//! MCP process spawner and stdio bridge.
//!
//! Manages the lifecycle of MCP server subprocesses and bridges
//! communication between SSE clients and the MCP stdio transport.

use std::process::Stdio;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;

use serde_json::{json, Value};
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
use tokio::process::{Child, Command};
use tokio::sync::mpsc;
use tracing::{debug, error, info, warn};

use crate::error::{ProxyError, Result};

/// Counter for generating unique session IDs.
static SESSION_COUNTER: AtomicU64 = AtomicU64::new(0);

/// An active MCP session with a subprocess.
pub struct McpSession {
    /// Unique session identifier.
    pub id: String,
    /// Tenant ID for this session (used for logging/debugging).
    #[allow(dead_code)]
    pub tenant_id: String,
    /// Channel to send requests to the MCP process.
    request_tx: mpsc::Sender<Value>,
    /// Handle to the child process.
    child: Option<Child>,
}

impl McpSession {
    /// Spawn a new MCP process and create a session.
    pub async fn spawn(
        binary_path: &str,
        tenant_id: String,
        storage_grpc_url: Option<&str>,
    ) -> Result<(Self, mpsc::Receiver<Value>)> {
        let session_id = format!(
            "sse-{}",
            SESSION_COUNTER.fetch_add(1, Ordering::Relaxed)
        );

        info!(
            "Spawning MCP process for session {} (tenant: {})",
            session_id,
            if tenant_id.is_empty() {
                "<default>"
            } else {
                &tenant_id
            }
        );

        // Build command with environment
        let mut cmd = Command::new(binary_path);
        cmd.stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit()); // MCP logs go to stderr

        // Pass tenant ID via environment
        if !tenant_id.is_empty() {
            cmd.env("DOCX_MCP_TENANT_ID", &tenant_id);
        }

        // Pass gRPC storage URL if configured
        if let Some(url) = storage_grpc_url {
            cmd.env("STORAGE_GRPC_URL", url);
        }

        let mut child = cmd
            .spawn()
            .map_err(|e| ProxyError::McpSpawnError(e.to_string()))?;

        let stdin = child
            .stdin
            .take()
            .ok_or_else(|| ProxyError::McpSpawnError("Failed to get stdin".to_string()))?;

        let stdout = child
            .stdout
            .take()
            .ok_or_else(|| ProxyError::McpSpawnError("Failed to get stdout".to_string()))?;

        // Create channels
        let (request_tx, mut request_rx) = mpsc::channel::<Value>(32);
        let (response_tx, response_rx) = mpsc::channel::<Value>(32);

        // Spawn stdin writer task
        let session_id_clone = session_id.clone();
        let tenant_id_clone = tenant_id.clone();
        tokio::spawn(async move {
            let mut stdin = stdin;
            while let Some(mut request) = request_rx.recv().await {
                // Inject tenant_id into params if present
                if let Some(params) = request.get_mut("params") {
                    if let Some(obj) = params.as_object_mut() {
                        if !tenant_id_clone.is_empty() {
                            obj.insert("tenant_id".to_string(), json!(tenant_id_clone));
                        }
                    }
                }

                let line = match serde_json::to_string(&request) {
                    Ok(s) => s,
                    Err(e) => {
                        error!("Failed to serialize request: {}", e);
                        continue;
                    }
                };

                debug!("[{}] -> MCP: {}", session_id_clone, &line[..line.len().min(200)]);

                if let Err(e) = stdin.write_all(line.as_bytes()).await {
                    error!("Failed to write to MCP stdin: {}", e);
                    break;
                }
                if let Err(e) = stdin.write_all(b"\n").await {
                    error!("Failed to write newline to MCP stdin: {}", e);
                    break;
                }
                if let Err(e) = stdin.flush().await {
                    error!("Failed to flush MCP stdin: {}", e);
                    break;
                }
            }
            debug!("[{}] stdin writer task ended", session_id_clone);
        });

        // Spawn stdout reader task
        let session_id_clone = session_id.clone();
        tokio::spawn(async move {
            let reader = BufReader::new(stdout);
            let mut lines = reader.lines();

            while let Ok(Some(line)) = lines.next_line().await {
                debug!("[{}] <- MCP: {}", session_id_clone, &line[..line.len().min(200)]);

                match serde_json::from_str::<Value>(&line) {
                    Ok(response) => {
                        if response_tx.send(response).await.is_err() {
                            debug!("[{}] Response receiver dropped", session_id_clone);
                            break;
                        }
                    }
                    Err(e) => {
                        warn!("[{}] Failed to parse MCP response: {}", session_id_clone, e);
                    }
                }
            }
            debug!("[{}] stdout reader task ended", session_id_clone);
        });

        let session = McpSession {
            id: session_id,
            tenant_id,
            request_tx,
            child: Some(child),
        };

        Ok((session, response_rx))
    }

    /// Send a request to the MCP process.
    pub async fn send(&self, request: Value) -> Result<()> {
        self.request_tx
            .send(request)
            .await
            .map_err(|e| ProxyError::McpProcessError(format!("Failed to send request: {}", e)))
    }

    /// Gracefully shut down the MCP process.
    pub async fn shutdown(&mut self) {
        if let Some(mut child) = self.child.take() {
            info!("[{}] Shutting down MCP process", self.id);

            // Drop the request channel to signal the stdin writer to stop
            drop(self.request_tx.clone());

            // Give the process a moment to exit gracefully
            tokio::select! {
                result = child.wait() => {
                    match result {
                        Ok(status) => info!("[{}] MCP process exited with {}", self.id, status),
                        Err(e) => warn!("[{}] Failed to wait for MCP process: {}", self.id, e),
                    }
                }
                _ = tokio::time::sleep(std::time::Duration::from_secs(5)) => {
                    warn!("[{}] MCP process did not exit in time, killing", self.id);
                    if let Err(e) = child.kill().await {
                        error!("[{}] Failed to kill MCP process: {}", self.id, e);
                    }
                }
            }
        }
    }
}

impl Drop for McpSession {
    fn drop(&mut self) {
        if self.child.is_some() {
            warn!("[{}] McpSession dropped without shutdown", self.id);
        }
    }
}

/// Manages multiple MCP sessions.
pub struct McpSessionManager {
    binary_path: String,
    storage_grpc_url: Option<String>,
}

impl McpSessionManager {
    /// Create a new session manager.
    pub fn new(binary_path: String, storage_grpc_url: Option<String>) -> Self {
        Self {
            binary_path,
            storage_grpc_url,
        }
    }

    /// Spawn a new MCP session for a tenant.
    pub async fn spawn_session(
        &self,
        tenant_id: String,
    ) -> Result<(McpSession, mpsc::Receiver<Value>)> {
        McpSession::spawn(
            &self.binary_path,
            tenant_id,
            self.storage_grpc_url.as_deref(),
        )
        .await
    }
}

/// Shared session manager.
pub type SharedMcpSessionManager = Arc<McpSessionManager>;

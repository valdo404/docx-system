//! HTTP handlers for the SSE proxy.
//!
//! Implements:
//! - POST /mcp - Streamable HTTP MCP endpoint with SSE responses
//! - GET /health - Health check endpoint

use std::convert::Infallible;
use std::time::Duration;

use axum::extract::{Request, State};
use axum::http::header;
use axum::response::sse::{Event, Sse};
use axum::response::{IntoResponse, Response};
use axum::Json;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use tokio_stream::wrappers::ReceiverStream;
use tokio_stream::StreamExt;
use tracing::{debug, info};

use crate::auth::SharedPatValidator;
use crate::error::ProxyError;
use crate::mcp::SharedMcpSessionManager;

/// Application state shared across handlers.
#[derive(Clone)]
pub struct AppState {
    pub validator: Option<SharedPatValidator>,
    pub session_manager: SharedMcpSessionManager,
}

/// Health check response.
#[derive(Serialize)]
pub struct HealthResponse {
    pub healthy: bool,
    pub version: &'static str,
    pub auth_enabled: bool,
}

/// GET /health - Health check endpoint.
pub async fn health_handler(State(state): State<AppState>) -> Json<HealthResponse> {
    Json(HealthResponse {
        healthy: true,
        version: env!("CARGO_PKG_VERSION"),
        auth_enabled: state.validator.is_some(),
    })
}

/// Extract Bearer token from Authorization header.
fn extract_bearer_token(req: &Request) -> Option<&str> {
    req.headers()
        .get(header::AUTHORIZATION)
        .and_then(|v| v.to_str().ok())
        .and_then(|v| v.strip_prefix("Bearer "))
}

/// MCP JSON-RPC request structure.
#[derive(Deserialize)]
struct McpRequest {
    jsonrpc: String,
    method: String,
    params: Option<Value>,
    id: Option<Value>,
}

/// POST /mcp - Streamable HTTP MCP endpoint.
///
/// This implements the MCP Streamable HTTP transport:
/// - Accepts JSON-RPC requests in the body
/// - Returns SSE stream of responses
/// - Injects tenant_id into request params based on authenticated PAT
pub async fn mcp_handler(
    State(state): State<AppState>,
    req: Request,
) -> std::result::Result<Response, ProxyError> {
    // Authenticate if validator is configured
    let tenant_id = if let Some(ref validator) = state.validator {
        let token = extract_bearer_token(&req).ok_or(ProxyError::Unauthorized)?;

        let validation = validator.validate(token).await?;
        info!(
            "Authenticated request for tenant {} (PAT: {}...)",
            validation.tenant_id,
            &validation.pat_id[..8.min(validation.pat_id.len())]
        );
        validation.tenant_id
    } else {
        // No auth configured - use empty tenant (local mode)
        debug!("Auth not configured, using default tenant");
        String::new()
    };

    // Parse request body
    let body = axum::body::to_bytes(req.into_body(), 1024 * 1024) // 1MB limit
        .await
        .map_err(|e| ProxyError::Internal(format!("Failed to read body: {}", e)))?;

    let mcp_request: McpRequest = serde_json::from_slice(&body)?;

    debug!(
        "MCP request: method={}, id={:?}",
        mcp_request.method, mcp_request.id
    );

    // Spawn MCP session
    let (mut session, response_rx) = state.session_manager.spawn_session(tenant_id).await?;

    // Build the JSON-RPC request to forward
    let mut forward_request = json!({
        "jsonrpc": mcp_request.jsonrpc,
        "method": mcp_request.method,
    });

    if let Some(params) = mcp_request.params {
        forward_request["params"] = params;
    }
    if let Some(id) = mcp_request.id.clone() {
        forward_request["id"] = id;
    }

    // Send request to MCP process
    session.send(forward_request).await?;

    // Create SSE stream from response channel
    let session_id = session.id.clone();

    let stream = ReceiverStream::new(response_rx).map(move |response| {
        let event_data = serde_json::to_string(&response).unwrap_or_else(|e| {
            json!({
                "jsonrpc": "2.0",
                "error": {
                    "code": -32603,
                    "message": format!("Failed to serialize response: {}", e)
                }
            })
            .to_string()
        });

        Ok::<_, Infallible>(Event::default().data(event_data))
    });

    // Spawn cleanup task
    let session_id_clone = session_id.clone();
    tokio::spawn(async move {
        // Wait a bit for the stream to complete, then clean up
        tokio::time::sleep(Duration::from_secs(60)).await;
        session.shutdown().await;
        debug!("[{}] Session cleaned up", session_id_clone);
    });

    Ok(Sse::new(stream)
        .keep_alive(
            axum::response::sse::KeepAlive::new()
                .interval(Duration::from_secs(15))
                .text("keep-alive"),
        )
        .into_response())
}

/// POST /mcp/message - Simpler request/response endpoint (non-streaming).
///
/// For clients that don't need SSE, this provides a simple JSON request/response.
pub async fn mcp_message_handler(
    State(state): State<AppState>,
    req: Request,
) -> std::result::Result<Response, ProxyError> {
    // Authenticate if validator is configured
    let tenant_id = if let Some(ref validator) = state.validator {
        let token = extract_bearer_token(&req).ok_or(ProxyError::Unauthorized)?;
        validator.validate(token).await?.tenant_id
    } else {
        String::new()
    };

    // Parse request body
    let body = axum::body::to_bytes(req.into_body(), 1024 * 1024)
        .await
        .map_err(|e| ProxyError::Internal(format!("Failed to read body: {}", e)))?;

    let mcp_request: McpRequest = serde_json::from_slice(&body)?;
    let request_id = mcp_request.id.clone();

    // Spawn MCP session
    let (mut session, mut response_rx) = state.session_manager.spawn_session(tenant_id).await?;

    // Build and send request
    let mut forward_request = json!({
        "jsonrpc": mcp_request.jsonrpc,
        "method": mcp_request.method,
    });

    if let Some(params) = mcp_request.params {
        forward_request["params"] = params;
    }
    if let Some(id) = mcp_request.id {
        forward_request["id"] = id;
    }

    session.send(forward_request).await?;

    // Wait for response with timeout
    let response = tokio::time::timeout(Duration::from_secs(30), async {
        while let Some(response) = response_rx.recv().await {
            // Return when we get a response (has result or error)
            if response.get("result").is_some() || response.get("error").is_some() {
                // Check ID matches if we have one
                if let Some(ref req_id) = request_id {
                    if response.get("id") == Some(req_id) {
                        return Some(response);
                    }
                } else {
                    return Some(response);
                }
            }
        }
        None
    })
    .await
    .map_err(|_| ProxyError::McpProcessError("Request timed out".to_string()))?
    .ok_or_else(|| ProxyError::McpProcessError("No response from MCP process".to_string()))?;

    // Clean up
    session.shutdown().await;

    Ok(Json(response).into_response())
}

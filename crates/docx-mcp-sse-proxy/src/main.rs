//! SSE/HTTP proxy for docx-mcp multi-tenant architecture.
//!
//! This proxy:
//! - Receives HTTP Streamable MCP requests
//! - Validates PAT tokens via Cloudflare D1
//! - Extracts tenant_id from validated tokens
//! - Forwards requests to MCP .NET process via stdio
//! - Streams responses back to clients via SSE

use std::sync::Arc;

use axum::routing::{get, post};
use axum::Router;
use clap::Parser;
use tokio::net::TcpListener;
use tokio::signal;
use tower_http::cors::{Any, CorsLayer};
use tower_http::trace::TraceLayer;
use tracing::{info, warn};
use tracing_subscriber::EnvFilter;

mod auth;
mod config;
mod error;
mod handlers;
mod mcp;

use auth::{PatValidator, SharedPatValidator};
use config::Config;
use handlers::{health_handler, mcp_handler, mcp_message_handler, AppState};
use mcp::{McpSessionManager, SharedMcpSessionManager};

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // Initialize logging
    tracing_subscriber::fmt()
        .with_env_filter(
            EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("info")),
        )
        .init();

    let config = Config::parse();

    info!("Starting docx-mcp-sse-proxy v{}", env!("CARGO_PKG_VERSION"));
    info!("  Host: {}", config.host);
    info!("  Port: {}", config.port);

    // Create PAT validator if D1 credentials are configured
    let validator: Option<SharedPatValidator> = if config.cloudflare_account_id.is_some()
        && config.cloudflare_api_token.is_some()
        && config.d1_database_id.is_some()
    {
        info!("  Auth: D1 PAT validation enabled");
        info!(
            "  PAT cache TTL: {}s (negative: {}s)",
            config.pat_cache_ttl_secs, config.pat_negative_cache_ttl_secs
        );

        Some(Arc::new(PatValidator::new(
            config.cloudflare_account_id.clone().unwrap(),
            config.cloudflare_api_token.clone().unwrap(),
            config.d1_database_id.clone().unwrap(),
            config.pat_cache_ttl_secs,
            config.pat_negative_cache_ttl_secs,
        )))
    } else {
        warn!("  Auth: DISABLED (no D1 credentials configured)");
        warn!("  Set CLOUDFLARE_ACCOUNT_ID, CLOUDFLARE_API_TOKEN, and D1_DATABASE_ID to enable auth");
        None
    };

    // Determine MCP binary path
    let binary_path = config.docx_mcp_binary.clone().unwrap_or_else(|| {
        // Try to find the binary in common locations
        let candidates = [
            "docx-mcp",
            "./docx-mcp",
            "../dist/docx-mcp",
            "/usr/local/bin/docx-mcp",
        ];

        for candidate in candidates {
            if std::path::Path::new(candidate).exists() {
                return candidate.to_string();
            }
        }

        // Default to PATH lookup
        "docx-mcp".to_string()
    });

    info!("  MCP binary: {}", binary_path);

    if let Some(ref url) = config.storage_grpc_url {
        info!("  Storage gRPC: {}", url);
    }

    // Create session manager
    let session_manager: SharedMcpSessionManager = Arc::new(McpSessionManager::new(
        binary_path,
        config.storage_grpc_url.clone(),
    ));

    // Build application state
    let state = AppState {
        validator,
        session_manager,
    };

    // Configure CORS
    let cors = CorsLayer::new()
        .allow_origin(Any)
        .allow_methods(Any)
        .allow_headers(Any);

    // Build router
    let app = Router::new()
        .route("/health", get(health_handler))
        .route("/mcp", post(mcp_handler))
        .route("/mcp/message", post(mcp_message_handler))
        .layer(cors)
        .layer(TraceLayer::new_for_http())
        .with_state(state);

    // Bind and serve
    let addr = format!("{}:{}", config.host, config.port);
    let listener = TcpListener::bind(&addr).await?;
    info!("Listening on http://{}", addr);

    axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal())
        .await?;

    info!("Server shutdown complete");
    Ok(())
}

async fn shutdown_signal() {
    let ctrl_c = async {
        signal::ctrl_c()
            .await
            .expect("Failed to install Ctrl+C handler");
        info!("Received Ctrl+C, initiating shutdown");
    };

    #[cfg(unix)]
    let terminate = async {
        signal::unix::signal(signal::unix::SignalKind::terminate())
            .expect("Failed to install SIGTERM handler")
            .recv()
            .await;
        info!("Received SIGTERM, initiating shutdown");
    };

    #[cfg(not(unix))]
    let terminate = std::future::pending::<()>();

    tokio::select! {
        _ = ctrl_c => {},
        _ = terminate => {},
    }
}

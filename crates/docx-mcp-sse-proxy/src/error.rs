//! Error types for the SSE proxy.

use axum::http::StatusCode;
use axum::response::{IntoResponse, Response};
use serde::Serialize;

/// Application-level errors.
#[derive(Debug, thiserror::Error)]
pub enum ProxyError {
    #[error("Authentication required")]
    Unauthorized,

    #[error("Invalid or expired PAT token")]
    InvalidToken,

    #[error("D1 API error: {0}")]
    D1Error(String),

    #[error("Failed to spawn MCP process: {0}")]
    McpSpawnError(String),

    #[error("MCP process error: {0}")]
    McpProcessError(String),

    #[error("Invalid JSON: {0}")]
    JsonError(#[from] serde_json::Error),

    #[error("Internal error: {0}")]
    Internal(String),
}

impl IntoResponse for ProxyError {
    fn into_response(self) -> Response {
        #[derive(Serialize)]
        struct ErrorBody {
            error: String,
            code: &'static str,
        }

        let (status, code) = match &self {
            ProxyError::Unauthorized => (StatusCode::UNAUTHORIZED, "UNAUTHORIZED"),
            ProxyError::InvalidToken => (StatusCode::UNAUTHORIZED, "INVALID_TOKEN"),
            ProxyError::D1Error(_) => (StatusCode::BAD_GATEWAY, "D1_ERROR"),
            ProxyError::McpSpawnError(_) => (StatusCode::INTERNAL_SERVER_ERROR, "MCP_SPAWN_ERROR"),
            ProxyError::McpProcessError(_) => {
                (StatusCode::INTERNAL_SERVER_ERROR, "MCP_PROCESS_ERROR")
            }
            ProxyError::JsonError(_) => (StatusCode::BAD_REQUEST, "INVALID_JSON"),
            ProxyError::Internal(_) => (StatusCode::INTERNAL_SERVER_ERROR, "INTERNAL_ERROR"),
        };

        let body = ErrorBody {
            error: self.to_string(),
            code,
        };

        (status, axum::Json(body)).into_response()
    }
}

pub type Result<T> = std::result::Result<T, ProxyError>;

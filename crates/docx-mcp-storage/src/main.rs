mod config;
mod error;
mod lock;
mod service;
mod storage;

use std::sync::Arc;

use clap::Parser;
use tokio::net::UnixListener;
use tokio::signal;
use tonic::transport::Server;
use tracing::info;
use tracing_subscriber::EnvFilter;

use config::{Config, StorageBackend, Transport};
use lock::FileLock;
use service::proto::storage_service_server::StorageServiceServer;
use service::StorageServiceImpl;
use storage::LocalStorage;

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // Initialize logging
    tracing_subscriber::fmt()
        .with_env_filter(
            EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("info")),
        )
        .init();

    let config = Config::parse();

    info!("Starting docx-mcp-storage server");
    info!("  Transport: {}", config.transport);
    info!("  Backend: {}", config.storage_backend);

    // Create storage backend
    let storage: Arc<dyn crate::storage::StorageBackend> = match config.storage_backend {
        StorageBackend::Local => {
            let dir = config.effective_local_storage_dir();
            info!("  Local storage dir: {}", dir.display());
            Arc::new(LocalStorage::new(&dir))
        }
        #[cfg(feature = "cloud")]
        StorageBackend::R2 => {
            todo!("R2 storage backend not yet implemented")
        }
    };

    // Create lock manager (using same base dir as storage for local)
    let lock_manager: Arc<dyn crate::lock::LockManager> = match config.storage_backend {
        StorageBackend::Local => {
            let dir = config.effective_local_storage_dir();
            Arc::new(FileLock::new(&dir))
        }
        #[cfg(feature = "cloud")]
        StorageBackend::R2 => {
            todo!("KV lock manager not yet implemented")
        }
    };

    // Create gRPC service
    let service = StorageServiceImpl::new(storage, lock_manager);
    let svc = StorageServiceServer::new(service);

    // Start server based on transport
    match config.transport {
        Transport::Tcp => {
            let addr = format!("{}:{}", config.host, config.port).parse()?;
            info!("Listening on tcp://{}", addr);

            Server::builder()
                .add_service(svc)
                .serve_with_shutdown(addr, shutdown_signal())
                .await?;
        }
        Transport::Unix => {
            let socket_path = config.effective_unix_socket();

            // Remove existing socket file if it exists
            if socket_path.exists() {
                std::fs::remove_file(&socket_path)?;
            }

            // Ensure parent directory exists
            if let Some(parent) = socket_path.parent() {
                std::fs::create_dir_all(parent)?;
            }

            info!("Listening on unix://{}", socket_path.display());

            let uds = UnixListener::bind(&socket_path)?;
            let uds_stream = tokio_stream::wrappers::UnixListenerStream::new(uds);

            Server::builder()
                .add_service(svc)
                .serve_with_incoming_shutdown(uds_stream, shutdown_signal())
                .await?;

            // Clean up socket on shutdown
            if socket_path.exists() {
                let _ = std::fs::remove_file(&socket_path);
            }
        }
    }

    info!("Server shutdown complete");
    Ok(())
}

async fn shutdown_signal() {
    let ctrl_c = async {
        signal::ctrl_c()
            .await
            .expect("Failed to install Ctrl+C handler");
    };

    #[cfg(unix)]
    let terminate = async {
        signal::unix::signal(signal::unix::SignalKind::terminate())
            .expect("Failed to install SIGTERM handler")
            .recv()
            .await;
    };

    #[cfg(not(unix))]
    let terminate = std::future::pending::<()>();

    tokio::select! {
        _ = ctrl_c => {
            info!("Received Ctrl+C, initiating shutdown");
        },
        _ = terminate => {
            info!("Received SIGTERM, initiating shutdown");
        },
    }
}

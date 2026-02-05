use std::path::PathBuf;

use clap::Parser;

/// Configuration for the docx-mcp-storage server.
#[derive(Parser, Debug, Clone)]
#[command(name = "docx-mcp-storage")]
#[command(about = "gRPC storage server for docx-mcp multi-tenant architecture")]
pub struct Config {
    /// Transport type: tcp or unix
    #[arg(long, default_value = "tcp", env = "GRPC_TRANSPORT")]
    pub transport: Transport,

    /// TCP host to bind to (only used with --transport tcp)
    #[arg(long, default_value = "0.0.0.0", env = "GRPC_HOST")]
    pub host: String,

    /// TCP port to bind to (only used with --transport tcp)
    #[arg(long, default_value = "50051", env = "GRPC_PORT")]
    pub port: u16,

    /// Unix socket path (only used with --transport unix)
    #[arg(long, env = "GRPC_UNIX_SOCKET")]
    pub unix_socket: Option<PathBuf>,

    /// Storage backend: local or r2
    #[arg(long, default_value = "local", env = "STORAGE_BACKEND")]
    pub storage_backend: StorageBackend,

    /// Base directory for local storage
    #[arg(long, env = "LOCAL_STORAGE_DIR")]
    pub local_storage_dir: Option<PathBuf>,

    /// R2 endpoint URL (for r2 backend)
    #[arg(long, env = "R2_ENDPOINT")]
    pub r2_endpoint: Option<String>,

    /// R2 access key ID
    #[arg(long, env = "R2_ACCESS_KEY_ID")]
    pub r2_access_key_id: Option<String>,

    /// R2 secret access key
    #[arg(long, env = "R2_SECRET_ACCESS_KEY")]
    pub r2_secret_access_key: Option<String>,

    /// R2 bucket name
    #[arg(long, env = "R2_BUCKET_NAME")]
    pub r2_bucket_name: Option<String>,
}

impl Config {
    /// Get the effective local storage directory.
    pub fn effective_local_storage_dir(&self) -> PathBuf {
        self.local_storage_dir.clone().unwrap_or_else(|| {
            dirs::data_local_dir()
                .unwrap_or_else(|| PathBuf::from("."))
                .join("docx-mcp")
                .join("sessions")
        })
    }

    /// Get the effective Unix socket path.
    pub fn effective_unix_socket(&self) -> PathBuf {
        self.unix_socket.clone().unwrap_or_else(|| {
            std::env::var("XDG_RUNTIME_DIR")
                .map(PathBuf::from)
                .unwrap_or_else(|_| PathBuf::from("/tmp"))
                .join("docx-mcp-storage.sock")
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, clap::ValueEnum)]
pub enum Transport {
    Tcp,
    Unix,
}

impl std::fmt::Display for Transport {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Transport::Tcp => write!(f, "tcp"),
            Transport::Unix => write!(f, "unix"),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, clap::ValueEnum)]
pub enum StorageBackend {
    Local,
    #[cfg(feature = "cloud")]
    R2,
}

impl std::fmt::Display for StorageBackend {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            StorageBackend::Local => write!(f, "local"),
            #[cfg(feature = "cloud")]
            StorageBackend::R2 => write!(f, "r2"),
        }
    }
}

use async_trait::async_trait;
use serde::{Deserialize, Serialize};

use crate::error::StorageError;

/// Information about a session stored in the backend.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionInfo {
    pub session_id: String,
    pub source_path: Option<String>,
    pub created_at: chrono::DateTime<chrono::Utc>,
    pub modified_at: chrono::DateTime<chrono::Utc>,
    pub size_bytes: u64,
}

/// A single WAL entry representing an edit operation.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WalEntry {
    pub position: u64,
    pub operation: String,
    pub path: String,
    pub patch_json: Vec<u8>,
    pub timestamp: chrono::DateTime<chrono::Utc>,
}

/// Information about a checkpoint.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CheckpointInfo {
    pub position: u64,
    pub created_at: chrono::DateTime<chrono::Utc>,
    pub size_bytes: u64,
}

/// The session index containing metadata about all sessions for a tenant.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SessionIndex {
    pub sessions: std::collections::HashMap<String, SessionIndexEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionIndexEntry {
    pub source_path: Option<String>,
    pub created_at: chrono::DateTime<chrono::Utc>,
    pub modified_at: chrono::DateTime<chrono::Utc>,
    pub wal_position: u64,
    pub checkpoint_positions: Vec<u64>,
}

/// Storage backend abstraction for tenant-aware document storage.
///
/// All methods take `tenant_id` as the first parameter to ensure isolation.
/// Implementations must organize data by tenant (e.g., `{base}/{tenant_id}/`).
#[async_trait]
pub trait StorageBackend: Send + Sync {
    /// Returns the backend identifier (e.g., "local", "r2").
    fn backend_name(&self) -> &'static str;

    // =========================================================================
    // Session Operations
    // =========================================================================

    /// Load a session's DOCX bytes.
    async fn load_session(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Option<Vec<u8>>, StorageError>;

    /// Save a session's DOCX bytes.
    async fn save_session(
        &self,
        tenant_id: &str,
        session_id: &str,
        data: &[u8],
    ) -> Result<(), StorageError>;

    /// Delete a session and all associated data (WAL, checkpoints).
    async fn delete_session(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<bool, StorageError>;

    /// List all sessions for a tenant.
    async fn list_sessions(&self, tenant_id: &str) -> Result<Vec<SessionInfo>, StorageError>;

    /// Check if a session exists.
    async fn session_exists(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<bool, StorageError>;

    // =========================================================================
    // Index Operations
    // =========================================================================

    /// Load the session index for a tenant.
    async fn load_index(&self, tenant_id: &str) -> Result<Option<SessionIndex>, StorageError>;

    /// Save the session index for a tenant.
    async fn save_index(
        &self,
        tenant_id: &str,
        index: &SessionIndex,
    ) -> Result<(), StorageError>;

    // =========================================================================
    // WAL Operations
    // =========================================================================

    /// Append entries to a session's WAL.
    async fn append_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        entries: &[WalEntry],
    ) -> Result<u64, StorageError>;

    /// Read WAL entries starting from a position.
    async fn read_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        from_position: u64,
        limit: Option<u64>,
    ) -> Result<(Vec<WalEntry>, bool), StorageError>;

    /// Truncate WAL, keeping only entries at or after the given position.
    async fn truncate_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        keep_from: u64,
    ) -> Result<u64, StorageError>;

    // =========================================================================
    // Checkpoint Operations
    // =========================================================================

    /// Save a checkpoint at a specific WAL position.
    async fn save_checkpoint(
        &self,
        tenant_id: &str,
        session_id: &str,
        position: u64,
        data: &[u8],
    ) -> Result<(), StorageError>;

    /// Load a checkpoint. If position is 0, load the latest.
    async fn load_checkpoint(
        &self,
        tenant_id: &str,
        session_id: &str,
        position: u64,
    ) -> Result<Option<(Vec<u8>, u64)>, StorageError>;

    /// List all checkpoints for a session.
    async fn list_checkpoints(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Vec<CheckpointInfo>, StorageError>;
}

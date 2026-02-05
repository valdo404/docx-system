use std::path::{Path, PathBuf};

use async_trait::async_trait;
use tokio::fs;
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
use tracing::{debug, instrument, warn};

use super::traits::{
    CheckpointInfo, SessionIndex, SessionInfo, StorageBackend, WalEntry,
};
use crate::error::StorageError;

/// Local filesystem storage backend.
///
/// Organizes data by tenant:
/// ```
/// {base_dir}/
///   {tenant_id}/
///     sessions/
///       index.json
///       {session_id}.docx
///       {session_id}.wal
///       {session_id}.ckpt.{position}.docx
/// ```
#[derive(Debug, Clone)]
pub struct LocalStorage {
    base_dir: PathBuf,
}

impl LocalStorage {
    /// Create a new LocalStorage with the given base directory.
    pub fn new(base_dir: impl AsRef<Path>) -> Self {
        Self {
            base_dir: base_dir.as_ref().to_path_buf(),
        }
    }

    /// Get the sessions directory for a tenant.
    fn sessions_dir(&self, tenant_id: &str) -> PathBuf {
        self.base_dir.join(tenant_id).join("sessions")
    }

    /// Get the path to a session file.
    fn session_path(&self, tenant_id: &str, session_id: &str) -> PathBuf {
        self.sessions_dir(tenant_id)
            .join(format!("{}.docx", session_id))
    }

    /// Get the path to a session's WAL file.
    fn wal_path(&self, tenant_id: &str, session_id: &str) -> PathBuf {
        self.sessions_dir(tenant_id)
            .join(format!("{}.wal", session_id))
    }

    /// Get the path to a checkpoint file.
    fn checkpoint_path(&self, tenant_id: &str, session_id: &str, position: u64) -> PathBuf {
        self.sessions_dir(tenant_id)
            .join(format!("{}.ckpt.{}.docx", session_id, position))
    }

    /// Get the path to the index file.
    fn index_path(&self, tenant_id: &str) -> PathBuf {
        self.sessions_dir(tenant_id).join("index.json")
    }

    /// Ensure the sessions directory exists.
    async fn ensure_sessions_dir(&self, tenant_id: &str) -> Result<(), StorageError> {
        let dir = self.sessions_dir(tenant_id);
        fs::create_dir_all(&dir).await.map_err(|e| {
            StorageError::Io(format!("Failed to create sessions dir {}: {}", dir.display(), e))
        })?;
        Ok(())
    }
}

#[async_trait]
impl StorageBackend for LocalStorage {
    fn backend_name(&self) -> &'static str {
        "local"
    }

    // =========================================================================
    // Session Operations
    // =========================================================================

    #[instrument(skip(self), level = "debug")]
    async fn load_session(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Option<Vec<u8>>, StorageError> {
        let path = self.session_path(tenant_id, session_id);
        match fs::read(&path).await {
            Ok(data) => {
                debug!("Loaded session {} ({} bytes)", session_id, data.len());
                Ok(Some(data))
            }
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
            Err(e) => Err(StorageError::Io(format!(
                "Failed to read {}: {}",
                path.display(),
                e
            ))),
        }
    }

    #[instrument(skip(self, data), level = "debug", fields(data_len = data.len()))]
    async fn save_session(
        &self,
        tenant_id: &str,
        session_id: &str,
        data: &[u8],
    ) -> Result<(), StorageError> {
        self.ensure_sessions_dir(tenant_id).await?;
        let path = self.session_path(tenant_id, session_id);

        // Write atomically via temp file
        let temp_path = path.with_extension("docx.tmp");
        fs::write(&temp_path, data).await.map_err(|e| {
            StorageError::Io(format!("Failed to write {}: {}", temp_path.display(), e))
        })?;
        fs::rename(&temp_path, &path).await.map_err(|e| {
            StorageError::Io(format!("Failed to rename to {}: {}", path.display(), e))
        })?;

        debug!("Saved session {} ({} bytes)", session_id, data.len());
        Ok(())
    }

    #[instrument(skip(self), level = "debug")]
    async fn delete_session(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<bool, StorageError> {
        let session_path = self.session_path(tenant_id, session_id);
        let wal_path = self.wal_path(tenant_id, session_id);

        let existed = session_path.exists();

        // Delete session file
        if let Err(e) = fs::remove_file(&session_path).await {
            if e.kind() != std::io::ErrorKind::NotFound {
                warn!("Failed to delete session file: {}", e);
            }
        }

        // Delete WAL
        if let Err(e) = fs::remove_file(&wal_path).await {
            if e.kind() != std::io::ErrorKind::NotFound {
                warn!("Failed to delete WAL file: {}", e);
            }
        }

        // Delete all checkpoints
        let checkpoints = self.list_checkpoints(tenant_id, session_id).await?;
        for ckpt in checkpoints {
            let ckpt_path = self.checkpoint_path(tenant_id, session_id, ckpt.position);
            if let Err(e) = fs::remove_file(&ckpt_path).await {
                if e.kind() != std::io::ErrorKind::NotFound {
                    warn!("Failed to delete checkpoint: {}", e);
                }
            }
        }

        debug!("Deleted session {} (existed: {})", session_id, existed);
        Ok(existed)
    }

    #[instrument(skip(self), level = "debug")]
    async fn list_sessions(&self, tenant_id: &str) -> Result<Vec<SessionInfo>, StorageError> {
        let dir = self.sessions_dir(tenant_id);
        if !dir.exists() {
            return Ok(vec![]);
        }

        let mut sessions = Vec::new();
        let mut entries = fs::read_dir(&dir).await.map_err(|e| {
            StorageError::Io(format!("Failed to read dir {}: {}", dir.display(), e))
        })?;

        while let Some(entry) = entries.next_entry().await.map_err(|e| {
            StorageError::Io(format!("Failed to read dir entry: {}", e))
        })? {
            let path = entry.path();
            if path.extension().is_some_and(|ext| ext == "docx")
                && !path
                    .file_stem()
                    .is_some_and(|s| s.to_string_lossy().contains(".ckpt."))
            {
                let session_id = path
                    .file_stem()
                    .map(|s| s.to_string_lossy().to_string())
                    .unwrap_or_default();

                let metadata = entry.metadata().await.map_err(|e| {
                    StorageError::Io(format!("Failed to get metadata: {}", e))
                })?;

                let created_at = metadata
                    .created()
                    .map(chrono::DateTime::from)
                    .unwrap_or_else(|_| chrono::Utc::now());
                let modified_at = metadata
                    .modified()
                    .map(chrono::DateTime::from)
                    .unwrap_or_else(|_| chrono::Utc::now());

                sessions.push(SessionInfo {
                    session_id,
                    source_path: None, // Would need to read from index
                    created_at,
                    modified_at,
                    size_bytes: metadata.len(),
                });
            }
        }

        debug!("Listed {} sessions for tenant {}", sessions.len(), tenant_id);
        Ok(sessions)
    }

    #[instrument(skip(self), level = "debug")]
    async fn session_exists(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<bool, StorageError> {
        let path = self.session_path(tenant_id, session_id);
        Ok(path.exists())
    }

    // =========================================================================
    // Index Operations
    // =========================================================================

    #[instrument(skip(self), level = "debug")]
    async fn load_index(&self, tenant_id: &str) -> Result<Option<SessionIndex>, StorageError> {
        let path = self.index_path(tenant_id);
        match fs::read_to_string(&path).await {
            Ok(json) => {
                let index: SessionIndex = serde_json::from_str(&json).map_err(|e| {
                    StorageError::Serialization(format!("Failed to parse index: {}", e))
                })?;
                debug!("Loaded index with {} sessions", index.sessions.len());
                Ok(Some(index))
            }
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
            Err(e) => Err(StorageError::Io(format!(
                "Failed to read index {}: {}",
                path.display(),
                e
            ))),
        }
    }

    #[instrument(skip(self, index), level = "debug", fields(sessions = index.sessions.len()))]
    async fn save_index(
        &self,
        tenant_id: &str,
        index: &SessionIndex,
    ) -> Result<(), StorageError> {
        self.ensure_sessions_dir(tenant_id).await?;
        let path = self.index_path(tenant_id);

        let json = serde_json::to_string_pretty(index).map_err(|e| {
            StorageError::Serialization(format!("Failed to serialize index: {}", e))
        })?;

        // Write atomically
        let temp_path = path.with_extension("json.tmp");
        fs::write(&temp_path, &json).await.map_err(|e| {
            StorageError::Io(format!("Failed to write index: {}", e))
        })?;
        fs::rename(&temp_path, &path).await.map_err(|e| {
            StorageError::Io(format!("Failed to rename index: {}", e))
        })?;

        debug!("Saved index with {} sessions", index.sessions.len());
        Ok(())
    }

    // =========================================================================
    // WAL Operations
    // =========================================================================

    #[instrument(skip(self, entries), level = "debug", fields(entries_count = entries.len()))]
    async fn append_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        entries: &[WalEntry],
    ) -> Result<u64, StorageError> {
        if entries.is_empty() {
            return Ok(0);
        }

        self.ensure_sessions_dir(tenant_id).await?;
        let path = self.wal_path(tenant_id, session_id);

        let mut file = fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&path)
            .await
            .map_err(|e| StorageError::Io(format!("Failed to open WAL {}: {}", path.display(), e)))?;

        let mut last_position = 0u64;
        for entry in entries {
            let line = serde_json::to_string(entry).map_err(|e| {
                StorageError::Serialization(format!("Failed to serialize WAL entry: {}", e))
            })?;
            file.write_all(line.as_bytes()).await.map_err(|e| {
                StorageError::Io(format!("Failed to write WAL: {}", e))
            })?;
            file.write_all(b"\n").await.map_err(|e| {
                StorageError::Io(format!("Failed to write WAL newline: {}", e))
            })?;
            last_position = entry.position;
        }

        file.flush().await.map_err(|e| {
            StorageError::Io(format!("Failed to flush WAL: {}", e))
        })?;

        debug!(
            "Appended {} WAL entries, last position: {}",
            entries.len(),
            last_position
        );
        Ok(last_position)
    }

    #[instrument(skip(self), level = "debug")]
    async fn read_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        from_position: u64,
        limit: Option<u64>,
    ) -> Result<(Vec<WalEntry>, bool), StorageError> {
        let path = self.wal_path(tenant_id, session_id);

        let file = match fs::File::open(&path).await {
            Ok(f) => f,
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
                return Ok((vec![], false));
            }
            Err(e) => {
                return Err(StorageError::Io(format!(
                    "Failed to open WAL {}: {}",
                    path.display(),
                    e
                )));
            }
        };

        let reader = BufReader::new(file);
        let mut lines = reader.lines();
        let mut entries = Vec::new();
        let limit = limit.unwrap_or(u64::MAX);

        while let Some(line) = lines.next_line().await.map_err(|e| {
            StorageError::Io(format!("Failed to read WAL line: {}", e))
        })? {
            if line.trim().is_empty() {
                continue;
            }

            let entry: WalEntry = serde_json::from_str(&line).map_err(|e| {
                StorageError::Serialization(format!("Failed to parse WAL entry: {}", e))
            })?;

            if entry.position >= from_position {
                entries.push(entry);
                if entries.len() as u64 >= limit {
                    // Check if there are more
                    let has_more = lines.next_line().await.map_err(|e| {
                        StorageError::Io(format!("Failed to check for more WAL: {}", e))
                    })?.is_some();
                    return Ok((entries, has_more));
                }
            }
        }

        debug!(
            "Read {} WAL entries from position {}",
            entries.len(),
            from_position
        );
        Ok((entries, false))
    }

    #[instrument(skip(self), level = "debug")]
    async fn truncate_wal(
        &self,
        tenant_id: &str,
        session_id: &str,
        keep_from: u64,
    ) -> Result<u64, StorageError> {
        let (entries, _) = self.read_wal(tenant_id, session_id, 0, None).await?;

        let to_remove = entries.iter().filter(|e| e.position < keep_from).count() as u64;
        let to_keep: Vec<_> = entries
            .into_iter()
            .filter(|e| e.position >= keep_from)
            .collect();

        if to_remove == 0 {
            return Ok(0);
        }

        // Rewrite WAL with only kept entries
        let path = self.wal_path(tenant_id, session_id);
        let temp_path = path.with_extension("wal.tmp");

        let mut file = fs::File::create(&temp_path).await.map_err(|e| {
            StorageError::Io(format!("Failed to create temp WAL: {}", e))
        })?;

        for entry in &to_keep {
            let line = serde_json::to_string(entry).map_err(|e| {
                StorageError::Serialization(format!("Failed to serialize WAL entry: {}", e))
            })?;
            file.write_all(line.as_bytes()).await.map_err(|e| {
                StorageError::Io(format!("Failed to write WAL: {}", e))
            })?;
            file.write_all(b"\n").await.map_err(|e| {
                StorageError::Io(format!("Failed to write WAL newline: {}", e))
            })?;
        }

        file.flush().await.map_err(|e| {
            StorageError::Io(format!("Failed to flush temp WAL: {}", e))
        })?;

        fs::rename(&temp_path, &path).await.map_err(|e| {
            StorageError::Io(format!("Failed to rename temp WAL: {}", e))
        })?;

        debug!("Truncated WAL, removed {} entries", to_remove);
        Ok(to_remove)
    }

    // =========================================================================
    // Checkpoint Operations
    // =========================================================================

    #[instrument(skip(self, data), level = "debug", fields(data_len = data.len()))]
    async fn save_checkpoint(
        &self,
        tenant_id: &str,
        session_id: &str,
        position: u64,
        data: &[u8],
    ) -> Result<(), StorageError> {
        self.ensure_sessions_dir(tenant_id).await?;
        let path = self.checkpoint_path(tenant_id, session_id, position);

        // Write atomically
        let temp_path = path.with_extension("docx.tmp");
        fs::write(&temp_path, data).await.map_err(|e| {
            StorageError::Io(format!("Failed to write checkpoint: {}", e))
        })?;
        fs::rename(&temp_path, &path).await.map_err(|e| {
            StorageError::Io(format!("Failed to rename checkpoint: {}", e))
        })?;

        debug!(
            "Saved checkpoint at position {} ({} bytes)",
            position,
            data.len()
        );
        Ok(())
    }

    #[instrument(skip(self), level = "debug")]
    async fn load_checkpoint(
        &self,
        tenant_id: &str,
        session_id: &str,
        position: u64,
    ) -> Result<Option<(Vec<u8>, u64)>, StorageError> {
        if position == 0 {
            // Load latest checkpoint
            let checkpoints = self.list_checkpoints(tenant_id, session_id).await?;
            if let Some(latest) = checkpoints.last() {
                let path = self.checkpoint_path(tenant_id, session_id, latest.position);
                let data = fs::read(&path).await.map_err(|e| {
                    StorageError::Io(format!("Failed to read checkpoint: {}", e))
                })?;
                return Ok(Some((data, latest.position)));
            }
            return Ok(None);
        }

        let path = self.checkpoint_path(tenant_id, session_id, position);
        match fs::read(&path).await {
            Ok(data) => {
                debug!(
                    "Loaded checkpoint at position {} ({} bytes)",
                    position,
                    data.len()
                );
                Ok(Some((data, position)))
            }
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
            Err(e) => Err(StorageError::Io(format!(
                "Failed to read checkpoint: {}",
                e
            ))),
        }
    }

    #[instrument(skip(self), level = "debug")]
    async fn list_checkpoints(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Vec<CheckpointInfo>, StorageError> {
        let dir = self.sessions_dir(tenant_id);
        if !dir.exists() {
            return Ok(vec![]);
        }

        let prefix = format!("{}.ckpt.", session_id);
        let mut checkpoints = Vec::new();

        let mut entries = fs::read_dir(&dir).await.map_err(|e| {
            StorageError::Io(format!("Failed to read dir: {}", e))
        })?;

        while let Some(entry) = entries.next_entry().await.map_err(|e| {
            StorageError::Io(format!("Failed to read dir entry: {}", e))
        })? {
            let path = entry.path();
            let file_name = path
                .file_name()
                .map(|s| s.to_string_lossy().to_string())
                .unwrap_or_default();

            if file_name.starts_with(&prefix) && file_name.ends_with(".docx") {
                // Extract position from filename: {session_id}.ckpt.{position}.docx
                let position_str = file_name
                    .strip_prefix(&prefix)
                    .and_then(|s| s.strip_suffix(".docx"))
                    .unwrap_or("0");

                if let Ok(position) = position_str.parse::<u64>() {
                    let metadata = entry.metadata().await.map_err(|e| {
                        StorageError::Io(format!("Failed to get metadata: {}", e))
                    })?;

                    checkpoints.push(CheckpointInfo {
                        position,
                        created_at: metadata
                            .created()
                            .map(chrono::DateTime::from)
                            .unwrap_or_else(|_| chrono::Utc::now()),
                        size_bytes: metadata.len(),
                    });
                }
            }
        }

        // Sort by position
        checkpoints.sort_by_key(|c| c.position);

        debug!(
            "Listed {} checkpoints for session {}",
            checkpoints.len(),
            session_id
        );
        Ok(checkpoints)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    async fn setup() -> (LocalStorage, TempDir) {
        let temp_dir = TempDir::new().unwrap();
        let storage = LocalStorage::new(temp_dir.path());
        (storage, temp_dir)
    }

    #[tokio::test]
    async fn test_session_crud() {
        let (storage, _temp) = setup().await;
        let tenant = "test-tenant";
        let session = "test-session";
        let data = b"PK\x03\x04fake docx content";

        // Initially doesn't exist
        assert!(!storage.session_exists(tenant, session).await.unwrap());
        assert!(storage.load_session(tenant, session).await.unwrap().is_none());

        // Save
        storage.save_session(tenant, session, data).await.unwrap();

        // Now exists
        assert!(storage.session_exists(tenant, session).await.unwrap());

        // Load
        let loaded = storage.load_session(tenant, session).await.unwrap().unwrap();
        assert_eq!(loaded, data);

        // List
        let sessions = storage.list_sessions(tenant).await.unwrap();
        assert_eq!(sessions.len(), 1);
        assert_eq!(sessions[0].session_id, session);

        // Delete
        let existed = storage.delete_session(tenant, session).await.unwrap();
        assert!(existed);
        assert!(!storage.session_exists(tenant, session).await.unwrap());
    }

    #[tokio::test]
    async fn test_wal_operations() {
        let (storage, _temp) = setup().await;
        let tenant = "test-tenant";
        let session = "test-session";

        let entries = vec![
            WalEntry {
                position: 1,
                operation: "add".to_string(),
                path: "/body/paragraph[0]".to_string(),
                patch_json: b"{}".to_vec(),
                timestamp: chrono::Utc::now(),
            },
            WalEntry {
                position: 2,
                operation: "replace".to_string(),
                path: "/body/paragraph[0]/run[0]".to_string(),
                patch_json: b"{}".to_vec(),
                timestamp: chrono::Utc::now(),
            },
        ];

        // Append
        let last_pos = storage.append_wal(tenant, session, &entries).await.unwrap();
        assert_eq!(last_pos, 2);

        // Read all
        let (read_entries, has_more) = storage.read_wal(tenant, session, 0, None).await.unwrap();
        assert_eq!(read_entries.len(), 2);
        assert!(!has_more);

        // Read from position
        let (read_entries, _) = storage.read_wal(tenant, session, 2, None).await.unwrap();
        assert_eq!(read_entries.len(), 1);
        assert_eq!(read_entries[0].position, 2);

        // Truncate
        let removed = storage.truncate_wal(tenant, session, 2).await.unwrap();
        assert_eq!(removed, 1);

        let (read_entries, _) = storage.read_wal(tenant, session, 0, None).await.unwrap();
        assert_eq!(read_entries.len(), 1);
    }

    #[tokio::test]
    async fn test_checkpoint_operations() {
        let (storage, _temp) = setup().await;
        let tenant = "test-tenant";
        let session = "test-session";
        let data = b"checkpoint data";

        // Save checkpoints
        storage.save_checkpoint(tenant, session, 10, data).await.unwrap();
        storage.save_checkpoint(tenant, session, 20, data).await.unwrap();

        // List
        let checkpoints = storage.list_checkpoints(tenant, session).await.unwrap();
        assert_eq!(checkpoints.len(), 2);
        assert_eq!(checkpoints[0].position, 10);
        assert_eq!(checkpoints[1].position, 20);

        // Load specific
        let (loaded, pos) = storage.load_checkpoint(tenant, session, 10).await.unwrap().unwrap();
        assert_eq!(loaded, data);
        assert_eq!(pos, 10);

        // Load latest (position = 0)
        let (_, pos) = storage.load_checkpoint(tenant, session, 0).await.unwrap().unwrap();
        assert_eq!(pos, 20);
    }

    #[tokio::test]
    async fn test_tenant_isolation() {
        let (storage, _temp) = setup().await;
        let data = b"test data";

        // Save to tenant A
        storage.save_session("tenant-a", "session-1", data).await.unwrap();

        // Tenant B shouldn't see it
        assert!(!storage.session_exists("tenant-b", "session-1").await.unwrap());
        assert!(storage.list_sessions("tenant-b").await.unwrap().is_empty());
    }
}

use std::sync::Arc;

use async_trait::async_trait;
use aws_sdk_s3::primitives::ByteStream;
use aws_sdk_s3::Client as S3Client;
use docx_storage_core::{
    CheckpointInfo, SessionIndex, SessionInfo, StorageBackend, StorageError, WalEntry,
};
use tracing::{debug, instrument, warn};

use crate::kv::KvClient;

/// R2 storage backend using Cloudflare R2 (S3-compatible) for objects and KV for index.
///
/// Storage layout in R2:
/// ```
/// {bucket}/
///   {tenant_id}/
///     sessions/
///       {session_id}.docx          # Session document
///       {session_id}.wal           # WAL file (JSONL format)
///       {session_id}.ckpt.{pos}.docx  # Checkpoint files
/// ```
///
/// Index stored in KV:
/// ```
/// Key: index:{tenant_id}
/// Value: JSON-serialized SessionIndex
/// ```
#[derive(Clone)]
pub struct R2Storage {
    s3_client: S3Client,
    kv_client: Arc<KvClient>,
    bucket_name: String,
}

impl R2Storage {
    /// Create a new R2Storage backend.
    pub fn new(s3_client: S3Client, kv_client: Arc<KvClient>, bucket_name: String) -> Self {
        Self {
            s3_client,
            kv_client,
            bucket_name,
        }
    }

    /// Get the S3 key for a session document.
    fn session_key(&self, tenant_id: &str, session_id: &str) -> String {
        format!("{}/sessions/{}.docx", tenant_id, session_id)
    }

    /// Get the S3 key for a session WAL file.
    fn wal_key(&self, tenant_id: &str, session_id: &str) -> String {
        format!("{}/sessions/{}.wal", tenant_id, session_id)
    }

    /// Get the S3 key for a checkpoint.
    fn checkpoint_key(&self, tenant_id: &str, session_id: &str, position: u64) -> String {
        format!("{}/sessions/{}.ckpt.{}.docx", tenant_id, session_id, position)
    }

    /// Get the KV key for a tenant's index.
    fn index_kv_key(&self, tenant_id: &str) -> String {
        format!("index:{}", tenant_id)
    }

    /// Get an object from R2.
    async fn get_object(&self, key: &str) -> Result<Option<Vec<u8>>, StorageError> {
        let result = self
            .s3_client
            .get_object()
            .bucket(&self.bucket_name)
            .key(key)
            .send()
            .await;

        match result {
            Ok(output) => {
                let bytes = output
                    .body
                    .collect()
                    .await
                    .map_err(|e| StorageError::Io(format!("Failed to read R2 object body: {}", e)))?
                    .into_bytes();
                Ok(Some(bytes.to_vec()))
            }
            Err(e) => {
                let service_error = e.into_service_error();
                if service_error.is_no_such_key() {
                    Ok(None)
                } else {
                    Err(StorageError::Io(format!("R2 get_object error: {}", service_error)))
                }
            }
        }
    }

    /// Put an object to R2.
    async fn put_object(&self, key: &str, data: &[u8]) -> Result<(), StorageError> {
        self.s3_client
            .put_object()
            .bucket(&self.bucket_name)
            .key(key)
            .body(ByteStream::from(data.to_vec()))
            .send()
            .await
            .map_err(|e| StorageError::Io(format!("R2 put_object error: {}", e)))?;
        Ok(())
    }

    /// Delete an object from R2.
    async fn delete_object(&self, key: &str) -> Result<(), StorageError> {
        self.s3_client
            .delete_object()
            .bucket(&self.bucket_name)
            .key(key)
            .send()
            .await
            .map_err(|e| StorageError::Io(format!("R2 delete_object error: {}", e)))?;
        Ok(())
    }

    /// List objects with a prefix.
    async fn list_objects(&self, prefix: &str) -> Result<Vec<String>, StorageError> {
        let mut keys = Vec::new();
        let mut continuation_token: Option<String> = None;

        loop {
            let mut request = self
                .s3_client
                .list_objects_v2()
                .bucket(&self.bucket_name)
                .prefix(prefix);

            if let Some(token) = continuation_token.take() {
                request = request.continuation_token(token);
            }

            let output = request
                .send()
                .await
                .map_err(|e| StorageError::Io(format!("R2 list_objects error: {}", e)))?;

            if let Some(contents) = output.contents {
                for obj in contents {
                    if let Some(key) = obj.key {
                        keys.push(key);
                    }
                }
            }

            if output.is_truncated.unwrap_or(false) {
                continuation_token = output.next_continuation_token;
            } else {
                break;
            }
        }

        Ok(keys)
    }
}

#[async_trait]
impl StorageBackend for R2Storage {
    fn backend_name(&self) -> &'static str {
        "r2"
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
        let key = self.session_key(tenant_id, session_id);
        let result = self.get_object(&key).await?;
        if result.is_some() {
            debug!("Loaded session {} from R2", session_id);
        }
        Ok(result)
    }

    #[instrument(skip(self, data), level = "debug", fields(data_len = data.len()))]
    async fn save_session(
        &self,
        tenant_id: &str,
        session_id: &str,
        data: &[u8],
    ) -> Result<(), StorageError> {
        let key = self.session_key(tenant_id, session_id);
        self.put_object(&key, data).await?;
        debug!("Saved session {} to R2 ({} bytes)", session_id, data.len());
        Ok(())
    }

    #[instrument(skip(self), level = "debug")]
    async fn delete_session(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<bool, StorageError> {
        let session_key = self.session_key(tenant_id, session_id);
        let wal_key = self.wal_key(tenant_id, session_id);

        // Check if session exists
        let existed = self.get_object(&session_key).await?.is_some();

        // Delete session file
        if let Err(e) = self.delete_object(&session_key).await {
            warn!("Failed to delete session file: {}", e);
        }

        // Delete WAL
        if let Err(e) = self.delete_object(&wal_key).await {
            warn!("Failed to delete WAL file: {}", e);
        }

        // Delete all checkpoints
        let checkpoints = self.list_checkpoints(tenant_id, session_id).await?;
        for ckpt in checkpoints {
            let ckpt_key = self.checkpoint_key(tenant_id, session_id, ckpt.position);
            if let Err(e) = self.delete_object(&ckpt_key).await {
                warn!("Failed to delete checkpoint: {}", e);
            }
        }

        debug!("Deleted session {} (existed: {})", session_id, existed);
        Ok(existed)
    }

    #[instrument(skip(self), level = "debug")]
    async fn list_sessions(&self, tenant_id: &str) -> Result<Vec<SessionInfo>, StorageError> {
        let prefix = format!("{}/sessions/", tenant_id);
        let keys = self.list_objects(&prefix).await?;

        let mut sessions = Vec::new();
        for key in keys {
            // Only include .docx files that aren't checkpoints
            if key.ends_with(".docx") && !key.contains(".ckpt.") {
                let session_id = key
                    .strip_prefix(&prefix)
                    .and_then(|s| s.strip_suffix(".docx"))
                    .unwrap_or_default()
                    .to_string();

                if !session_id.is_empty() {
                    // Get object metadata for size/timestamps
                    let head = self
                        .s3_client
                        .head_object()
                        .bucket(&self.bucket_name)
                        .key(&key)
                        .send()
                        .await;

                    let (size_bytes, modified_at) = match head {
                        Ok(output) => {
                            let size = output.content_length.unwrap_or(0) as u64;
                            let modified = output
                                .last_modified
                                .and_then(|dt| {
                                    chrono::DateTime::from_timestamp(dt.secs(), dt.subsec_nanos())
                                })
                                .unwrap_or_else(chrono::Utc::now);
                            (size, modified)
                        }
                        Err(_) => (0, chrono::Utc::now()),
                    };

                    sessions.push(SessionInfo {
                        session_id,
                        source_path: None,
                        created_at: modified_at, // R2 doesn't store creation time
                        modified_at,
                        size_bytes,
                    });
                }
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
        let key = self.session_key(tenant_id, session_id);
        let result = self
            .s3_client
            .head_object()
            .bucket(&self.bucket_name)
            .key(&key)
            .send()
            .await;

        match result {
            Ok(_) => Ok(true),
            Err(e) => {
                let service_error = e.into_service_error();
                if service_error.is_not_found() {
                    Ok(false)
                } else {
                    Err(StorageError::Io(format!("R2 head_object error: {}", service_error)))
                }
            }
        }
    }

    // =========================================================================
    // Index Operations (stored in KV for fast access)
    // =========================================================================

    #[instrument(skip(self), level = "debug")]
    async fn load_index(&self, tenant_id: &str) -> Result<Option<SessionIndex>, StorageError> {
        let key = self.index_kv_key(tenant_id);
        match self.kv_client.get(&key).await? {
            Some(json) => {
                let index: SessionIndex = serde_json::from_str(&json).map_err(|e| {
                    StorageError::Serialization(format!("Failed to parse index: {}", e))
                })?;
                debug!("Loaded index with {} sessions from KV", index.sessions.len());
                Ok(Some(index))
            }
            None => Ok(None),
        }
    }

    #[instrument(skip(self, index), level = "debug", fields(sessions = index.sessions.len()))]
    async fn save_index(
        &self,
        tenant_id: &str,
        index: &SessionIndex,
    ) -> Result<(), StorageError> {
        let key = self.index_kv_key(tenant_id);
        let json = serde_json::to_string(index).map_err(|e| {
            StorageError::Serialization(format!("Failed to serialize index: {}", e))
        })?;
        self.kv_client.put(&key, &json).await?;
        debug!("Saved index with {} sessions to KV", index.sessions.len());
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

        let key = self.wal_key(tenant_id, session_id);

        // .NET MappedWal format:
        // - 8 bytes: little-endian i64 = data length (NOT including header)
        // - JSONL data: each entry is a JSON line ending with \n

        // Read existing WAL or create new
        let mut wal_data = match self.get_object(&key).await? {
            Some(data) if data.len() >= 8 => {
                // Parse header to get data length
                let data_len = i64::from_le_bytes(data[..8].try_into().unwrap()) as usize;
                let used_len = 8 + data_len;
                let mut truncated = data;
                truncated.truncate(used_len.min(truncated.len()));
                truncated
            }
            _ => {
                // New file - start with 8-byte header (data_len = 0)
                vec![0u8; 8]
            }
        };

        // Append new entries as JSONL
        let mut last_position = 0u64;
        for entry in entries {
            wal_data.extend_from_slice(&entry.patch_json);
            if !entry.patch_json.ends_with(b"\n") {
                wal_data.push(b'\n');
            }
            last_position = entry.position;
        }

        // Update header with data length
        let data_len = (wal_data.len() - 8) as i64;
        wal_data[..8].copy_from_slice(&data_len.to_le_bytes());

        // Write back to R2
        self.put_object(&key, &wal_data).await?;

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
        let key = self.wal_key(tenant_id, session_id);

        let raw_data = match self.get_object(&key).await? {
            Some(data) => data,
            None => return Ok((vec![], false)),
        };

        if raw_data.len() < 8 {
            return Ok((vec![], false));
        }

        // Parse header
        let data_len = i64::from_le_bytes(raw_data[..8].try_into().unwrap()) as usize;
        if data_len == 0 {
            return Ok((vec![], false));
        }

        // Extract JSONL portion
        let end = (8 + data_len).min(raw_data.len());
        let jsonl_data = &raw_data[8..end];

        let content = std::str::from_utf8(jsonl_data).map_err(|e| {
            StorageError::Io(format!("WAL is not valid UTF-8: {}", e))
        })?;

        // Parse JSONL - each line is a .NET WalEntry JSON
        let mut entries = Vec::new();
        let limit = limit.unwrap_or(u64::MAX);
        let mut position = 1u64;

        for line in content.lines() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }

            if position >= from_position {
                let value: serde_json::Value = serde_json::from_str(line).map_err(|e| {
                    StorageError::Serialization(format!(
                        "Failed to parse WAL entry at position {}: {}",
                        position, e
                    ))
                })?;

                let timestamp = value
                    .get("timestamp")
                    .and_then(|v| v.as_str())
                    .and_then(|s| chrono::DateTime::parse_from_rfc3339(s).ok())
                    .map(|dt| dt.with_timezone(&chrono::Utc))
                    .unwrap_or_else(chrono::Utc::now);

                entries.push(WalEntry {
                    position,
                    operation: String::new(),
                    path: String::new(),
                    patch_json: line.as_bytes().to_vec(),
                    timestamp,
                });

                if entries.len() as u64 >= limit {
                    return Ok((entries, true));
                }
            }

            position += 1;
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
        keep_count: u64,
    ) -> Result<u64, StorageError> {
        let (entries, _) = self.read_wal(tenant_id, session_id, 0, None).await?;

        let (to_keep, to_remove): (Vec<_>, Vec<_>) =
            entries.into_iter().partition(|e| e.position <= keep_count);

        let removed_count = to_remove.len() as u64;

        if removed_count == 0 {
            return Ok(0);
        }

        // Rewrite WAL with only kept entries
        let key = self.wal_key(tenant_id, session_id);
        let mut wal_data = vec![0u8; 8]; // Header placeholder

        for entry in &to_keep {
            wal_data.extend_from_slice(&entry.patch_json);
            if !entry.patch_json.ends_with(b"\n") {
                wal_data.push(b'\n');
            }
        }

        // Update header
        let data_len = (wal_data.len() - 8) as i64;
        wal_data[..8].copy_from_slice(&data_len.to_le_bytes());

        self.put_object(&key, &wal_data).await?;

        debug!(
            "Truncated WAL, removed {} entries, kept {}",
            removed_count,
            to_keep.len()
        );
        Ok(removed_count)
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
        let key = self.checkpoint_key(tenant_id, session_id, position);
        self.put_object(&key, data).await?;
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
                let key = self.checkpoint_key(tenant_id, session_id, latest.position);
                if let Some(data) = self.get_object(&key).await? {
                    debug!(
                        "Loaded latest checkpoint at position {} ({} bytes)",
                        latest.position,
                        data.len()
                    );
                    return Ok(Some((data, latest.position)));
                }
            }
            return Ok(None);
        }

        let key = self.checkpoint_key(tenant_id, session_id, position);
        match self.get_object(&key).await? {
            Some(data) => {
                debug!(
                    "Loaded checkpoint at position {} ({} bytes)",
                    position,
                    data.len()
                );
                Ok(Some((data, position)))
            }
            None => Ok(None),
        }
    }

    #[instrument(skip(self), level = "debug")]
    async fn list_checkpoints(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Vec<CheckpointInfo>, StorageError> {
        let prefix = format!("{}/sessions/{}.ckpt.", tenant_id, session_id);
        let keys = self.list_objects(&prefix).await?;

        let mut checkpoints = Vec::new();
        for key in keys {
            if key.ends_with(".docx") {
                // Extract position from key: {tenant}/sessions/{session}.ckpt.{position}.docx
                let position_str = key
                    .strip_prefix(&prefix)
                    .and_then(|s| s.strip_suffix(".docx"))
                    .unwrap_or("0");

                if let Ok(position) = position_str.parse::<u64>() {
                    // Get object metadata
                    let head = self
                        .s3_client
                        .head_object()
                        .bucket(&self.bucket_name)
                        .key(&key)
                        .send()
                        .await;

                    let (size_bytes, created_at) = match head {
                        Ok(output) => {
                            let size = output.content_length.unwrap_or(0) as u64;
                            let created = output
                                .last_modified
                                .and_then(|dt| {
                                    chrono::DateTime::from_timestamp(dt.secs(), dt.subsec_nanos())
                                })
                                .unwrap_or_else(chrono::Utc::now);
                            (size, created)
                        }
                        Err(_) => (0, chrono::Utc::now()),
                    };

                    checkpoints.push(CheckpointInfo {
                        position,
                        created_at,
                        size_bytes,
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

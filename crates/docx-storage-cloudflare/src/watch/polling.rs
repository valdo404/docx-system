use async_trait::async_trait;
use aws_sdk_s3::Client as S3Client;
use dashmap::DashMap;
use docx_storage_core::{
    ExternalChangeEvent, ExternalChangeType, SourceDescriptor, SourceMetadata, SourceType,
    StorageError, WatchBackend,
};
use tracing::{debug, instrument};

/// State for a watched source
#[derive(Debug, Clone)]
struct WatchedSource {
    source: SourceDescriptor,
    #[allow(dead_code)]
    watch_id: String,
    known_metadata: Option<SourceMetadata>,
    #[allow(dead_code)]
    poll_interval_secs: u32,
}

/// Polling-based watch backend for R2/S3 sources.
///
/// R2 doesn't support push notifications, so we poll for changes
/// by checking ETag/LastModified metadata.
pub struct PollingWatchBackend {
    /// S3 client for R2 operations
    s3_client: S3Client,
    /// Default bucket
    default_bucket: String,
    /// Watched sources: (tenant_id, session_id) -> WatchedSource
    sources: DashMap<(String, String), WatchedSource>,
    /// Pending change events detected during polling
    pending_changes: DashMap<(String, String), ExternalChangeEvent>,
    /// Default poll interval (seconds)
    default_poll_interval: u32,
}

impl PollingWatchBackend {
    /// Create a new PollingWatchBackend.
    pub fn new(s3_client: S3Client, default_bucket: String, default_poll_interval: u32) -> Self {
        Self {
            s3_client,
            default_bucket,
            sources: DashMap::new(),
            pending_changes: DashMap::new(),
            default_poll_interval,
        }
    }

    /// Get the key for the sources map.
    fn key(tenant_id: &str, session_id: &str) -> (String, String) {
        (tenant_id.to_string(), session_id.to_string())
    }

    /// Parse R2/S3 URI into bucket and key.
    fn parse_uri(uri: &str) -> Option<(String, String)> {
        let uri = uri
            .strip_prefix("r2://")
            .or_else(|| uri.strip_prefix("s3://"))?;

        let mut parts = uri.splitn(2, '/');
        let bucket = parts.next()?.to_string();
        let key = parts.next().unwrap_or("").to_string();
        Some((bucket, key))
    }

    /// Get metadata for an R2/S3 object.
    async fn get_object_metadata(
        &self,
        bucket: &str,
        key: &str,
    ) -> Result<Option<SourceMetadata>, StorageError> {
        let bucket = if bucket.is_empty() {
            &self.default_bucket
        } else {
            bucket
        };

        let result = self
            .s3_client
            .head_object()
            .bucket(bucket)
            .key(key)
            .send()
            .await;

        match result {
            Ok(output) => {
                let size_bytes = output.content_length.unwrap_or(0) as u64;
                let modified_at = output
                    .last_modified
                    .and_then(|dt| Some(dt.secs()))
                    .unwrap_or(0);
                let etag = output.e_tag;
                let version_id = output.version_id;

                // For R2, we don't have direct content hash access,
                // but ETag is typically the MD5 hash (or multipart upload identifier)
                // We could compute SHA256 if needed, but ETag is sufficient for change detection
                let content_hash = etag.as_ref().and_then(|e| {
                    // Strip quotes from ETag
                    let e = e.trim_matches('"');
                    // If it's a valid hex string (MD5), use it
                    hex::decode(e).ok()
                });

                Ok(Some(SourceMetadata {
                    size_bytes,
                    modified_at,
                    etag,
                    version_id,
                    content_hash,
                }))
            }
            Err(e) => {
                let service_error = e.into_service_error();
                if service_error.is_not_found() {
                    Ok(None)
                } else {
                    Err(StorageError::Watch(format!(
                        "R2 head_object error: {}",
                        service_error
                    )))
                }
            }
        }
    }

    /// Compare metadata to detect changes.
    fn has_changed(old: &SourceMetadata, new: &SourceMetadata) -> bool {
        // Prefer ETag comparison (most reliable for R2)
        if let (Some(old_etag), Some(new_etag)) = (&old.etag, &new.etag) {
            return old_etag != new_etag;
        }

        // Fall back to version ID
        if let (Some(old_ver), Some(new_ver)) = (&old.version_id, &new.version_id) {
            return old_ver != new_ver;
        }

        // Fall back to content hash
        if let (Some(old_hash), Some(new_hash)) = (&old.content_hash, &new.content_hash) {
            return old_hash != new_hash;
        }

        // Last resort: size and mtime
        old.size_bytes != new.size_bytes || old.modified_at != new.modified_at
    }
}

#[async_trait]
impl WatchBackend for PollingWatchBackend {
    #[instrument(skip(self), level = "debug")]
    async fn start_watch(
        &self,
        tenant_id: &str,
        session_id: &str,
        source: &SourceDescriptor,
        poll_interval_secs: u32,
    ) -> Result<String, StorageError> {
        // Validate source type
        if source.source_type != SourceType::R2 && source.source_type != SourceType::S3 {
            return Err(StorageError::Watch(format!(
                "PollingWatchBackend only supports R2/S3 sources, got {:?}",
                source.source_type
            )));
        }

        let (bucket, key) = Self::parse_uri(&source.uri).ok_or_else(|| {
            StorageError::Watch(format!("Invalid R2/S3 URI: {}", source.uri))
        })?;

        let watch_id = uuid::Uuid::new_v4().to_string();
        let map_key = Self::key(tenant_id, session_id);

        // Get initial metadata
        let known_metadata = self.get_object_metadata(&bucket, &key).await?;

        let poll_interval = if poll_interval_secs > 0 {
            poll_interval_secs
        } else {
            self.default_poll_interval
        };

        // Store the watch info
        self.sources.insert(
            map_key,
            WatchedSource {
                source: source.clone(),
                watch_id: watch_id.clone(),
                known_metadata,
                poll_interval_secs: poll_interval,
            },
        );

        debug!(
            "Started polling watch for {} (tenant {} session {}, interval {} secs)",
            source.uri, tenant_id, session_id, poll_interval
        );

        Ok(watch_id)
    }

    #[instrument(skip(self), level = "debug")]
    async fn stop_watch(&self, tenant_id: &str, session_id: &str) -> Result<(), StorageError> {
        let key = Self::key(tenant_id, session_id);

        if let Some((_, watched)) = self.sources.remove(&key) {
            debug!(
                "Stopped watching {} for tenant {} session {}",
                watched.source.uri, tenant_id, session_id
            );
        }

        // Also remove any pending changes
        self.pending_changes.remove(&key);

        Ok(())
    }

    #[instrument(skip(self), level = "debug")]
    async fn check_for_changes(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Option<ExternalChangeEvent>, StorageError> {
        let key = Self::key(tenant_id, session_id);

        // Check for pending changes first
        if let Some((_, event)) = self.pending_changes.remove(&key) {
            return Ok(Some(event));
        }

        // Get watched source
        let watched = match self.sources.get(&key) {
            Some(w) => w.clone(),
            None => return Ok(None),
        };

        // Parse URI
        let (bucket, obj_key) = match Self::parse_uri(&watched.source.uri) {
            Some((b, k)) => (b, k),
            None => return Ok(None),
        };

        // Get current metadata
        let current_metadata = match self.get_object_metadata(&bucket, &obj_key).await? {
            Some(m) => m,
            None => {
                // Object was deleted
                if watched.known_metadata.is_some() {
                    let event = ExternalChangeEvent {
                        session_id: session_id.to_string(),
                        change_type: ExternalChangeType::Deleted,
                        old_metadata: watched.known_metadata.clone(),
                        new_metadata: None,
                        detected_at: chrono::Utc::now().timestamp(),
                        new_uri: None,
                    };
                    return Ok(Some(event));
                }
                return Ok(None);
            }
        };

        // Compare with known metadata
        if let Some(known) = &watched.known_metadata {
            if Self::has_changed(known, &current_metadata) {
                debug!(
                    "Detected change in {} (ETag: {:?} -> {:?})",
                    watched.source.uri, known.etag, current_metadata.etag
                );

                let event = ExternalChangeEvent {
                    session_id: session_id.to_string(),
                    change_type: ExternalChangeType::Modified,
                    old_metadata: Some(known.clone()),
                    new_metadata: Some(current_metadata),
                    detected_at: chrono::Utc::now().timestamp(),
                    new_uri: None,
                };

                return Ok(Some(event));
            }
        }

        Ok(None)
    }

    #[instrument(skip(self), level = "debug")]
    async fn get_source_metadata(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Option<SourceMetadata>, StorageError> {
        let key = Self::key(tenant_id, session_id);

        let watched = match self.sources.get(&key) {
            Some(w) => w.clone(),
            None => return Ok(None),
        };

        let (bucket, obj_key) = match Self::parse_uri(&watched.source.uri) {
            Some((b, k)) => (b, k),
            None => return Ok(None),
        };

        self.get_object_metadata(&bucket, &obj_key).await
    }

    #[instrument(skip(self), level = "debug")]
    async fn get_known_metadata(
        &self,
        tenant_id: &str,
        session_id: &str,
    ) -> Result<Option<SourceMetadata>, StorageError> {
        let key = Self::key(tenant_id, session_id);

        Ok(self
            .sources
            .get(&key)
            .and_then(|w| w.known_metadata.clone()))
    }

    #[instrument(skip(self, metadata), level = "debug")]
    async fn update_known_metadata(
        &self,
        tenant_id: &str,
        session_id: &str,
        metadata: SourceMetadata,
    ) -> Result<(), StorageError> {
        let key = Self::key(tenant_id, session_id);

        if let Some(mut watched) = self.sources.get_mut(&key) {
            watched.known_metadata = Some(metadata);
            debug!(
                "Updated known metadata for tenant {} session {}",
                tenant_id, session_id
            );
        }

        Ok(())
    }
}

use std::collections::HashMap;
use std::sync::Mutex;
use std::time::Duration;

use async_trait::async_trait;
use docx_storage_core::{LockAcquireResult, LockManager, StorageError};
use serde::{Deserialize, Serialize};
use tracing::{debug, instrument};

use crate::kv::KvClient;
use std::sync::Arc;

/// Lock data stored in KV.
#[derive(Debug, Clone, Serialize, Deserialize)]
struct LockData {
    holder_id: String,
    acquired_at: i64,
    expires_at: i64,
}

/// KV-based distributed lock manager.
///
/// Uses Cloudflare KV for distributed locking with TTL-based expiration.
/// This is eventually consistent, so there's a small window for races,
/// but it's acceptable for our use case (optimistic locking with retries).
///
/// Lock keys: `lock:{tenant_id}:{resource_id}`
pub struct KvLock {
    kv_client: Arc<KvClient>,
    /// Local cache of acquired locks to avoid unnecessary KV calls
    local_locks: Mutex<HashMap<(String, String), String>>,
}

impl KvLock {
    /// Create a new KvLock.
    pub fn new(kv_client: Arc<KvClient>) -> Self {
        Self {
            kv_client,
            local_locks: Mutex::new(HashMap::new()),
        }
    }

    /// Get the KV key for a lock.
    fn lock_key(tenant_id: &str, resource_id: &str) -> String {
        format!("lock:{}:{}", tenant_id, resource_id)
    }
}

#[async_trait]
impl LockManager for KvLock {
    #[instrument(skip(self), level = "debug")]
    async fn acquire(
        &self,
        tenant_id: &str,
        resource_id: &str,
        holder_id: &str,
        ttl: Duration,
    ) -> Result<LockAcquireResult, StorageError> {
        let key = Self::lock_key(tenant_id, resource_id);
        let local_key = (tenant_id.to_string(), resource_id.to_string());

        // Check if we already hold this lock locally
        {
            let local_locks = self.local_locks.lock().unwrap();
            if let Some(existing_holder) = local_locks.get(&local_key) {
                if existing_holder == holder_id {
                    debug!(
                        "Lock on {}/{} already held by {} (local cache)",
                        tenant_id, resource_id, holder_id
                    );
                    return Ok(LockAcquireResult::acquired());
                } else {
                    debug!(
                        "Lock on {}/{} held by {} (requested by {})",
                        tenant_id, resource_id, existing_holder, holder_id
                    );
                    return Ok(LockAcquireResult::not_acquired());
                }
            }
        }

        let now = chrono::Utc::now().timestamp();
        let expires_at = now + ttl.as_secs() as i64;

        // Check if lock exists and is still valid
        if let Some(existing) = self.kv_client.get(&key).await? {
            if let Ok(lock_data) = serde_json::from_str::<LockData>(&existing) {
                if lock_data.expires_at > now {
                    // Lock is still held
                    if lock_data.holder_id == holder_id {
                        // We already hold it (reentrant)
                        debug!(
                            "Lock on {}/{} already held by {} (reentrant)",
                            tenant_id, resource_id, holder_id
                        );
                        let mut local_locks = self.local_locks.lock().unwrap();
                        local_locks.insert(local_key, holder_id.to_string());
                        return Ok(LockAcquireResult::acquired());
                    } else {
                        // Someone else holds it
                        debug!(
                            "Lock on {}/{} held by {} until {} (requested by {})",
                            tenant_id,
                            resource_id,
                            lock_data.holder_id,
                            lock_data.expires_at,
                            holder_id
                        );
                        return Ok(LockAcquireResult::not_acquired());
                    }
                }
                // Lock expired, we can take it
                debug!(
                    "Lock on {}/{} expired (was held by {}), acquiring for {}",
                    tenant_id, resource_id, lock_data.holder_id, holder_id
                );
            }
        }

        // Try to acquire the lock
        let lock_data = LockData {
            holder_id: holder_id.to_string(),
            acquired_at: now,
            expires_at,
        };
        let lock_json = serde_json::to_string(&lock_data).map_err(|e| {
            StorageError::Serialization(format!("Failed to serialize lock data: {}", e))
        })?;

        self.kv_client.put(&key, &lock_json).await?;

        // Add to local cache
        {
            let mut local_locks = self.local_locks.lock().unwrap();
            local_locks.insert(local_key, holder_id.to_string());
        }

        debug!(
            "Acquired lock on {}/{} for {} (expires at {})",
            tenant_id, resource_id, holder_id, expires_at
        );
        Ok(LockAcquireResult::acquired())
    }

    #[instrument(skip(self), level = "debug")]
    async fn release(
        &self,
        tenant_id: &str,
        resource_id: &str,
        holder_id: &str,
    ) -> Result<(), StorageError> {
        let key = Self::lock_key(tenant_id, resource_id);
        let local_key = (tenant_id.to_string(), resource_id.to_string());

        // Check if we hold this lock
        {
            let mut local_locks = self.local_locks.lock().unwrap();
            if let Some(existing_holder) = local_locks.get(&local_key) {
                if existing_holder != holder_id {
                    debug!(
                        "Cannot release lock on {}/{}: held by {} not {}",
                        tenant_id, resource_id, existing_holder, holder_id
                    );
                    return Ok(());
                }
                local_locks.remove(&local_key);
            }
        }

        // Verify in KV and delete
        if let Some(existing) = self.kv_client.get(&key).await? {
            if let Ok(lock_data) = serde_json::from_str::<LockData>(&existing) {
                if lock_data.holder_id == holder_id {
                    self.kv_client.delete(&key).await?;
                    debug!(
                        "Released lock on {}/{} by {}",
                        tenant_id, resource_id, holder_id
                    );
                } else {
                    debug!(
                        "Lock on {}/{} held by {} not {} (no-op)",
                        tenant_id, resource_id, lock_data.holder_id, holder_id
                    );
                }
            }
        }

        Ok(())
    }
}

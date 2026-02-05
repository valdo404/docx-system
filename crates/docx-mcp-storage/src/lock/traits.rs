use std::time::Duration;

use async_trait::async_trait;

use crate::error::StorageError;

/// Result of a lock acquisition attempt.
#[derive(Debug, Clone)]
pub struct LockAcquireResult {
    /// Whether the lock was acquired.
    pub acquired: bool,
    /// If not acquired, who currently holds the lock.
    pub current_holder: Option<String>,
    /// Lock expiration timestamp (Unix epoch seconds).
    pub expires_at: i64,
}

/// Result of a lock release attempt.
#[derive(Debug, Clone)]
pub struct LockReleaseResult {
    /// Whether the lock was released.
    pub released: bool,
    /// Reason: "ok", "not_owner", "not_found", "expired"
    pub reason: String,
}

/// Result of a lock renewal attempt.
#[derive(Debug, Clone)]
pub struct LockRenewResult {
    /// Whether the lock was renewed.
    pub renewed: bool,
    /// New expiration timestamp.
    pub expires_at: i64,
    /// Reason: "ok", "not_owner", "not_found"
    pub reason: String,
}

/// Lock manager abstraction for tenant-aware distributed locking.
///
/// Locks are on the pair `(tenant_id, resource_id)` to ensure tenant isolation.
/// The maximum number of concurrent locks = T tenants × F files per tenant.
#[async_trait]
pub trait LockManager: Send + Sync {
    /// Returns the lock manager identifier (e.g., "file", "kv").
    fn lock_type(&self) -> &'static str;

    /// Attempt to acquire a lock on `(tenant_id, resource_id)`.
    ///
    /// # Arguments
    /// * `tenant_id` - Tenant identifier for isolation
    /// * `resource_id` - Resource to lock (e.g., session_id)
    /// * `holder_id` - Unique identifier for this lock holder (UUID recommended)
    /// * `ttl` - Time-to-live for the lock to prevent orphaned locks
    ///
    /// # Returns
    /// * `Ok(result)` - Lock result with acquisition status
    async fn acquire(
        &self,
        tenant_id: &str,
        resource_id: &str,
        holder_id: &str,
        ttl: Duration,
    ) -> Result<LockAcquireResult, StorageError>;

    /// Release a lock.
    ///
    /// The lock is only released if `holder_id` matches the current holder.
    async fn release(
        &self,
        tenant_id: &str,
        resource_id: &str,
        holder_id: &str,
    ) -> Result<LockReleaseResult, StorageError>;

    /// Renew a lock's TTL.
    ///
    /// The lock is only renewed if `holder_id` matches the current holder.
    async fn renew(
        &self,
        tenant_id: &str,
        resource_id: &str,
        holder_id: &str,
        ttl: Duration,
    ) -> Result<LockRenewResult, StorageError>;
}

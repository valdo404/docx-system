mod kv_lock;

pub use kv_lock::KvLock;

// Re-export from core
pub use docx_storage_core::LockManager;

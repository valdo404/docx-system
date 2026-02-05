use thiserror::Error;

/// Errors that can occur in the storage layer.
#[derive(Error, Debug)]
pub enum StorageError {
    #[error("I/O error: {0}")]
    Io(String),

    #[error("Serialization error: {0}")]
    Serialization(String),

    #[error("Not found: {0}")]
    NotFound(String),

    #[error("Lock error: {0}")]
    Lock(String),

    #[error("Invalid argument: {0}")]
    InvalidArgument(String),

    #[error("Internal error: {0}")]
    Internal(String),
}

impl From<StorageError> for tonic::Status {
    fn from(err: StorageError) -> Self {
        match err {
            StorageError::Io(msg) => tonic::Status::internal(msg),
            StorageError::Serialization(msg) => tonic::Status::internal(msg),
            StorageError::NotFound(msg) => tonic::Status::not_found(msg),
            StorageError::Lock(msg) => tonic::Status::failed_precondition(msg),
            StorageError::InvalidArgument(msg) => tonic::Status::invalid_argument(msg),
            StorageError::Internal(msg) => tonic::Status::internal(msg),
        }
    }
}

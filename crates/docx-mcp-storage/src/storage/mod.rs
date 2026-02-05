mod traits;
mod local;

pub use traits::*;
pub use local::LocalStorage;

#[cfg(feature = "cloud")]
mod r2;
#[cfg(feature = "cloud")]
pub use r2::R2Storage;

use docx_storage_core::StorageError;
use reqwest::Client as HttpClient;
use tracing::{debug, instrument};

/// Cloudflare KV REST API client.
///
/// Uses the Cloudflare API v4 to interact with KV namespaces.
/// This provides faster access for index data compared to R2.
pub struct KvClient {
    http_client: HttpClient,
    account_id: String,
    namespace_id: String,
    api_token: String,
}

impl KvClient {
    /// Create a new KV client.
    pub fn new(
        account_id: String,
        namespace_id: String,
        api_token: String,
    ) -> Self {
        Self {
            http_client: HttpClient::new(),
            account_id,
            namespace_id,
            api_token,
        }
    }

    /// Base URL for KV API.
    fn base_url(&self) -> String {
        format!(
            "https://api.cloudflare.com/client/v4/accounts/{}/storage/kv/namespaces/{}",
            self.account_id, self.namespace_id
        )
    }

    /// Get a value from KV.
    #[instrument(skip(self), level = "debug")]
    pub async fn get(&self, key: &str) -> Result<Option<String>, StorageError> {
        let url = format!("{}/values/{}", self.base_url(), urlencoding::encode(key));

        let response = self
            .http_client
            .get(&url)
            .header("Authorization", format!("Bearer {}", self.api_token))
            .send()
            .await
            .map_err(|e| StorageError::Io(format!("KV GET request failed: {}", e)))?;

        let status = response.status();
        if status == reqwest::StatusCode::NOT_FOUND {
            debug!("KV key not found: {}", key);
            return Ok(None);
        }

        if !status.is_success() {
            let text = response.text().await.unwrap_or_default();
            return Err(StorageError::Io(format!(
                "KV GET failed with status {}: {}",
                status, text
            )));
        }

        // KV GET returns raw value, not JSON-wrapped
        let value = response
            .text()
            .await
            .map_err(|e| StorageError::Io(format!("Failed to read KV response: {}", e)))?;

        debug!("KV GET {} ({} bytes)", key, value.len());
        Ok(Some(value))
    }

    /// Put a value to KV.
    #[instrument(skip(self, value), level = "debug", fields(value_len = value.len()))]
    pub async fn put(&self, key: &str, value: &str) -> Result<(), StorageError> {
        let url = format!("{}/values/{}", self.base_url(), urlencoding::encode(key));

        let response = self
            .http_client
            .put(&url)
            .header("Authorization", format!("Bearer {}", self.api_token))
            .header("Content-Type", "text/plain")
            .body(value.to_string())
            .send()
            .await
            .map_err(|e| StorageError::Io(format!("KV PUT request failed: {}", e)))?;

        let status = response.status();
        if !status.is_success() {
            let text = response.text().await.unwrap_or_default();
            return Err(StorageError::Io(format!(
                "KV PUT failed with status {}: {}",
                status, text
            )));
        }

        debug!("KV PUT {} ({} bytes)", key, value.len());
        Ok(())
    }

    /// Delete a value from KV.
    #[instrument(skip(self), level = "debug")]
    pub async fn delete(&self, key: &str) -> Result<bool, StorageError> {
        let url = format!("{}/values/{}", self.base_url(), urlencoding::encode(key));

        let response = self
            .http_client
            .delete(&url)
            .header("Authorization", format!("Bearer {}", self.api_token))
            .send()
            .await
            .map_err(|e| StorageError::Io(format!("KV DELETE request failed: {}", e)))?;

        let status = response.status();
        if status == reqwest::StatusCode::NOT_FOUND {
            return Ok(false);
        }

        if !status.is_success() {
            let text = response.text().await.unwrap_or_default();
            return Err(StorageError::Io(format!(
                "KV DELETE failed with status {}: {}",
                status, text
            )));
        }

        debug!("KV DELETE {}", key);
        Ok(true)
    }
}

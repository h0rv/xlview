//! Structured error types for xlview.
//!
//! Replaces `Result<T, String>` throughout the codebase with proper error types.

/// All errors that can occur in xlview parsing and rendering.
#[derive(Debug, thiserror::Error)]
pub enum XlviewError {
    /// XML parsing error from quick-xml.
    #[error("XML parsing: {0}")]
    Xml(#[from] quick_xml::Error),

    /// ZIP archive error.
    #[error("ZIP archive: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// Invalid cell reference.
    #[error("Invalid cell reference: {0}")]
    CellRef(String),

    /// Style resolution failure.
    #[error("Style resolution failed: {0}")]
    Style(String),

    /// Rendering error.
    #[error("Render error: {0}")]
    Render(String),

    /// General parse error.
    #[error("Parse error: {0}")]
    Parse(String),

    /// I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// Catch-all for string errors during migration.
    #[error("{0}")]
    Other(String),
}

/// Convenience alias used throughout the crate.
pub type Result<T> = std::result::Result<T, XlviewError>;

impl From<String> for XlviewError {
    fn from(s: String) -> Self {
        Self::Other(s)
    }
}

impl From<&str> for XlviewError {
    fn from(s: &str) -> Self {
        Self::Other(s.to_string())
    }
}

#[cfg(target_arch = "wasm32")]
impl From<XlviewError> for wasm_bindgen::JsValue {
    fn from(e: XlviewError) -> Self {
        wasm_bindgen::JsValue::from_str(&e.to_string())
    }
}

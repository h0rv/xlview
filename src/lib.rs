//! xlview - XLSX viewer for the web
//!
//! Parses and renders Excel files in the browser via WebAssembly and Canvas 2D:
//! - Full styling (fonts, colors, borders, fills, conditional formatting)
//! - Charts, images, shapes, sparklines
//! - Frozen panes, merged cells, multiple sheets
//! - 100k+ cells at 120fps
//! - Zero runtime dependencies
//!
//! # Usage (JavaScript)
//!
//! ```javascript
//! import init, { XlView } from 'xlview';
//! await init();
//! const viewer = XlView.new_with_overlay(canvas, overlay, dpr);
//! viewer.load(data);
//! viewer.render();
//! ```

// Parsing modules
pub mod auto_filter;
pub mod cell_ref;
pub mod charts;
pub mod color;
pub mod comments;
pub mod conditional;
pub mod data_validation;
pub mod drawings;
pub mod error;
pub mod formula;
pub mod hyperlinks;
pub mod named_styles;
pub mod namespaces;
pub mod numfmt;
pub mod outlines;
pub mod page_setup;
pub mod parser;
pub mod protection;
pub mod rich_text;
pub mod sparklines;
pub mod styles;
pub mod theme_parser;
pub mod types;
pub mod workbook_meta;
pub mod xml_helpers;

// Rendering modules (Canvas 2D)
pub mod layout;
pub mod render;
pub mod viewer;

use wasm_bindgen::prelude::*;

// Re-export the main viewer struct
pub use viewer::XlView;

pub use types::*;

/// Parse an XLSX file and return a JSON string representing the workbook
///
/// # Arguments
/// * `data` - The raw bytes of the XLSX file
///
/// # Returns
/// A JSON string containing the parsed workbook structure
///
/// # Errors
/// Returns an error if the XLSX file is invalid or cannot be parsed.
#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8]) -> Result<String, JsValue> {
    let workbook = parser::parse(data).map_err(|e| JsValue::from_str(&e.to_string()))?;

    serde_json::to_string(&workbook)
        .map_err(|e| JsValue::from_str(&format!("JSON serialization error: {e}")))
}

/// Parse an XLSX file and return the workbook as a `JsValue`
///
/// This is more efficient than `parse_xlsx` when the result will be
/// used directly in JavaScript.
///
/// # Errors
/// Returns an error if the XLSX file is invalid or cannot be parsed.
#[wasm_bindgen]
pub fn parse_xlsx_to_js(data: &[u8]) -> Result<JsValue, JsValue> {
    let workbook = parser::parse(data).map_err(|e| JsValue::from_str(&e.to_string()))?;

    serde_wasm_bindgen::to_value(&workbook)
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

/// Parse an XLSX file and return internal timing metrics (WASM only).
#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
pub fn parse_xlsx_metrics(data: &[u8]) -> Result<JsValue, JsValue> {
    let (_workbook, metrics) =
        parser::parse_with_metrics_lazy(data).map_err(|e| JsValue::from_str(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&metrics)
        .map_err(|e| JsValue::from_str(&format!("Serialization error: {e}")))
}

/// Get the library version
#[must_use]
#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

use serde::{Deserialize, Serialize};
use std::rc::Rc;

use super::{RichTextRun, StyleRef, TextRunData};

/// Cell with position
#[derive(Debug, Serialize, Deserialize)]
pub struct CellData {
    pub r: u32, // row (0-indexed)
    pub c: u32, // col (0-indexed)
    pub cell: Cell,
}

/// A single cell's data and style
#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    /// The display value (already formatted, or concatenated plain text for rich text)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v: Option<String>,
    /// Cell type: s=string, n=number, b=boolean, e=error, d=date
    pub t: CellType,
    /// Resolved style (not an index)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub s: Option<StyleRef>,
    /// Style index into workbook-resolved styles (lazy path).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_idx: Option<u32>,
    /// Raw cell value (lazy path; not serialized).
    #[serde(skip_serializing, skip_deserializing, default)]
    pub raw: Option<CellRawValue>,
    /// Cached display value for lazy formatting (not serialized).
    #[serde(skip_serializing, skip_deserializing, default)]
    pub cached_display: Option<String>,
    /// Rich text runs for cells with mixed formatting
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<Vec<RichTextRun>>,
    /// Cached render-friendly rich text runs (not serialized).
    #[serde(skip_serializing, skip_deserializing, default)]
    pub cached_rich_text: Option<Rc<Vec<TextRunData>>>,
    /// Whether the cell has a comment/note attached
    #[serde(skip_serializing_if = "Option::is_none")]
    pub has_comment: Option<bool>,
    /// Hyperlink associated with this cell
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<Hyperlink>,
    /// Formula text (preserved for roundtrip save; not serialized)
    #[serde(skip)]
    pub formula: Option<String>,
}

#[derive(Debug, Clone)]
pub enum CellRawValue {
    SharedString(u32),
    String(String),
    Number(f64),
    Boolean(bool),
    Error(String),
    Date(f64),
}

/// Hyperlink information for a cell
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    /// URL or internal reference target
    pub target: String,
    /// Bookmark/location within target (e.g., cell reference for internal links)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
    /// Hover text / tooltip
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    /// Whether this is an external URL (true) or internal reference (false)
    pub is_external: bool,
}

/// Hyperlink definition with cell reference (for bulk storage on sheet)
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HyperlinkDef {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// The hyperlink data
    pub hyperlink: Hyperlink,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    #[serde(rename = "s")]
    String,
    #[serde(rename = "n")]
    Number,
    #[serde(rename = "b")]
    Boolean,
    #[serde(rename = "e")]
    Error,
    #[serde(rename = "d")]
    Date,
}

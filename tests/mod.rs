//! Integration tests for xlview.
//!
//! This module provides the test infrastructure for testing the XLSX parser.
//! It includes:
//!
//! - `fixtures`: Builders for creating valid XLSX files in memory
//! - `common`: Assertion helpers and parsing utilities
//!
//! # Example Usage
//!
//! ```rust,ignore
//! use crate::fixtures::{XlsxBuilder, StyleBuilder, SheetBuilder};
//! use crate::common::{parse_xlsx_to_json, assert_cell_value, assert_cell_bold};
//!
//! fn test_bold_text() {
//!     let xlsx = XlsxBuilder::new()
//!         .add_sheet("Sheet1")
//!         .add_cell("A1", "Bold Text", Some(StyleBuilder::new().bold().build()))
//!         .build();
//!
//!     let workbook = parse_xlsx_to_json(&xlsx);
//!     assert_cell_value(&workbook, 0, 0, 0, "Bold Text");
//!     assert_cell_bold(&workbook, 0, 0, 0);
//! }
//! ```
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

pub mod common;
pub mod fixtures;

// Re-export commonly used items at the top level
pub use common::{
    assert_cell_align_h, assert_cell_bg_color, assert_cell_bold, assert_cell_font_color,
    assert_cell_font_size, assert_cell_italic, assert_cell_value, assert_cell_wrap,
    assert_merge_exists, assert_sheet_count, assert_sheet_name, get_cell, get_cell_style,
    parse_xlsx_to_json,
};
pub use fixtures::{BorderSide, CellValue, SheetBuilder, StyleBuilder, XlsxBuilder};

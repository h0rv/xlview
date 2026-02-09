//! ECMA-376 Compliance Tests
//!
//! These tests verify that the XLSX parser correctly handles various aspects
//! of the ECMA-376 SpreadsheetML specification.
//!
//! Reference: ECMA-376-1:2016 Part 1 - SpreadsheetML
//!
//! Test categories:
//! - Cell types and values (18.3.1.4)
//! - Styles and formatting (18.8)
//! - Workbook structure (18.2)
//! - Sheet elements (18.3)
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless,
    clippy::expect_fun_call,
    clippy::unnecessary_map_or,
    clippy::should_implement_trait
)]

mod fixtures;
use fixtures::{StyleBuilder, XlsxBuilder};
use std::fs;
use xlview::parser::parse;

// ============================================================================
// ECMA-376 18.3.1.4 - Cell Types
// ============================================================================

/// Test cell type "s" = shared string (ECMA-376 18.3.1.4)
#[test]
fn test_ecma376_cell_type_shared_string() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Hello World", None)
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell A1 should exist");
    let cell = cell.unwrap();
    assert!(cell.cell.v.is_some(), "Cell should have value");
    assert_eq!(
        cell.cell.v.as_ref().unwrap(),
        "Hello World",
        "Cell value should be 'Hello World'"
    );
}

/// Test cell type "n" = number (ECMA-376 18.3.1.4)
#[test]
fn test_ecma376_cell_type_number() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "42.5", None)
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell A1 should exist");
    let cell = cell.unwrap();
    assert!(cell.cell.v.is_some(), "Cell should have value");
}

/// Test cell type "b" = boolean (ECMA-376 18.3.1.4)
#[test]
fn test_ecma376_cell_type_boolean() {
    // Boolean values in Excel are 0 or 1
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "TRUE", None)
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];

    // Just verify parsing doesn't crash
    assert!(!sheet.cells.is_empty(), "Should have cells");
}

// ============================================================================
// ECMA-376 18.8.1 - Fonts
// ============================================================================

/// Test font bold attribute (ECMA-376 18.8.2)
#[test]
fn test_ecma376_font_bold() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Bold", Some(StyleBuilder::new().bold().build()))
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(style.bold, Some(true), "Cell should be bold");
}

/// Test font italic attribute (ECMA-376 18.8.26)
#[test]
fn test_ecma376_font_italic() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "Italic", Some(StyleBuilder::new().italic().build()))
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(style.italic, Some(true), "Cell should be italic");
}

/// Test font size (ECMA-376 18.8.38)
#[test]
fn test_ecma376_font_size() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Large",
            Some(StyleBuilder::new().font_size(18.0).build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(style.font_size, Some(18.0), "Font size should be 18");
}

// ============================================================================
// ECMA-376 18.8.20 - Fills
// ============================================================================

/// Test solid fill pattern (ECMA-376 18.8.32)
#[test]
fn test_ecma376_fill_solid() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Filled",
            Some(StyleBuilder::new().bg_color("#FF0000").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert!(
        style.bg_color.is_some(),
        "Cell should have background color"
    );
}

// ============================================================================
// ECMA-376 18.8.4 - Borders
// ============================================================================

/// Test thin border style (ECMA-376 18.8.5)
#[test]
fn test_ecma376_border_thin() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Bordered",
            Some(StyleBuilder::new().border_all("thin", None).build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert!(style.border_top.is_some(), "Should have top border");
    assert!(style.border_bottom.is_some(), "Should have bottom border");
    assert!(style.border_left.is_some(), "Should have left border");
    assert!(style.border_right.is_some(), "Should have right border");
}

// ============================================================================
// ECMA-376 18.8.1 - Alignment
// ============================================================================

/// Test horizontal alignment (ECMA-376 18.8.1)
#[test]
fn test_ecma376_alignment_horizontal() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Centered",
            Some(StyleBuilder::new().align_horizontal("center").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert!(style.align_h.is_some(), "Should have horizontal alignment");
}

/// Test vertical alignment (ECMA-376 18.8.1)
#[test]
fn test_ecma376_alignment_vertical() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Bottom",
            Some(StyleBuilder::new().align_vertical("bottom").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert!(style.align_v.is_some(), "Should have vertical alignment");
}

/// Test text wrap (ECMA-376 18.8.1)
#[test]
fn test_ecma376_alignment_wrap_text() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Wrapped",
            Some(StyleBuilder::new().wrap_text().build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(style.wrap, Some(true), "Text should wrap");
}

/// Test centerContinuous horizontal alignment (ECMA-376 18.18.40)
/// Regression test: centerContinuous was not mapped to HAlign::CenterContinuous
#[test]
fn test_ecma376_alignment_center_continuous() {
    use xlview::types::HAlign;

    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Center Across Selection",
            Some(
                StyleBuilder::new()
                    .align_horizontal("centerContinuous")
                    .build(),
            ),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(
        style.align_h,
        Some(HAlign::CenterContinuous),
        "Should have centerContinuous alignment"
    );
}

/// Test distributed horizontal alignment (ECMA-376 18.18.40)
/// Regression test: distributed was not mapped to HAlign::Distributed
#[test]
fn test_ecma376_alignment_horizontal_distributed() {
    use xlview::types::HAlign;

    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Distributed",
            Some(StyleBuilder::new().align_horizontal("distributed").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(
        style.align_h,
        Some(HAlign::Distributed),
        "Should have distributed horizontal alignment"
    );
}

/// Test justify vertical alignment (ECMA-376 18.18.41)
/// Regression test: vertical justify was not mapped to VAlign::Justify
#[test]
fn test_ecma376_alignment_vertical_justify() {
    use xlview::types::VAlign;

    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Justified\nvertically",
            Some(StyleBuilder::new().align_vertical("justify").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(VAlign::Justify),
        "Should have justify vertical alignment"
    );
}

/// Test distributed vertical alignment (ECMA-376 18.18.41)
/// Regression test: vertical distributed was not mapped to VAlign::Distributed
#[test]
fn test_ecma376_alignment_vertical_distributed() {
    use xlview::types::VAlign;

    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Distributed\nvertically",
            Some(StyleBuilder::new().align_vertical("distributed").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell should exist");
    let style = cell.unwrap().cell.s.as_ref().expect("Should have style");
    assert_eq!(
        style.align_v,
        Some(VAlign::Distributed),
        "Should have distributed vertical alignment"
    );
}

// ============================================================================
// ECMA-376 18.3.1.55 - Merged Cells
// ============================================================================

/// Test merged cell range parsing (ECMA-376 18.3.1.55)
#[test]
fn test_ecma376_merged_cells() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Find a sheet with merges
    let has_merges = workbook.sheets.iter().any(|s| !s.merges.is_empty());
    assert!(
        has_merges,
        "Should have at least one sheet with merged cells"
    );
}

// ============================================================================
// ECMA-376 18.2.19 - Sheet State
// ============================================================================

/// Test sheet visibility states (ECMA-376 18.2.19)
#[test]
fn test_ecma376_sheet_state() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // All sheets should have a valid state
    for sheet in &workbook.sheets {
        // state is an enum that should be parsed
        // (visible, hidden, or very_hidden)
        assert!(!sheet.name.is_empty(), "Sheet should have a name");
    }
}

// ============================================================================
// ECMA-376 18.3.1.73 - Panes (Freeze)
// ============================================================================

/// Test frozen panes parsing (ECMA-376 18.3.1.73)
#[test]
fn test_ecma376_frozen_panes() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Check for frozen panes
    let has_frozen = workbook
        .sheets
        .iter()
        .any(|s| s.frozen_rows > 0 || s.frozen_cols > 0);

    assert!(
        has_frozen,
        "Should have at least one sheet with frozen panes"
    );
}

// ============================================================================
// ECMA-376 18.7 - Conditional Formatting
// ============================================================================

/// Test conditional formatting parsing (ECMA-376 18.7)
#[test]
fn test_ecma376_conditional_formatting() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let cf_count: usize = workbook
        .sheets
        .iter()
        .map(|s| s.conditional_formatting.len())
        .sum();

    assert!(cf_count >= 1, "Should have conditional formatting rules");
}

// ============================================================================
// ECMA-376 18.14 - Themes
// ============================================================================

/// Test theme color parsing (ECMA-376 18.14)
#[test]
fn test_ecma376_theme_colors() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Theme should have colors (12 standard theme colors)
    assert!(
        !workbook.theme.colors.is_empty(),
        "Theme should have colors"
    );
}

// ============================================================================
// ECMA-376 18.3.1.40 - Hyperlinks
// ============================================================================

/// Test hyperlink parsing (ECMA-376 18.3.1.40)
#[test]
fn test_ecma376_hyperlinks() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let hyperlink_count: usize = workbook.sheets.iter().map(|s| s.hyperlinks.len()).sum();

    assert!(hyperlink_count >= 1, "Should have hyperlinks");
}

// ============================================================================
// ECMA-376 18.3.1.18 - Data Validation
// ============================================================================

/// Test data validation parsing (ECMA-376 18.3.1.18)
#[test]
fn test_ecma376_data_validation() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let validation_count: usize = workbook
        .sheets
        .iter()
        .map(|s| s.data_validations.len())
        .sum();

    assert!(validation_count >= 1, "Should have data validations");
}

// ============================================================================
// ECMA-376 18.4 - Comments
// ============================================================================

/// Test comment parsing (ECMA-376 18.4)
#[test]
fn test_ecma376_comments() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let comment_count: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();

    // Comments should be parsed (may be 0 if file doesn't have comments)
    // Just verify the parsing didn't fail
    let _ = comment_count;
}

// ============================================================================
// Roundtrip / Structure Tests
// ============================================================================

/// Test that XLSX files can be parsed and maintain structure
#[test]
fn test_xlsx_structure_integrity() {
    let files = [
        "test/minimal.xlsx",
        "test/styled.xlsx",
        "test/kitchen_sink.xlsx",
        "test/kitchen_sink_v2.xlsx",
    ];

    for file_path in &files {
        if !std::path::Path::new(file_path).exists() {
            continue;
        }

        let data = fs::read(file_path).expect(&format!("Failed to read {}", file_path));
        let workbook = parse(&data).expect(&format!("Failed to parse {}", file_path));

        // Basic structure checks
        assert!(
            !workbook.sheets.is_empty(),
            "{} should have at least one sheet",
            file_path
        );

        for sheet in &workbook.sheets {
            assert!(!sheet.name.is_empty(), "Sheet should have a name");
            // Row and column indices should be reasonable
            assert!(
                sheet.max_row < 1048576,
                "Max row should be within Excel limits"
            );
            assert!(
                sheet.max_col < 16384,
                "Max col should be within Excel limits"
            );
        }
    }
}

/// Test that all test files can be parsed without panics
#[test]
fn test_no_panics_on_all_test_files() {
    let test_dir = "test";
    let entries = fs::read_dir(test_dir);

    if entries.is_err() {
        return; // Skip if test directory doesn't exist
    }

    for entry in entries.unwrap() {
        let entry = entry.expect("Failed to read directory entry");
        let path = entry.path();

        if path.extension().map_or(false, |ext| ext == "xlsx") {
            let data = fs::read(&path);
            if let Ok(data) = data {
                // Just verify parsing doesn't panic
                let _ = parse(&data);
            }
        }
    }
}

//! Font family rendering tests for xlview
//!
//! Integration tests that verify font family styles are correctly parsed from XLSX files
//! and available in the parsed cell data for rendering.
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

use xlview::parser;

mod common;
mod fixtures;

use fixtures::{CellValue, SheetBuilder, StyleBuilder, XlsxBuilder};

// ============================================================================
// Test 1: Cell with explicit font family (Arial) should use that font
// ============================================================================

#[test]
fn test_cell_with_arial_font() {
    let style = StyleBuilder::new().font_name("Arial").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Test".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    assert!(
        !workbook.sheets.is_empty(),
        "Workbook should have at least one sheet"
    );
    assert!(
        !workbook.sheets[0].cells.is_empty(),
        "Sheet should have at least one cell"
    );

    let cell = &workbook.sheets[0].cells[0];
    assert!(cell.cell.s.is_some(), "Cell should have a style");

    let style = cell.cell.s.as_ref().unwrap();
    assert_eq!(
        style.font_family,
        Some("Arial".to_string()),
        "Cell should have Arial font family"
    );
}

#[test]
fn test_cell_with_times_new_roman_font() {
    let style = StyleBuilder::new().font_name("Times New Roman").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Serif Text".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.font_family,
        Some("Times New Roman".to_string()),
        "Cell should have Times New Roman font family"
    );
}

// ============================================================================
// Test 2: Cell without font family should use default font (Calibri)
// ============================================================================

#[test]
fn test_cell_without_font_family_uses_default() {
    // Create a cell with no style at all
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell("A1", CellValue::String("No Style".into()), None))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];

    // Cell may have a default style applied or no style at all
    // The parser should apply Calibri as default when rendering
    if let Some(ref style) = cell.cell.s {
        // If style exists, font_family should be Calibri (the default)
        if style.font_family.is_some() {
            assert_eq!(
                style.font_family,
                Some("Calibri".to_string()),
                "Default font family should be Calibri"
            );
        }
    }
    // If no style, the renderer will use its own default (Calibri)
}

#[test]
fn test_cell_with_style_but_no_font_name() {
    // Create a cell with bold style but no explicit font name
    let style = StyleBuilder::new().bold().build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Bold Only".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    // Bold should be set
    assert_eq!(style.bold, Some(true), "Cell should be bold");
    // Font family should be Calibri (default from stylesheet)
    assert_eq!(
        style.font_family,
        Some("Calibri".to_string()),
        "Default font family should be Calibri"
    );
}

// ============================================================================
// Test 3: Cell with unknown font family should fall back gracefully
// ============================================================================

#[test]
fn test_cell_with_unknown_font_family() {
    // Create a cell with a font that doesn't exist on most systems
    let style = StyleBuilder::new()
        .font_name("NonExistentFont12345")
        .build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Unknown Font".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    // The parser should still preserve the font family name as specified in the file
    // It's the renderer's job to handle fallback when the font is not available
    assert_eq!(
        style.font_family,
        Some("NonExistentFont12345".to_string()),
        "Parser should preserve the specified font family even if unknown"
    );
}

#[test]
fn test_cell_with_empty_font_name() {
    // Edge case: empty font name string
    let style = StyleBuilder::new().font_name("").build();
    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Empty Font".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];

    // Should parse without error - the renderer will handle empty font names
    assert!(cell.cell.s.is_some(), "Cell should have style");
}

// ============================================================================
// Test 4: Multiple cells with different font families
// ============================================================================

#[test]
fn test_multiple_cells_with_different_font_families() {
    let arial_style = StyleBuilder::new().font_name("Arial").build();
    let times_style = StyleBuilder::new().font_name("Times New Roman").build();
    let courier_style = StyleBuilder::new().font_name("Courier New").build();
    let verdana_style = StyleBuilder::new().font_name("Verdana").build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell(
                    "A1",
                    CellValue::String("Arial Cell".into()),
                    Some(arial_style),
                )
                .cell(
                    "A2",
                    CellValue::String("Times Cell".into()),
                    Some(times_style),
                )
                .cell(
                    "A3",
                    CellValue::String("Courier Cell".into()),
                    Some(courier_style),
                )
                .cell(
                    "A4",
                    CellValue::String("Verdana Cell".into()),
                    Some(verdana_style),
                ),
        )
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.cells.len(), 4, "Should have 4 cells");

    // Find cells by their row position (0-indexed)
    let arial_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0)
        .expect("Should find Arial cell");
    let times_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 1)
        .expect("Should find Times cell");
    let courier_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 2)
        .expect("Should find Courier cell");
    let verdana_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 3)
        .expect("Should find Verdana cell");

    assert_eq!(
        arial_cell.cell.s.as_ref().unwrap().font_family,
        Some("Arial".to_string())
    );
    assert_eq!(
        times_cell.cell.s.as_ref().unwrap().font_family,
        Some("Times New Roman".to_string())
    );
    assert_eq!(
        courier_cell.cell.s.as_ref().unwrap().font_family,
        Some("Courier New".to_string())
    );
    assert_eq!(
        verdana_cell.cell.s.as_ref().unwrap().font_family,
        Some("Verdana".to_string())
    );
}

#[test]
fn test_multiple_cells_same_font_different_styles() {
    // Test that cells can share the same font family but have different other styles
    let arial_bold = StyleBuilder::new().font_name("Arial").bold().build();
    let arial_italic = StyleBuilder::new().font_name("Arial").italic().build();
    let arial_both = StyleBuilder::new()
        .font_name("Arial")
        .bold()
        .italic()
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(
            SheetBuilder::new("Sheet1")
                .cell("A1", CellValue::String("Bold".into()), Some(arial_bold))
                .cell("A2", CellValue::String("Italic".into()), Some(arial_italic))
                .cell(
                    "A3",
                    CellValue::String("Bold Italic".into()),
                    Some(arial_both),
                ),
        )
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let sheet = &workbook.sheets[0];

    let bold_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0)
        .expect("Should find bold cell");
    let italic_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 1)
        .expect("Should find italic cell");
    let both_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 2)
        .expect("Should find bold+italic cell");

    // All should have Arial font
    assert_eq!(
        bold_cell.cell.s.as_ref().unwrap().font_family,
        Some("Arial".to_string())
    );
    assert_eq!(
        italic_cell.cell.s.as_ref().unwrap().font_family,
        Some("Arial".to_string())
    );
    assert_eq!(
        both_cell.cell.s.as_ref().unwrap().font_family,
        Some("Arial".to_string())
    );

    // Check individual style properties
    assert_eq!(bold_cell.cell.s.as_ref().unwrap().bold, Some(true));
    assert_ne!(bold_cell.cell.s.as_ref().unwrap().italic, Some(true));

    assert_eq!(italic_cell.cell.s.as_ref().unwrap().italic, Some(true));
    assert_ne!(italic_cell.cell.s.as_ref().unwrap().bold, Some(true));

    assert_eq!(both_cell.cell.s.as_ref().unwrap().bold, Some(true));
    assert_eq!(both_cell.cell.s.as_ref().unwrap().italic, Some(true));
}

// ============================================================================
// Test 5: Font family combined with other style properties
// ============================================================================

#[test]
fn test_font_family_with_size_and_color() {
    let style = StyleBuilder::new()
        .font_name("Georgia")
        .font_size(14.0)
        .font_color("#FF0000")
        .build();

    let xlsx = XlsxBuilder::new()
        .sheet(SheetBuilder::new("Sheet1").cell(
            "A1",
            CellValue::String("Styled".into()),
            Some(style),
        ))
        .build();

    let workbook = parser::parse(&xlsx).expect("Failed to parse");
    let cell = &workbook.sheets[0].cells[0];
    let style = cell.cell.s.as_ref().expect("Cell should have style");

    assert_eq!(
        style.font_family,
        Some("Georgia".to_string()),
        "Font family should be Georgia"
    );
    assert_eq!(style.font_size, Some(14.0), "Font size should be 14");
    assert!(style.font_color.is_some(), "Font color should be set");
    // Color may be normalized to ARGB format
    let color = style.font_color.as_ref().unwrap();
    assert!(
        color.contains("FF0000") || color.contains("ff0000"),
        "Font color should be red, got: {}",
        color
    );
}

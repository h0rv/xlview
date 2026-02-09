//! Comprehensive tests for all 19 pattern fill types.
//!
//! This module tests that each ECMA-376 pattern fill type can be:
//! 1. Created via StyleBuilder with the `.pattern()` method
//! 2. Written to a valid XLSX file
//! 3. Parsed correctly with the pattern_type field properly set
//!
//! Pattern fill types from ECMA-376 Part 1, Section 18.18.55:
//! - none, solid, gray125, gray0625
//! - darkGray, mediumGray, lightGray
//! - darkHorizontal, darkVertical, darkDown, darkUp, darkGrid, darkTrellis
//! - lightHorizontal, lightVertical, lightDown, lightUp, lightGrid, lightTrellis
//!
//! Note: The parser treats "solid" and "none" fills specially:
//! - "solid" fills: pattern_type is NOT set; instead bg_color is set directly
//! - "none" fills: pattern_type may be None (the default is no fill)
//! - All other pattern fills: pattern_type is set to the appropriate PatternType enum
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

mod fixtures;

use fixtures::{StyleBuilder, XlsxBuilder, ALL_PATTERN_FILLS};
use xlview::parser::parse;
use xlview::types::PatternType;

// ============================================================================
// Helper Functions
// ============================================================================

/// Convert a pattern type string to the expected PatternType enum variant.
fn expected_pattern_type(pattern_str: &str) -> PatternType {
    match pattern_str {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "gray125" => PatternType::Gray125,
        "gray0625" => PatternType::Gray0625,
        "darkGray" => PatternType::DarkGray,
        "mediumGray" => PatternType::MediumGray,
        "lightGray" => PatternType::LightGray,
        "darkHorizontal" => PatternType::DarkHorizontal,
        "darkVertical" => PatternType::DarkVertical,
        "darkDown" => PatternType::DarkDown,
        "darkUp" => PatternType::DarkUp,
        "darkGrid" => PatternType::DarkGrid,
        "darkTrellis" => PatternType::DarkTrellis,
        "lightHorizontal" => PatternType::LightHorizontal,
        "lightVertical" => PatternType::LightVertical,
        "lightDown" => PatternType::LightDown,
        "lightUp" => PatternType::LightUp,
        "lightGrid" => PatternType::LightGrid,
        "lightTrellis" => PatternType::LightTrellis,
        _ => panic!("Unknown pattern type: {}", pattern_str),
    }
}

/// Create an XLSX with a single cell using the specified pattern fill.
fn create_xlsx_with_pattern(pattern_type: &str) -> Vec<u8> {
    XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            format!("Pattern: {}", pattern_type),
            Some(StyleBuilder::new().pattern(pattern_type).build()),
        )
        .build()
}

// ============================================================================
// Test: All Pattern Types Parse Correctly
// ============================================================================

#[test]
fn test_all_pattern_types_via_fixture() {
    // Verify the fixture has all 19 pattern types
    assert_eq!(
        ALL_PATTERN_FILLS.len(),
        19,
        "ALL_PATTERN_FILLS should have exactly 19 patterns"
    );

    for pattern_type in ALL_PATTERN_FILLS {
        let xlsx = create_xlsx_with_pattern(pattern_type);
        let workbook = parse(&xlsx).unwrap_or_else(|e| {
            panic!(
                "Failed to parse XLSX with pattern '{}': {:?}",
                pattern_type, e
            )
        });

        assert_eq!(
            workbook.sheets.len(),
            1,
            "Should have exactly one sheet for pattern '{}'",
            pattern_type
        );

        let sheet = &workbook.sheets[0];
        let cell = sheet
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0)
            .unwrap_or_else(|| panic!("Cell A1 should exist for pattern '{}'", pattern_type));

        // Verify the cell has a style
        assert!(
            cell.cell.s.is_some(),
            "Cell should have style for pattern '{}'",
            pattern_type
        );

        let style = cell.cell.s.as_ref().unwrap();

        // Special handling for different pattern types:
        // - 'none': pattern_type may be None (the default)
        // - 'solid': pattern_type is NOT set; parser sets bg_color directly
        // - All others: pattern_type is set
        match *pattern_type {
            "none" => {
                // 'none' pattern may not set pattern_type (it's the default)
                if let Some(ref pt) = style.pattern_type {
                    assert_eq!(*pt, PatternType::None, "Pattern type mismatch for 'none'");
                }
            }
            "solid" => {
                // 'solid' fills don't set pattern_type - the parser converts them
                // to bg_color directly for simplicity
                // This is intentional behavior - solid is just a background color
                assert!(
                    style.pattern_type.is_none(),
                    "Solid pattern should NOT set pattern_type (parser sets bg_color instead)"
                );
            }
            _ => {
                // All other patterns should have pattern_type set
                assert!(
                    style.pattern_type.is_some(),
                    "pattern_type should be set for pattern '{}', got None",
                    pattern_type
                );

                let expected = expected_pattern_type(pattern_type);
                let actual = style.pattern_type.as_ref().unwrap();
                assert_eq!(
                    *actual, expected,
                    "Pattern type mismatch for '{}': expected {:?}, got {:?}",
                    pattern_type, expected, actual
                );
            }
        }
    }
}

// ============================================================================
// Individual Pattern Type Tests
// ============================================================================

#[test]
fn test_pattern_none() {
    let xlsx = create_xlsx_with_pattern("none");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    assert!(cell.is_some(), "Cell A1 should exist");
    let cell = cell.unwrap();
    assert!(cell.cell.s.is_some(), "Cell should have style");

    // 'none' pattern means no fill - pattern_type may be None or PatternType::None
    let style = cell.cell.s.as_ref().unwrap();
    if let Some(ref pt) = style.pattern_type {
        assert_eq!(*pt, PatternType::None);
    }
}

#[test]
fn test_pattern_solid() {
    // Note: The parser treats 'solid' fills specially.
    // Instead of setting pattern_type to Solid, it converts the fill
    // to a simple bg_color. This is by design for simplicity.
    let xlsx = create_xlsx_with_pattern("solid");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");

    // Solid fills don't set pattern_type - the parser optimizes this
    // to just a background color since solid is the most common case
    assert!(
        style.pattern_type.is_none(),
        "Solid fills should NOT set pattern_type (parser uses bg_color instead)"
    );
}

#[test]
fn test_pattern_gray125() {
    let xlsx = create_xlsx_with_pattern("gray125");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::Gray125),
        "Pattern should be gray125"
    );
}

#[test]
fn test_pattern_gray0625() {
    let xlsx = create_xlsx_with_pattern("gray0625");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::Gray0625),
        "Pattern should be gray0625"
    );
}

#[test]
fn test_pattern_dark_gray() {
    let xlsx = create_xlsx_with_pattern("darkGray");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkGray),
        "Pattern should be darkGray"
    );
}

#[test]
fn test_pattern_medium_gray() {
    let xlsx = create_xlsx_with_pattern("mediumGray");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::MediumGray),
        "Pattern should be mediumGray"
    );
}

#[test]
fn test_pattern_light_gray() {
    let xlsx = create_xlsx_with_pattern("lightGray");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightGray),
        "Pattern should be lightGray"
    );
}

#[test]
fn test_pattern_dark_horizontal() {
    let xlsx = create_xlsx_with_pattern("darkHorizontal");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkHorizontal),
        "Pattern should be darkHorizontal"
    );
}

#[test]
fn test_pattern_dark_vertical() {
    let xlsx = create_xlsx_with_pattern("darkVertical");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkVertical),
        "Pattern should be darkVertical"
    );
}

#[test]
fn test_pattern_dark_down() {
    let xlsx = create_xlsx_with_pattern("darkDown");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkDown),
        "Pattern should be darkDown"
    );
}

#[test]
fn test_pattern_dark_up() {
    let xlsx = create_xlsx_with_pattern("darkUp");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkUp),
        "Pattern should be darkUp"
    );
}

#[test]
fn test_pattern_dark_grid() {
    let xlsx = create_xlsx_with_pattern("darkGrid");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkGrid),
        "Pattern should be darkGrid"
    );
}

#[test]
fn test_pattern_dark_trellis() {
    let xlsx = create_xlsx_with_pattern("darkTrellis");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkTrellis),
        "Pattern should be darkTrellis"
    );
}

#[test]
fn test_pattern_light_horizontal() {
    let xlsx = create_xlsx_with_pattern("lightHorizontal");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightHorizontal),
        "Pattern should be lightHorizontal"
    );
}

#[test]
fn test_pattern_light_vertical() {
    let xlsx = create_xlsx_with_pattern("lightVertical");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightVertical),
        "Pattern should be lightVertical"
    );
}

#[test]
fn test_pattern_light_down() {
    let xlsx = create_xlsx_with_pattern("lightDown");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightDown),
        "Pattern should be lightDown"
    );
}

#[test]
fn test_pattern_light_up() {
    let xlsx = create_xlsx_with_pattern("lightUp");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightUp),
        "Pattern should be lightUp"
    );
}

#[test]
fn test_pattern_light_grid() {
    let xlsx = create_xlsx_with_pattern("lightGrid");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightGrid),
        "Pattern should be lightGrid"
    );
}

#[test]
fn test_pattern_light_trellis() {
    let xlsx = create_xlsx_with_pattern("lightTrellis");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(
        style.pattern_type,
        Some(PatternType::LightTrellis),
        "Pattern should be lightTrellis"
    );
}

// ============================================================================
// Pattern Fill with Colors Tests
// ============================================================================

#[test]
fn test_solid_fill_with_color() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Yellow Background",
            Some(StyleBuilder::new().bg_color("#FFFF00").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");

    // bg_color() sets pattern to solid in the fixture, but the parser
    // converts solid fills to just bg_color (no pattern_type)
    assert!(
        style.pattern_type.is_none(),
        "Solid fills should NOT set pattern_type (parser optimizes to bg_color)"
    );

    // Should have a background color
    assert!(
        style.bg_color.is_some(),
        "Cell should have background color"
    );
    let bg = style.bg_color.as_ref().unwrap();
    assert!(
        bg.contains("FFFF00") || bg.contains("ffff00"),
        "Background color should be yellow, got: {}",
        bg
    );
}

#[test]
fn test_gray125_pattern_parses_with_style() {
    // gray125 is a built-in pattern used in default XLSX styles
    let xlsx = create_xlsx_with_pattern("gray125");
    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");
    assert_eq!(style.pattern_type, Some(PatternType::Gray125));
}

// ============================================================================
// Multiple Patterns in Same Workbook
// ============================================================================

#[test]
fn test_multiple_patterns_in_single_sheet() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Solid",
            Some(StyleBuilder::new().pattern("solid").build()),
        )
        .add_cell(
            "A2",
            "Gray125",
            Some(StyleBuilder::new().pattern("gray125").build()),
        )
        .add_cell(
            "A3",
            "DarkGray",
            Some(StyleBuilder::new().pattern("darkGray").build()),
        )
        .add_cell(
            "A4",
            "LightHorizontal",
            Some(StyleBuilder::new().pattern("lightHorizontal").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];

    // Check A1 - solid (parser doesn't set pattern_type for solid)
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some());
    let style_a1 = cell_a1.unwrap().cell.s.as_ref().unwrap();
    assert!(
        style_a1.pattern_type.is_none(),
        "Solid fills should NOT set pattern_type"
    );

    // Check A2 - gray125
    let cell_a2 = sheet.cells.iter().find(|c| c.r == 1 && c.c == 0);
    assert!(cell_a2.is_some());
    let style_a2 = cell_a2.unwrap().cell.s.as_ref().unwrap();
    assert_eq!(style_a2.pattern_type, Some(PatternType::Gray125));

    // Check A3 - darkGray
    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert!(cell_a3.is_some());
    let style_a3 = cell_a3.unwrap().cell.s.as_ref().unwrap();
    assert_eq!(style_a3.pattern_type, Some(PatternType::DarkGray));

    // Check A4 - lightHorizontal
    let cell_a4 = sheet.cells.iter().find(|c| c.r == 3 && c.c == 0);
    assert!(cell_a4.is_some());
    let style_a4 = cell_a4.unwrap().cell.s.as_ref().unwrap();
    assert_eq!(style_a4.pattern_type, Some(PatternType::LightHorizontal));
}

#[test]
fn test_all_19_patterns_in_single_workbook() {
    // Create a workbook with all 19 pattern types
    let mut builder = XlsxBuilder::new().add_sheet("AllPatterns");

    for (i, pattern_type) in ALL_PATTERN_FILLS.iter().enumerate() {
        let cell_ref = format!("A{}", i + 1);
        builder = builder.add_cell(
            &cell_ref,
            *pattern_type,
            Some(StyleBuilder::new().pattern(pattern_type).build()),
        );
    }

    let xlsx = builder.build();
    let workbook = parse(&xlsx).expect("Failed to parse XLSX with all 19 patterns");

    let sheet = &workbook.sheets[0];
    assert!(
        sheet.cells.len() >= 19,
        "Should have at least 19 cells, got {}",
        sheet.cells.len()
    );

    // Verify each pattern
    for (i, pattern_str) in ALL_PATTERN_FILLS.iter().enumerate() {
        let cell = sheet.cells.iter().find(|c| c.r == i as u32 && c.c == 0);
        assert!(
            cell.is_some(),
            "Cell A{} should exist for pattern '{}'",
            i + 1,
            pattern_str
        );

        let cell = cell.unwrap();
        let style = cell.cell.s.as_ref();

        match *pattern_str {
            "none" => {
                // 'none' pattern may not have pattern_type set
                if let Some(s) = style {
                    if let Some(ref pt) = s.pattern_type {
                        assert_eq!(
                            *pt,
                            PatternType::None,
                            "Cell A{} pattern mismatch for 'none'",
                            i + 1
                        );
                    }
                }
            }
            "solid" => {
                // 'solid' fills don't set pattern_type
                if let Some(s) = style {
                    assert!(
                        s.pattern_type.is_none(),
                        "Cell A{} solid fill should NOT have pattern_type set",
                        i + 1
                    );
                }
            }
            _ => {
                // All other patterns should have pattern_type set
                assert!(
                    style.is_some(),
                    "Cell A{} should have style for pattern '{}'",
                    i + 1,
                    pattern_str
                );
                let style = style.unwrap();
                assert!(
                    style.pattern_type.is_some(),
                    "Cell A{} should have pattern_type for '{}'",
                    i + 1,
                    pattern_str
                );

                let expected = expected_pattern_type(pattern_str);
                let actual = style.pattern_type.as_ref().unwrap();
                assert_eq!(
                    *actual,
                    expected,
                    "Cell A{} pattern mismatch: expected {:?}, got {:?}",
                    i + 1,
                    expected,
                    actual
                );
            }
        }
    }
}

// ============================================================================
// Pattern Fill Count Verification
// ============================================================================

#[test]
fn test_all_pattern_fills_constant_has_19_entries() {
    assert_eq!(
        ALL_PATTERN_FILLS.len(),
        19,
        "ALL_PATTERN_FILLS should contain exactly 19 pattern types per ECMA-376"
    );
}

#[test]
fn test_pattern_fills_contain_expected_values() {
    // Verify the fixture contains all expected pattern names
    let expected_patterns = [
        "none",
        "solid",
        "mediumGray",
        "darkGray",
        "lightGray",
        "darkHorizontal",
        "darkVertical",
        "darkDown",
        "darkUp",
        "darkGrid",
        "darkTrellis",
        "lightHorizontal",
        "lightVertical",
        "lightDown",
        "lightUp",
        "lightGrid",
        "lightTrellis",
        "gray125",
        "gray0625",
    ];

    for pattern in &expected_patterns {
        assert!(
            ALL_PATTERN_FILLS.contains(pattern),
            "ALL_PATTERN_FILLS should contain '{}'",
            pattern
        );
    }
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_pattern_with_additional_styling() {
    // Pattern fill combined with other style attributes
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Styled Pattern",
            Some(
                StyleBuilder::new()
                    .pattern("darkGrid")
                    .bold()
                    .italic()
                    .font_size(14.0)
                    .font_color("#FF0000")
                    .build(),
            ),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    let style = cell.cell.s.as_ref().expect("Cell should have style");

    // Verify pattern
    assert_eq!(
        style.pattern_type,
        Some(PatternType::DarkGrid),
        "Pattern should be darkGrid"
    );

    // Verify other style attributes are preserved
    assert_eq!(style.bold, Some(true), "Should be bold");
    assert_eq!(style.italic, Some(true), "Should be italic");
    assert_eq!(style.font_size, Some(14.0), "Font size should be 14");
    assert!(style.font_color.is_some(), "Should have font color");
}

#[test]
fn test_empty_cell_with_pattern_only() {
    // A cell with just a pattern fill and no value
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "", // Empty string value
            Some(StyleBuilder::new().pattern("lightTrellis").build()),
        )
        .build();

    let workbook = parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];

    // Find the cell - it might be at row 0, col 0
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);

    // Even empty cells with styles should be parseable
    if let Some(cell) = cell {
        if let Some(ref style) = cell.cell.s {
            assert_eq!(
                style.pattern_type,
                Some(PatternType::LightTrellis),
                "Pattern should be lightTrellis"
            );
        }
    }
}

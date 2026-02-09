//! Fuzz Validation Tests
//!
//! Comprehensive validation of Excel features through generated test files
//! with randomized parameters to catch edge cases.
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
use fixtures::{SheetBuilder, StyleBuilder, XlsxBuilder};
use xlview::parser::parse;

// ============================================================================
// Style Fuzzing
// ============================================================================

/// Fuzz all combinations of font styles
#[test]
fn fuzz_font_style_combinations() {
    let combinations = [
        (true, false, false, false),  // bold only
        (false, true, false, false),  // italic only
        (false, false, true, false),  // underline only
        (false, false, false, true),  // strikethrough only
        (true, true, false, false),   // bold + italic
        (true, true, true, false),    // bold + italic + underline
        (true, true, true, true),     // all styles
        (false, false, false, false), // no styles
    ];

    for (bold, italic, underline, strike) in combinations {
        let mut style = StyleBuilder::new();
        if bold {
            style = style.bold();
        }
        if italic {
            style = style.italic();
        }
        if underline {
            style = style.underline();
        }
        if strike {
            style = style.strikethrough();
        }

        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell("A1", "Styled Text", Some(style.build()))
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed with bold={}, italic={}, underline={}, strike={}: {:?}",
            bold,
            italic,
            underline,
            strike,
            result.err()
        );

        let workbook = result.unwrap();
        let cell = workbook.sheets[0]
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0);
        assert!(cell.is_some(), "Cell A1 should exist");

        if bold || italic || underline || strike {
            let style = cell.unwrap().cell.s.as_ref();
            assert!(style.is_some(), "Cell should have style");

            let s = style.unwrap();
            if bold {
                assert_eq!(s.bold, Some(true), "Bold should be set");
            }
            if italic {
                assert_eq!(s.italic, Some(true), "Italic should be set");
            }
        }
    }
}

/// Fuzz font sizes from very small to very large
#[test]
fn fuzz_font_sizes() {
    let sizes: Vec<f64> = vec![
        1.0, 6.0, 8.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0, 28.0, 32.0, 36.0, 48.0,
        72.0, 96.0, 144.0, 200.0, 409.0,
    ];

    for size in sizes {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Text",
                Some(StyleBuilder::new().font_size(size).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse font size {}: {:?}",
            size,
            result.err()
        );

        let workbook = result.unwrap();
        let cell = workbook.sheets[0]
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0)
            .unwrap();
        let parsed_size = cell.cell.s.as_ref().and_then(|s| s.font_size);
        assert_eq!(
            parsed_size,
            Some(size),
            "Font size should be {} but got {:?}",
            size,
            parsed_size
        );
    }
}

/// Fuzz RGB colors across the spectrum
#[test]
fn fuzz_rgb_colors() {
    let colors = [
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#808080", "#C0C0C0", "#800000", "#008000", "#000080", "#808000", "#800080", "#008080",
        "#123456", "#ABCDEF", "#FEDCBA", "#112233", "#AABBCC", "#001122",
    ];

    for color in colors {
        // Test as font color
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Colored",
                Some(StyleBuilder::new().font_color(color).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse font color {}: {:?}",
            color,
            result.err()
        );

        // Test as background color
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Filled",
                Some(StyleBuilder::new().bg_color(color).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse bg color {}: {:?}",
            color,
            result.err()
        );
    }
}

/// Fuzz all border styles on all sides
#[test]
fn fuzz_border_styles() {
    let styles = [
        "thin",
        "medium",
        "thick",
        "dashed",
        "dotted",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "slantDashDot",
    ];

    for style in styles {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Bordered",
                Some(StyleBuilder::new().border_all(style, None).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse border style {}: {:?}",
            style,
            result.err()
        );

        let workbook = result.unwrap();
        let cell = workbook.sheets[0]
            .cells
            .iter()
            .find(|c| c.r == 0 && c.c == 0)
            .unwrap();
        let s = cell.cell.s.as_ref().expect("Should have style");

        assert!(
            s.border_top.is_some(),
            "Should have top border for {}",
            style
        );
        assert!(
            s.border_bottom.is_some(),
            "Should have bottom border for {}",
            style
        );
        assert!(
            s.border_left.is_some(),
            "Should have left border for {}",
            style
        );
        assert!(
            s.border_right.is_some(),
            "Should have right border for {}",
            style
        );
    }
}

/// Fuzz all fill patterns
#[test]
fn fuzz_fill_patterns() {
    let patterns = [
        "solid",
        "gray125",
        "gray0625",
        "darkGray",
        "mediumGray",
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
    ];

    for pattern in patterns {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Pattern",
                Some(
                    StyleBuilder::new()
                        .pattern(pattern)
                        .bg_color("#FF0000")
                        .build(),
                ),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse pattern {}: {:?}",
            pattern,
            result.err()
        );
    }
}

// ============================================================================
// Alignment Fuzzing
// ============================================================================

/// Fuzz all alignment combinations
#[test]
fn fuzz_alignment_combinations() {
    let horizontal = ["left", "center", "right", "justify", "fill"];
    let vertical = ["top", "center", "bottom"];

    for h in horizontal {
        for v in vertical {
            let xlsx = XlsxBuilder::new()
                .add_sheet("Sheet1")
                .add_cell(
                    "A1",
                    "Aligned",
                    Some(
                        StyleBuilder::new()
                            .align_horizontal(h)
                            .align_vertical(v)
                            .build(),
                    ),
                )
                .build();

            let result = parse(&xlsx);
            assert!(
                result.is_ok(),
                "Failed with h={}, v={}: {:?}",
                h,
                v,
                result.err()
            );
        }
    }
}

/// Fuzz text rotation angles
#[test]
fn fuzz_text_rotation() {
    for angle in (0..=90).step_by(15) {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Rotated",
                Some(StyleBuilder::new().rotation(angle).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse rotation {}: {:?}",
            angle,
            result.err()
        );
    }
}

/// Fuzz text indent levels
#[test]
fn fuzz_text_indent() {
    for indent in 0..=15 {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "Indented",
                Some(StyleBuilder::new().indent(indent).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse indent {}: {:?}",
            indent,
            result.err()
        );
    }
}

// ============================================================================
// Cell Content Fuzzing
// ============================================================================

/// Fuzz various string content types
#[test]
fn fuzz_string_content() {
    let strings = [
        "",                    // empty
        " ",                   // single space
        "Hello",               // simple
        "Hello World",         // with space
        "Hello\nWorld",        // with newline
        "Hello\tWorld",        // with tab
        "Привет мир",          // Cyrillic
        "你好世界",            // Chinese
        "🎉🎊🎈",              // Emoji
        "Line1\nLine2\nLine3", // multiple lines
        "<>&\"'",              // XML special chars
        "=SUM(A1:A10)",        // formula-like
        "123",                 // number-like
        "TRUE",                // boolean-like
        "#N/A",                // error-like
    ];

    for (i, s) in strings.iter().enumerate() {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell("A1", *s, None)
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse string {}: {:?}",
            i,
            result.err()
        );
    }
}

/// Fuzz numeric values including edge cases
#[test]
fn fuzz_numeric_values() {
    let numbers = [
        "0",
        "1",
        "-1",
        "0.5",
        "-0.5",
        "0.0",
        "1.0",
        "100",
        "1000",
        "1000000",
        "1.23456789",
        "0.000001",
        "999999999999",
        "-999999999999",
        "1E10",
        "1E-10",
        "1.5E5",
        "-1.5E-5",
    ];

    for num in numbers {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell("A1", num, None)
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse number {}: {:?}",
            num,
            result.err()
        );
    }
}

/// Fuzz long strings
#[test]
fn fuzz_long_strings() {
    let lengths = [100, 500, 1000, 5000];

    for len in lengths {
        let long_string = "A".repeat(len);
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell("A1", long_string, None)
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse string of length {}: {:?}",
            len,
            result.err()
        );
    }
}

// ============================================================================
// Sheet Structure Fuzzing
// ============================================================================

/// Fuzz various merge configurations
#[test]
fn fuzz_merge_configurations() {
    let merges = [
        "A1:B1",  // horizontal 2
        "A1:D1",  // horizontal 4
        "A1:A4",  // vertical 4
        "A1:B2",  // 2x2
        "A1:C3",  // 3x3
        "A1:E5",  // 5x5
        "B2:D4",  // offset 3x3
        "A1:J10", // 10x10
    ];

    for merge_range in merges {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell("A1", "Merged", None)
            .add_merge(merge_range)
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse merge {}: {:?}",
            merge_range,
            result.err()
        );

        let workbook = result.unwrap();
        assert!(
            !workbook.sheets[0].merges.is_empty(),
            "Should have merges for {}",
            merge_range
        );
    }
}

/// Fuzz multiple sheets
#[test]
fn fuzz_multiple_sheets() {
    for sheet_count in [1, 2, 3, 5, 10] {
        let mut builder = XlsxBuilder::new();

        for i in 0..sheet_count {
            let sheet = SheetBuilder::new(&format!("Sheet{}", i + 1)).cell(
                "A1",
                format!("Content {}", i + 1),
                None,
            );
            builder = builder.sheet(sheet);
        }

        let xlsx = builder.build();
        let result = parse(&xlsx);

        assert!(
            result.is_ok(),
            "Failed to parse {} sheets: {:?}",
            sheet_count,
            result.err()
        );

        let workbook = result.unwrap();
        assert_eq!(
            workbook.sheets.len(),
            sheet_count,
            "Should have {} sheets",
            sheet_count
        );
    }
}

/// Fuzz frozen pane configurations
#[test]
fn fuzz_frozen_panes() {
    let configs = [
        (1, 0), // freeze first row
        (0, 1), // freeze first column
        (1, 1), // freeze first row and column
        (2, 0), // freeze two rows
        (0, 2), // freeze two columns
        (2, 2), // freeze 2x2
    ];

    for (rows, cols) in configs {
        let sheet = SheetBuilder::new("Sheet1")
            .cell("A1", "Frozen", None)
            .freeze_panes(rows, cols);

        let xlsx = XlsxBuilder::new().sheet(sheet).build();
        let result = parse(&xlsx);

        assert!(
            result.is_ok(),
            "Failed to parse frozen panes ({}, {}): {:?}",
            rows,
            cols,
            result.err()
        );

        let workbook = result.unwrap();
        assert_eq!(
            workbook.sheets[0].frozen_rows, rows,
            "Frozen rows should be {}",
            rows
        );
        assert_eq!(
            workbook.sheets[0].frozen_cols, cols,
            "Frozen cols should be {}",
            cols
        );
    }
}

/// Fuzz column widths
#[test]
fn fuzz_column_widths() {
    let widths: Vec<f64> = vec![1.0, 8.43, 10.0, 20.0, 50.0, 100.0, 255.0];

    for width in &widths {
        let sheet = SheetBuilder::new("Sheet1")
            .cell("A1", "Wide", None)
            .col_width(1, 1, *width);

        let xlsx = XlsxBuilder::new().sheet(sheet).build();
        let result = parse(&xlsx);

        assert!(
            result.is_ok(),
            "Failed to parse col width {}: {:?}",
            width,
            result.err()
        );
    }
}

/// Fuzz row heights
#[test]
fn fuzz_row_heights() {
    let heights: Vec<f64> = vec![1.0, 15.0, 20.0, 30.0, 50.0, 100.0, 409.0];

    for height in &heights {
        let sheet = SheetBuilder::new("Sheet1")
            .cell("A1", "Tall", None)
            .row_height(1, *height);

        let xlsx = XlsxBuilder::new().sheet(sheet).build();
        let result = parse(&xlsx);

        assert!(
            result.is_ok(),
            "Failed to parse row height {}: {:?}",
            height,
            result.err()
        );
    }
}

// ============================================================================
// Stress Tests
// ============================================================================

/// Stress test with many cells
#[test]
fn stress_many_cells() {
    let cell_counts = [100, 500, 1000];

    for count in cell_counts {
        let mut sheet = SheetBuilder::new("Sheet1");

        for i in 0..count {
            let col = (i % 26) as u8;
            let row = i / 26;
            let cell_ref = format!("{}{}", (b'A' + col) as char, row + 1);
            sheet = sheet.cell(&cell_ref, format!("Cell {}", i), None);
        }

        let xlsx = XlsxBuilder::new().sheet(sheet).build();
        let result = parse(&xlsx);

        assert!(
            result.is_ok(),
            "Failed to parse {} cells: {:?}",
            count,
            result.err()
        );

        let workbook = result.unwrap();
        assert!(
            workbook.sheets[0].cells.len() >= count,
            "Should have at least {} cells",
            count
        );
    }
}

/// Stress test with many styled cells
#[test]
fn stress_many_styled_cells() {
    let mut sheet = SheetBuilder::new("Sheet1");
    let colors = ["#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF"];

    for i in 0..200 {
        let col = (i % 26) as u8;
        let row = i / 26;
        let cell_ref = format!("{}{}", (b'A' + col) as char, row + 1);
        let color = colors[i % colors.len()];

        sheet = sheet.cell(
            &cell_ref,
            format!("Cell {}", i),
            Some(
                StyleBuilder::new()
                    .bold()
                    .bg_color(color)
                    .border_all("thin", None),
            ),
        );
    }

    let xlsx = XlsxBuilder::new().sheet(sheet).build();
    let result = parse(&xlsx);

    assert!(
        result.is_ok(),
        "Failed to parse many styled cells: {:?}",
        result.err()
    );
}

/// Stress test with many merges
#[test]
fn stress_many_merges() {
    let mut sheet = SheetBuilder::new("Sheet1");

    // Create a grid of 2x2 merges
    for row in (0..20).step_by(2) {
        for col in (0..10).step_by(2) {
            let start_col = (b'A' + col) as char;
            let end_col = (b'A' + col + 1) as char;
            let start = format!("{}{}", start_col, row + 1);
            let range = format!("{}{}:{}{}", start_col, row + 1, end_col, row + 2);

            sheet = sheet
                .cell(&start, format!("Merge {},{}", row, col), None)
                .merge(&range);
        }
    }

    let xlsx = XlsxBuilder::new().sheet(sheet).build();
    let result = parse(&xlsx);

    assert!(
        result.is_ok(),
        "Failed to parse many merges: {:?}",
        result.err()
    );

    let workbook = result.unwrap();
    assert!(
        workbook.sheets[0].merges.len() >= 50,
        "Should have at least 50 merges"
    );
}

/// Stress test combining multiple features
#[test]
fn stress_combined_features() {
    let mut sheet = SheetBuilder::new("Sheet1").freeze_panes(2, 1);

    // Add header row with bold (columns A-J)
    for col in 0u8..10 {
        let cell_ref = format!("{}{}", (b'A' + col) as char, 1);
        sheet = sheet.cell(
            &cell_ref,
            format!("Header {}", col + 1),
            Some(StyleBuilder::new().bold().bg_color("#CCCCCC")),
        );
    }

    // Add data rows with alternating colors
    for row in 1usize..50 {
        for col in 0u8..10 {
            let cell_ref = format!("{}{}", (b'A' + col) as char, row + 1);
            let bg = if row % 2 == 0 { "#FFFFFF" } else { "#F0F0F0" };
            sheet = sheet.cell(
                &cell_ref,
                format!("{}", row * 10 + col as usize),
                Some(StyleBuilder::new().bg_color(bg)),
            );
        }
    }

    // Add a merge
    sheet = sheet
        .cell(
            "A52",
            "Merged Footer",
            Some(StyleBuilder::new().bold().align_horizontal("center")),
        )
        .merge("A52:E52");

    let xlsx = XlsxBuilder::new().sheet(sheet).build();
    let result = parse(&xlsx);

    assert!(
        result.is_ok(),
        "Failed to parse combined features: {:?}",
        result.err()
    );

    let workbook = result.unwrap();
    assert_eq!(workbook.sheets[0].frozen_rows, 2);
    assert_eq!(workbook.sheets[0].frozen_cols, 1);
    assert!(!workbook.sheets[0].merges.is_empty());
}

// ============================================================================
// Edge Cases
// ============================================================================

/// Test empty sheet
#[test]
fn edge_case_empty_sheet() {
    let xlsx = XlsxBuilder::new().sheet(SheetBuilder::new("Empty")).build();
    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse empty sheet: {:?}",
        result.err()
    );
}

/// Test very long sheet name
#[test]
fn edge_case_long_sheet_name() {
    let long_name = "A".repeat(31); // Excel max is 31 chars
    let sheet = SheetBuilder::new(&long_name).cell("A1", "Content", None);
    let xlsx = XlsxBuilder::new().sheet(sheet).build();

    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse long sheet name: {:?}",
        result.err()
    );
}

/// Test special characters in sheet name
#[test]
fn edge_case_special_sheet_name() {
    let names = [
        "Sheet 1",
        "Sheet-1",
        "Sheet_1",
        "Sheet.1",
        "Sheet(1)",
        "日本語",
        "Données",
    ];

    for name in names {
        let sheet = SheetBuilder::new(name).cell("A1", "Content", None);
        let xlsx = XlsxBuilder::new().sheet(sheet).build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse sheet name '{}': {:?}",
            name,
            result.err()
        );
    }
}

/// Test cell references at extremes
#[test]
fn edge_case_extreme_cell_refs() {
    let refs = ["A1", "Z1", "AA1", "AZ1", "BA1", "A100", "A1000", "Z100"];

    for cell_ref in refs {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(cell_ref, "Content", None)
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse cell ref '{}': {:?}",
            cell_ref,
            result.err()
        );
    }
}

/// Test number format strings
#[test]
fn fuzz_number_formats() {
    let formats = [
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "0%",
        "0.00%",
        "0.00E+00",
        "@",
        "yyyy-mm-dd",
        "hh:mm:ss",
        "mm/dd/yyyy",
        "$#,##0.00",
    ];

    for format in formats {
        let xlsx = XlsxBuilder::new()
            .add_sheet("Sheet1")
            .add_cell(
                "A1",
                "12345.67",
                Some(StyleBuilder::new().number_format(format).build()),
            )
            .build();

        let result = parse(&xlsx);
        assert!(
            result.is_ok(),
            "Failed to parse number format '{}': {:?}",
            format,
            result.err()
        );
    }
}

/// Test subscript and superscript
#[test]
fn fuzz_vert_align() {
    // Test subscript
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "H2O", Some(StyleBuilder::new().subscript().build()))
        .build();

    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse subscript: {:?}",
        result.err()
    );

    // Test superscript
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "x2", Some(StyleBuilder::new().superscript().build()))
        .build();

    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse superscript: {:?}",
        result.err()
    );
}

/// Test wrap text
#[test]
fn fuzz_wrap_text() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Long text that wraps",
            Some(StyleBuilder::new().wrap_text().build()),
        )
        .build();

    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse wrap text: {:?}",
        result.err()
    );

    let workbook = result.unwrap();
    let cell = workbook.sheets[0]
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .unwrap();
    let wrap = cell.cell.s.as_ref().and_then(|s| s.wrap);
    assert_eq!(wrap, Some(true), "Wrap should be enabled");
}

/// Test combined styles
#[test]
fn fuzz_combined_styles() {
    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell(
            "A1",
            "Fully Styled",
            Some(
                StyleBuilder::new()
                    .bold()
                    .italic()
                    .underline()
                    .font_size(14.0)
                    .font_color("#0000FF")
                    .bg_color("#FFFF00")
                    .border_all("medium", Some("#000000"))
                    .align_horizontal("center")
                    .align_vertical("center")
                    .wrap_text()
                    .build(),
            ),
        )
        .build();

    let result = parse(&xlsx);
    assert!(
        result.is_ok(),
        "Failed to parse combined styles: {:?}",
        result.err()
    );

    let workbook = result.unwrap();
    let cell = workbook.sheets[0]
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .unwrap();
    let style = cell.cell.s.as_ref().expect("Should have style");

    assert_eq!(style.bold, Some(true));
    assert_eq!(style.italic, Some(true));
    assert!(style.underline.is_some());
    assert_eq!(style.font_size, Some(14.0));
    assert!(style.font_color.is_some());
    assert!(style.bg_color.is_some());
    assert_eq!(style.wrap, Some(true));
}

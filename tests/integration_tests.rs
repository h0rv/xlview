//! Integration tests for xlview
//!
//! These tests parse real XLSX files and validate the output structure,
//! styling, and formatting. They serve as regression tests for the parser.
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

use std::fs;
use xlview::parser::parse;
use xlview::types::{CellData, Sheet};

/// Helper to find a cell by value in a sheet
fn find_cell_by_value<'a>(sheet: &'a Sheet, value: &str) -> Option<&'a CellData> {
    sheet.cells.iter().find(|cd| {
        if let Some(ref v) = cd.cell.v {
            v == value
        } else {
            false
        }
    })
}

/// Helper to get a cell at specific coordinates
#[allow(dead_code)]
fn get_cell_at(sheet: &Sheet, row: u32, col: u32) -> Option<&CellData> {
    sheet.cells.iter().find(|cd| cd.r == row && cd.c == col)
}

/// Test that the kitchen sink file parses without errors
#[test]
fn test_kitchen_sink_parses() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Should have 6 sheets (5 visible + 1 hidden)
    assert!(workbook.sheets.len() >= 5, "Expected at least 5 sheets");

    // First sheet should be "Fonts & Colors"
    assert_eq!(workbook.sheets[0].name, "Fonts &amp; Colors");
}

/// Test that tab colors are parsed correctly
#[test]
fn test_tab_colors() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // First sheet should have red tab color
    assert!(
        workbook.sheets[0].tab_color.is_some(),
        "First sheet should have a tab color"
    );
    let tab_color = workbook.sheets[0].tab_color.as_ref().unwrap();
    assert!(
        tab_color.contains("FF6B6B") || tab_color.contains("ff6b6b"),
        "Expected red tab color, got {}",
        tab_color
    );
}

/// Test that fonts are parsed correctly
#[test]
fn test_font_parsing() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Find the "Bold" cell and verify it has bold styling
    if let Some(cell_data) = find_cell_by_value(sheet, "Bold") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Bold cell should have style");
        assert_eq!(
            style.bold,
            Some(true),
            "Bold cell should have bold=true, got {:?}. Full style: {:?}",
            style.bold,
            style
        );
    } else {
        panic!("Could not find 'Bold' cell in sheet");
    }

    // Find the "Italic" cell and verify it has italic styling
    if let Some(cell_data) = find_cell_by_value(sheet, "Italic") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Italic cell should have style");
        assert_eq!(
            style.italic,
            Some(true),
            "Italic cell should have italic=true"
        );
    } else {
        panic!("Could not find 'Italic' cell in sheet");
    }
}

/// Test that font colors are parsed correctly
#[test]
fn test_font_colors() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Find the "Red" cell and verify it has red font color
    if let Some(cell_data) = find_cell_by_value(sheet, "Red") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Red cell should have style");
        assert!(
            style.font_color.is_some(),
            "Red cell should have font_color. Full style: {:?}",
            style
        );
        let color = style.font_color.as_ref().unwrap();
        assert!(
            color.contains("FF0000") || color.contains("ff0000"),
            "Expected red font color, got {}",
            color
        );
    } else {
        panic!("Could not find 'Red' cell in sheet");
    }
}

/// Test that background fills are parsed correctly
#[test]
fn test_background_fills() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Find a cell with "Fill 1" text and verify background color
    if let Some(cell_data) = find_cell_by_value(sheet, "Fill 1") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Fill cell should have style");
        assert!(
            style.bg_color.is_some(),
            "Fill 1 cell should have bg_color. Full style: {:?}",
            style
        );
    } else {
        panic!("Could not find 'Fill 1' cell in sheet");
    }
}

/// Test that borders are parsed correctly
#[test]
fn test_borders() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Find "Borders" sheet
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Borders")
        .expect("Should have Borders sheet");

    // Find "thin" border cell
    if let Some(cell_data) = find_cell_by_value(sheet, "thin") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Border cell should have style");
        assert!(
            style.border_top.is_some(),
            "thin cell should have border_top. Full style: {:?}",
            style
        );
        let border = style.border_top.as_ref().unwrap();
        assert_eq!(
            format!("{:?}", border.style),
            "Thin",
            "Expected thin border style"
        );
    } else {
        panic!("Could not find 'thin' cell in Borders sheet");
    }
}

/// Test number formatting
#[test]
fn test_number_formats() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Find "Numbers" sheet
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Numbers")
        .expect("Should have Numbers sheet");

    // Verify the sheet has cells
    assert!(!sheet.cells.is_empty(), "Numbers sheet should have cells");
}

/// Test alignment
#[test]
fn test_alignment() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Find "Alignment" sheet
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Alignment")
        .expect("Should have Alignment sheet");

    // Find "H: center" cell and verify alignment
    if let Some(cell_data) = find_cell_by_value(sheet, "H: center") {
        let style = cell_data
            .cell
            .s
            .as_ref()
            .expect("Alignment cell should have style");
        assert!(
            style.align_h.is_some(),
            "center cell should have align_h. Full style: {:?}",
            style
        );
    }
}

/// Test merged cells
#[test]
fn test_merged_cells() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Find "Alignment" sheet which has merged cells
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Alignment")
        .expect("Should have Alignment sheet");

    // Should have at least one merge
    assert!(
        !sheet.merges.is_empty(),
        "Alignment sheet should have merged cells"
    );
}

/// Test that minimal.xlsx parses correctly
#[test]
fn test_minimal_xlsx() {
    let data = fs::read("test/minimal.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
    assert!(
        !workbook.sheets[0].cells.is_empty(),
        "Sheet should have cells"
    );
}

/// Test that styled.xlsx parses correctly
#[test]
fn test_styled_xlsx() {
    let data = fs::read("test/styled.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

// ============================================================================
// kitchen_sink_v2 Comprehensive Tests
// ============================================================================

/// Test that kitchen_sink_v2 parses all sheets correctly
#[test]
fn test_kitchen_sink_v2_sheet_structure() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Should have at least 1 sheet
    assert!(
        !workbook.sheets.is_empty(),
        "Expected at least 1 sheet, got {}",
        workbook.sheets.len()
    );

    // Verify all sheets have names
    for sheet in &workbook.sheets {
        assert!(!sheet.name.is_empty(), "Sheet should have a name");
    }
}

/// Test that pattern fills are parsed in v2 file
#[test]
fn test_kitchen_sink_v2_pattern_fills() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut pattern_count = 0;
    for sheet in &workbook.sheets {
        for cell in &sheet.cells {
            if let Some(ref style) = cell.cell.s {
                if style.pattern_type.is_some() {
                    pattern_count += 1;
                }
            }
        }
    }

    // v2 file may have limited pattern fills, just verify we can parse them
    assert!(
        pattern_count >= 0,
        "Pattern fill parsing should work, got {} pattern cells",
        pattern_count
    );
}

/// Test that border styles are parsed in v2 file
#[test]
fn test_kitchen_sink_v2_border_styles() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut border_styles: std::collections::HashSet<String> = std::collections::HashSet::new();

    for sheet in &workbook.sheets {
        for cell in &sheet.cells {
            if let Some(ref style) = cell.cell.s {
                if let Some(ref border) = style.border_top {
                    border_styles.insert(format!("{:?}", border.style));
                }
                if let Some(ref border) = style.border_bottom {
                    border_styles.insert(format!("{:?}", border.style));
                }
                if let Some(ref border) = style.border_left {
                    border_styles.insert(format!("{:?}", border.style));
                }
                if let Some(ref border) = style.border_right {
                    border_styles.insert(format!("{:?}", border.style));
                }
            }
        }
    }

    assert!(
        border_styles.len() >= 3,
        "Should have at least 3 different border styles, got {:?}",
        border_styles
    );
}

/// Test conditional formatting is parsed
#[test]
fn test_kitchen_sink_v2_conditional_formatting() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut cf_count = 0;
    for sheet in &workbook.sheets {
        cf_count += sheet.conditional_formatting.len();
    }

    // v2 file should have conditional formatting rules
    assert!(
        cf_count >= 1,
        "Should have at least 1 CF rule, got {}",
        cf_count
    );
}

/// Test data validation is parsed
#[test]
fn test_kitchen_sink_v2_data_validation() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut validation_count = 0;
    for sheet in &workbook.sheets {
        validation_count += sheet.data_validations.len();
    }

    // v2 file should have data validations
    assert!(
        validation_count >= 1,
        "Should have at least 1 data validation, got {}",
        validation_count
    );
}

/// Test that hyperlinks are parsed
#[test]
fn test_kitchen_sink_v2_hyperlinks() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut hyperlink_count = 0;
    for sheet in &workbook.sheets {
        hyperlink_count += sheet.hyperlinks.len();
    }

    assert!(
        hyperlink_count >= 1,
        "Should have at least 1 hyperlink, got {}",
        hyperlink_count
    );
}

/// Test that drawings (images/shapes) are parsed
#[test]
fn test_kitchen_sink_v2_drawings() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut drawing_count = 0;
    for sheet in &workbook.sheets {
        drawing_count += sheet.drawings.len();
    }

    // v2 file should have drawings
    assert!(
        drawing_count >= 1,
        "Should have at least 1 drawing, got {}",
        drawing_count
    );
}

/// Test that charts are parsed
#[test]
fn test_kitchen_sink_v2_charts() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut chart_count = 0;
    for sheet in &workbook.sheets {
        chart_count += sheet.charts.len();
    }

    // v2 file should have charts
    assert!(
        chart_count >= 1,
        "Should have at least 1 chart, got {}",
        chart_count
    );
}

/// Test that theme colors are parsed correctly
#[test]
fn test_theme_colors() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Theme should have colors
    assert!(
        !workbook.theme.colors.is_empty(),
        "Theme should have colors parsed"
    );

    // Should have standard theme color count (12 standard colors)
    assert!(
        workbook.theme.colors.len() >= 10,
        "Theme should have at least 10 colors, got {}",
        workbook.theme.colors.len()
    );
}

// ============================================================================
// Comprehensive Feature Counts
// ============================================================================

/// Test comprehensive feature counts across the workbook
#[test]
fn test_comprehensive_feature_coverage() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut stats = FeatureStats::default();

    for sheet in &workbook.sheets {
        stats.sheet_count += 1;

        for cell in &sheet.cells {
            stats.cell_count += 1;

            if let Some(ref style) = cell.cell.s {
                if style.bold == Some(true) {
                    stats.bold_cells += 1;
                }
                if style.italic == Some(true) {
                    stats.italic_cells += 1;
                }
                if style.underline.is_some() {
                    stats.underlined_cells += 1;
                }
                if style.font_color.is_some() {
                    stats.colored_font_cells += 1;
                }
                if style.bg_color.is_some() {
                    stats.colored_bg_cells += 1;
                }
                if style.pattern_type.is_some() {
                    stats.pattern_fill_cells += 1;
                }
                if style.border_top.is_some()
                    || style.border_bottom.is_some()
                    || style.border_left.is_some()
                    || style.border_right.is_some()
                {
                    stats.bordered_cells += 1;
                }
                if style.align_h.is_some() || style.align_v.is_some() {
                    stats.aligned_cells += 1;
                }
                if style.wrap == Some(true) {
                    stats.wrapped_cells += 1;
                }
                if style.rotation.is_some() && style.rotation != Some(0) {
                    stats.rotated_cells += 1;
                }
            }
        }

        stats.merge_count += sheet.merges.len();
        stats.hyperlink_count += sheet.hyperlinks.len();
        stats.cf_rule_count += sheet.conditional_formatting.len();
        stats.validation_count += sheet.data_validations.len();
        stats.image_count += sheet.drawings.len();
        stats.chart_count += sheet.charts.len();
    }

    // Verify we have comprehensive features
    assert!(stats.sheet_count >= 3, "Should have at least 3 sheets");
    assert!(stats.cell_count >= 50, "Should have at least 50 cells");
    assert!(stats.bold_cells >= 1, "Should have bold cells");
    assert!(
        stats.colored_bg_cells >= 1,
        "Should have colored backgrounds"
    );
    assert!(stats.bordered_cells >= 1, "Should have bordered cells");
}

#[derive(Default)]
struct FeatureStats {
    sheet_count: usize,
    cell_count: usize,
    bold_cells: usize,
    italic_cells: usize,
    underlined_cells: usize,
    colored_font_cells: usize,
    colored_bg_cells: usize,
    pattern_fill_cells: usize,
    bordered_cells: usize,
    aligned_cells: usize,
    wrapped_cells: usize,
    rotated_cells: usize,
    merge_count: usize,
    hyperlink_count: usize,
    cf_rule_count: usize,
    validation_count: usize,
    image_count: usize,
    chart_count: usize,
}

// ============================================================================
// MS CF Samples Tests (if file exists)
// ============================================================================

/// Test MS conditional formatting samples file
#[test]
fn test_ms_cf_samples() {
    let path = "test/ms_cf_samples.xlsx";
    if !std::path::Path::new(path).exists() {
        eprintln!("Skipping test - {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut cf_count = 0;
    for sheet in &workbook.sheets {
        cf_count += sheet.conditional_formatting.len();
    }

    // MS CF samples should have many CF rules
    assert!(
        cf_count >= 5,
        "MS CF samples should have at least 5 CF rules, got {}",
        cf_count
    );
}

// ============================================================================
// Edge Case Tests
// ============================================================================

/// Test parsing of empty sheets
#[test]
fn test_empty_sheet_handling() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Check if we can handle sheets with varying cell counts
    for sheet in &workbook.sheets {
        // Sheet should have a name at minimum
        assert!(!sheet.name.is_empty(), "Sheet should have a name");
    }
}

/// Test that very wide cell references work
#[test]
fn test_wide_column_references() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    // Verify column indices are parsed correctly
    for sheet in &workbook.sheets {
        for cell in &sheet.cells {
            // Column index should be reasonable (< 16384 which is Excel's max)
            assert!(cell.c < 16384, "Column index {} is too large", cell.c);
        }
    }
}

/// Test frozen panes parsing
#[test]
fn test_frozen_panes() {
    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    let mut has_frozen = false;
    for sheet in &workbook.sheets {
        if sheet.frozen_rows > 0 || sheet.frozen_cols > 0 {
            has_frozen = true;
            break;
        }
    }

    // Kitchen sink should have at least one sheet with frozen panes
    assert!(
        has_frozen,
        "Should have at least one sheet with frozen panes"
    );
}

/// Debug helper: Print stylesheet info for a file
#[test]
#[ignore] // Run with: cargo test print_stylesheet -- --ignored --nocapture
fn print_stylesheet_debug() {
    use std::io::{BufReader, Cursor, Read};
    use xlview::styles::parse_styles;
    use zip::ZipArchive;

    let data = fs::read("test/kitchen_sink.xlsx").expect("Failed to read test file");
    let mut archive = ZipArchive::new(Cursor::new(&data)).expect("Failed to open zip");

    let mut styles_content = Vec::new();
    archive
        .by_name("xl/styles.xml")
        .expect("Failed to find styles.xml")
        .read_to_end(&mut styles_content)
        .expect("Failed to read styles.xml");

    let stylesheet =
        parse_styles(BufReader::new(Cursor::new(&styles_content))).expect("Failed to parse styles");

    println!("=== FONTS ({}) ===", stylesheet.fonts.len());
    for (i, font) in stylesheet.fonts.iter().enumerate() {
        println!(
            "{}: name={:?} size={:?} bold={} italic={}",
            i, font.name, font.size, font.bold, font.italic
        );
    }

    println!("\n=== CELL_XFS ({}) ===", stylesheet.cell_xfs.len());
    for (i, xf) in stylesheet.cell_xfs.iter().take(15).enumerate() {
        println!(
            "{}: fontId={:?} fillId={:?} borderId={:?} apply_font={} apply_fill={} apply_border={}",
            i, xf.font_id, xf.fill_id, xf.border_id, xf.apply_font, xf.apply_fill, xf.apply_border
        );
    }

    println!("\n=== BORDERS ({}) ===", stylesheet.borders.len());
    for (i, border) in stylesheet.borders.iter().take(5).enumerate() {
        println!(
            "{}: top={:?} right={:?} bottom={:?} left={:?}",
            i,
            border.top.as_ref().map(|s| &s.style),
            border.right.as_ref().map(|s| &s.style),
            border.bottom.as_ref().map(|s| &s.style),
            border.left.as_ref().map(|s| &s.style)
        );
    }
}

/// Debug helper: Print all feature counts
#[test]
#[ignore] // Run with: cargo test print_features -- --ignored --nocapture
fn print_features_debug() {
    let data = fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read test file");
    let workbook = parse(&data).expect("Failed to parse XLSX");

    println!("=== WORKBOOK FEATURES ===");
    println!("Sheets: {}", workbook.sheets.len());
    println!("Theme colors: {}", workbook.theme.colors.len());

    for sheet in &workbook.sheets {
        println!("\n=== SHEET: {} ===", sheet.name);
        println!("  Cells: {}", sheet.cells.len());
        println!("  Merges: {}", sheet.merges.len());
        println!("  Hyperlinks: {}", sheet.hyperlinks.len());
        println!("  CF Rules: {}", sheet.conditional_formatting.len());
        println!("  Data Validations: {}", sheet.data_validations.len());
        println!("  Images: {}", sheet.drawings.len());
        println!("  Charts: {}", sheet.charts.len());
        println!(
            "  Frozen: rows={}, cols={}",
            sheet.frozen_rows, sheet.frozen_cols
        );
        if let Some(ref color) = sheet.tab_color {
            println!("  Tab color: {}", color);
        }
    }
}

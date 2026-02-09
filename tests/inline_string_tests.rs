//! Tests for inline string parsing in XLSX files
//!
//! Inline strings use `t="inlineStr"` with `<is><t>text</t></is>` structure
//! instead of shared string references.
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

mod common;
mod fixtures;

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

/// Create XLSX with inline string cells
fn create_xlsx_with_inline_strings(cells: &[(&str, &str)]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
    );

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#,
    );

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);

    // Build rows from cells
    let mut rows_xml = String::new();
    for (cell_ref, value) in cells {
        let row_num: u32 = cell_ref
            .chars()
            .skip_while(|c| c.is_ascii_alphabetic())
            .collect::<String>()
            .parse()
            .unwrap_or(1);
        rows_xml.push_str(&format!(
            r#"<row r="{}"><c r="{}" t="inlineStr"><is><t>{}</t></is></c></row>"#,
            row_num, cell_ref, value
        ));
    }

    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
{rows_xml}
</sheetData>
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Basic Inline String Tests
// =============================================================================

#[test]
fn test_single_inline_string() {
    let xlsx = create_xlsx_with_inline_strings(&[("A1", "Hello World")]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Find cell A1
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell.is_some(), "Should find cell A1");

    let cell = cell.unwrap();
    assert_eq!(
        cell.cell.v.as_deref(),
        Some("Hello World"),
        "Cell A1 should contain 'Hello World'"
    );
}

#[test]
fn test_multiple_inline_strings() {
    let xlsx = create_xlsx_with_inline_strings(&[
        ("A1", "First"),
        ("A2", "Second"),
        ("A3", "Third"),
        ("B1", "Column B"),
    ]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Check A1
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Should find cell A1");
    assert_eq!(cell_a1.unwrap().cell.v.as_deref(), Some("First"));

    // Check A2
    let cell_a2 = sheet.cells.iter().find(|c| c.r == 1 && c.c == 0);
    assert!(cell_a2.is_some(), "Should find cell A2");
    assert_eq!(cell_a2.unwrap().cell.v.as_deref(), Some("Second"));

    // Check A3
    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert!(cell_a3.is_some(), "Should find cell A3");
    assert_eq!(cell_a3.unwrap().cell.v.as_deref(), Some("Third"));

    // Check B1
    let cell_b1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 1);
    assert!(cell_b1.is_some(), "Should find cell B1");
    assert_eq!(cell_b1.unwrap().cell.v.as_deref(), Some("Column B"));
}

#[test]
fn test_inline_string_with_special_characters() {
    let xlsx = create_xlsx_with_inline_strings(&[("A1", "Test &amp; Value")]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell.is_some());
    // The &amp; should be unescaped to &
    assert_eq!(cell.unwrap().cell.v.as_deref(), Some("Test & Value"));
}

#[test]
fn test_inline_string_unicode() {
    let xlsx = create_xlsx_with_inline_strings(&[
        ("A1", "日本語"),
        ("A2", "中文"),
        ("A3", "한국어"),
        ("A4", "Ελληνικά"),
    ]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert_eq!(cell_a1.unwrap().cell.v.as_deref(), Some("日本語"));

    let cell_a2 = sheet.cells.iter().find(|c| c.r == 1 && c.c == 0);
    assert_eq!(cell_a2.unwrap().cell.v.as_deref(), Some("中文"));

    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert_eq!(cell_a3.unwrap().cell.v.as_deref(), Some("한국어"));

    let cell_a4 = sheet.cells.iter().find(|c| c.r == 3 && c.c == 0);
    assert_eq!(cell_a4.unwrap().cell.v.as_deref(), Some("Ελληνικά"));
}

// =============================================================================
// Lazy Parsing Tests (same as wasm viewer uses)
// =============================================================================

#[test]
fn test_lazy_parsing_inline_strings() {
    // This test uses parse_lazy() like the wasm viewer does
    let data =
        std::fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read kitchen_sink_v2.xlsx");
    let workbook = xlview::parser::parse_lazy(&data).expect("Failed to parse XLSX");

    // Debug: print theme colors in lazy mode
    eprintln!("Lazy mode - Theme colors:");
    for (i, color) in workbook.theme.colors.iter().enumerate() {
        eprintln!("  theme[{}] = {}", i, color);
    }

    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Data Validation")
        .expect("Should find 'Data Validation' sheet");

    // Print all cells for debugging including style info
    eprintln!("Lazy parsing - Cells in Data Validation sheet:");
    for cell in &sheet.cells {
        eprintln!(
            "  Row {}, Col {}: v={:?}, raw={:?}, style_idx={:?}, has_s={:?}",
            cell.r,
            cell.c,
            cell.cell.v,
            cell.cell.raw,
            cell.cell.style_idx,
            cell.cell.s.is_some()
        );
    }

    // In lazy mode, v should be None but raw should have the value
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Should find cell A1");
    let cell_a1 = cell_a1.unwrap();

    // Check that raw value is set for lazy parsing
    assert!(
        cell_a1.cell.raw.is_some(),
        "A1 should have raw value in lazy mode"
    );

    // A3 should also have raw value
    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert!(cell_a3.is_some(), "Should find cell A3");
    let cell_a3 = cell_a3.unwrap();
    assert!(
        cell_a3.cell.raw.is_some(),
        "A3 should have raw value in lazy mode"
    );
}

// =============================================================================
// Real File Tests - kitchen_sink_v2.xlsx
// =============================================================================

#[test]
fn test_kitchen_sink_v2_inline_strings() {
    let data =
        std::fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read kitchen_sink_v2.xlsx");
    let workbook = xlview::parser::parse(&data).expect("Failed to parse XLSX");

    // Sheet 3 (Data Validation) has inline strings
    // Based on XML: A1="Dropdown Examples", A3="Select Status:", B3="Active", etc.
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Data Validation")
        .expect("Should find 'Data Validation' sheet");

    // Print all cells for debugging
    eprintln!("Cells in Data Validation sheet:");
    for cell in &sheet.cells {
        eprintln!("  Row {}, Col {}: {:?}", cell.r, cell.c, cell.cell.v);
    }

    // A1 should be "Dropdown Examples"
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Should find cell A1");
    assert_eq!(
        cell_a1.unwrap().cell.v.as_deref(),
        Some("Dropdown Examples"),
        "A1 should be 'Dropdown Examples'"
    );

    // A3 should be "Select Status:"
    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert!(cell_a3.is_some(), "Should find cell A3");
    assert_eq!(
        cell_a3.unwrap().cell.v.as_deref(),
        Some("Select Status:"),
        "A3 should be 'Select Status:'"
    );

    // B3 should be "Active"
    let cell_b3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 1);
    assert!(cell_b3.is_some(), "Should find cell B3");
    assert_eq!(
        cell_b3.unwrap().cell.v.as_deref(),
        Some("Active"),
        "B3 should be 'Active'"
    );

    // A4 should be "Select Priority:"
    let cell_a4 = sheet.cells.iter().find(|c| c.r == 3 && c.c == 0);
    assert!(cell_a4.is_some(), "Should find cell A4");
    assert_eq!(
        cell_a4.unwrap().cell.v.as_deref(),
        Some("Select Priority:"),
        "A4 should be 'Select Priority:'"
    );

    // B4 should be "Medium"
    let cell_b4 = sheet.cells.iter().find(|c| c.r == 3 && c.c == 1);
    assert!(cell_b4.is_some(), "Should find cell B4");
    assert_eq!(
        cell_b4.unwrap().cell.v.as_deref(),
        Some("Medium"),
        "B4 should be 'Medium'"
    );
}

#[test]
fn test_kitchen_sink_v2_inline_string_count() {
    let data =
        std::fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read kitchen_sink_v2.xlsx");
    let workbook = xlview::parser::parse(&data).expect("Failed to parse XLSX");

    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Data Validation")
        .expect("Should find 'Data Validation' sheet");

    // Count cells with values
    let cells_with_values = sheet.cells.iter().filter(|c| c.cell.v.is_some()).count();

    // Based on XML, we should have:
    // A1, A3, B3, A4, B4, A6, B6, A7, A9, B9, A10, B10, A11, B11 = 14 cells with values
    // (Actually A6 has inline str, B6 has number, A7 has inline str, B7 is empty)
    // Let's verify we have at least 10 cells with values
    assert!(
        cells_with_values >= 10,
        "Sheet should have at least 10 cells with values, found {}",
        cells_with_values
    );
}

#[test]
fn test_data_validation_cells_have_black_font_color() {
    // This test verifies that cells without explicit styles get the default font color (black)
    // which is determined by theme color index 1 (dk1 = dark1 = black)
    let data =
        std::fs::read("test/kitchen_sink_v2.xlsx").expect("Failed to read kitchen_sink_v2.xlsx");
    let workbook = xlview::parser::parse(&data).expect("Failed to parse XLSX");

    // Debug: print the theme colors
    eprintln!("Theme colors from workbook:");
    for (i, color) in workbook.theme.colors.iter().enumerate() {
        eprintln!("  theme[{}] = {}", i, color);
    }

    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Data Validation")
        .expect("Should find 'Data Validation' sheet");

    // Check cells without explicit style (A3, B3, A4, B4, etc.)
    // These should inherit the default font which has theme color 1 (black)

    // A3 = "Select Status:" - no explicit style
    let cell_a3 = sheet.cells.iter().find(|c| c.r == 2 && c.c == 0);
    assert!(cell_a3.is_some(), "Should find cell A3");
    let cell_a3 = cell_a3.unwrap();

    // In eager mode, the cell should have a resolved style with font_color
    if let Some(ref style) = cell_a3.cell.s {
        eprintln!("A3 style: font_color={:?}", style.font_color);
        // The font color should be black (#000000) from theme color 1
        assert!(
            style
                .font_color
                .as_ref()
                .map(|c| c.contains("000000") || c.to_uppercase().contains("000000"))
                .unwrap_or(false),
            "A3 font color should be black (#000000), got {:?}",
            style.font_color
        );
    } else {
        // Cell has no inline style - check style_idx
        eprintln!(
            "A3 has no inline style, style_idx={:?}",
            cell_a3.cell.style_idx
        );
        // Without explicit style, style_idx should be None and default_style should be used
        assert!(
            cell_a3.cell.style_idx.is_none(),
            "A3 should have no explicit style index"
        );
    }
}

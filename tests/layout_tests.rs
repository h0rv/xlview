//! Layout feature tests for xlview
//!
//! Tests for column widths, row heights, hidden rows/columns, merged cells,
//! multiple sheets, sheet visibility, tab colors, frozen panes, and default dimensions.
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

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// Helper to create a minimal XLSX file in memory
fn create_xlsx(workbook_xml: &str, sheets: &[(&str, &str)]) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let options = FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        let mut content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#.to_string();

        for (i, _) in sheets.iter().enumerate() {
            content_types.push_str(&format!(
                r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i + 1
            ));
        }
        content_types.push_str("</Types>");
        zip.write_all(content_types.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        // xl/_rels/workbook.xml.rels
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        let mut rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
            .to_string();
        for (i, _) in sheets.iter().enumerate() {
            rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i + 1, i + 1
            ));
        }
        rels.push_str("</Relationships>");
        zip.write_all(rels.as_bytes()).unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        // xl/worksheets/sheet{n}.xml
        for (i, (_, sheet_xml)) in sheets.iter().enumerate() {
            zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), options)
                .unwrap();
            zip.write_all(sheet_xml.as_bytes()).unwrap();
        }

        zip.finish().unwrap();
    }
    buf.into_inner()
}

// Helper for single sheet XLSX
fn create_single_sheet_xlsx(sheet_name: &str, sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheets><sheet name="{}" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets>
</workbook>"#,
        sheet_name
    );
    create_xlsx(&workbook_xml, &[(sheet_name, sheet_xml)])
}

// Wrap sheet content in worksheet XML
fn wrap_sheet(content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{}
</worksheet>"#,
        content
    )
}

// =============================================================================
// COLUMN WIDTH TESTS
// =============================================================================

#[test]
fn test_column_width_custom_single() {
    // Test: col width="15" for single column
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let col_widths = &workbook["sheets"][0]["colWidths"];
    assert!(col_widths.is_array(), "colWidths should be an array");

    // Find column 0 (A)
    let col_a = col_widths
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["col"] == 0);
    assert!(col_a.is_some(), "Column A should have custom width");

    // Width 15 in Excel character units
    let width = col_a.unwrap()["width"].as_f64().unwrap();
    assert!(
        (width - 15.0).abs() < 1.0,
        "Width should be ~15 (Excel units), got {}",
        width
    );
}

#[test]
fn test_column_width_range() {
    // Test: col min="1" max="5" width="20" for multiple columns
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="5" width="20" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c><c r="E1"><v>E</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let col_widths = workbook["sheets"][0]["colWidths"].as_array().unwrap();

    // Should have 5 columns (0-4)
    assert!(
        col_widths.len() >= 5,
        "Should have at least 5 column widths"
    );

    // All columns 0-4 should have width 20 (Excel character units)
    for col_idx in 0..5 {
        let col = col_widths.iter().find(|c| c["col"] == col_idx);
        assert!(
            col.is_some(),
            "Column {} should have width defined",
            col_idx
        );
        let width = col.unwrap()["width"].as_f64().unwrap();
        assert!(
            (width - 20.0).abs() < 1.0,
            "Column {} width should be ~20 (Excel units), got {}",
            col_idx,
            width
        );
    }
}

#[test]
fn test_column_width_default() {
    // Test: No col definition, use default
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    // colWidths should be empty when no custom widths defined
    let col_widths = &workbook["sheets"][0]["colWidths"];
    assert!(
        col_widths.as_array().unwrap().is_empty(),
        "No custom column widths should be set"
    );

    // defaultColWidth should be set
    let default_width = workbook["sheets"][0]["defaultColWidth"].as_f64().unwrap();
    assert!(
        default_width > 0.0,
        "Default column width should be positive"
    );
}

#[test]
fn test_column_width_zero() {
    // Test: width="0" (hidden via width)
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="0" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let col_widths = workbook["sheets"][0]["colWidths"].as_array().unwrap();
    let col_b = col_widths.iter().find(|c| c["col"] == 1);
    assert!(col_b.is_some(), "Column B should be defined");

    // Width 0 in Excel units
    let width = col_b.unwrap()["width"].as_f64().unwrap();
    assert!(
        width < 1.0,
        "Zero width column should be ~0 (Excel units), got {}",
        width
    );
}

#[test]
fn test_column_width_very_wide() {
    // Test: width="100"
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="100" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>Wide</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let col_widths = workbook["sheets"][0]["colWidths"].as_array().unwrap();
    let col_a = col_widths.iter().find(|c| c["col"] == 0);
    assert!(col_a.is_some(), "Column A should have width");

    // Width 100 in Excel units
    let width = col_a.unwrap()["width"].as_f64().unwrap();
    assert!(
        (width - 100.0).abs() < 1.0,
        "Very wide column should be ~100 (Excel units), got {}",
        width
    );
}

// =============================================================================
// ROW HEIGHT TESTS
// =============================================================================

#[test]
fn test_row_height_custom() {
    // Test: row ht="30" customHeight="1"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="30" customHeight="1"><c r="A1"><v>Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let row_heights = workbook["sheets"][0]["rowHeights"].as_array().unwrap();
    let row_1 = row_heights.iter().find(|r| r["row"] == 0);
    assert!(row_1.is_some(), "Row 1 should have custom height");

    // Height 30 points * 1.33 = ~40 pixels
    let height = row_1.unwrap()["height"].as_f64().unwrap();
    assert!(
        (height - 39.9).abs() < 1.0,
        "Row height should be ~40 pixels, got {}",
        height
    );
}

#[test]
fn test_row_height_default() {
    // Test: No ht attribute
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Default height</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let row_heights = workbook["sheets"][0]["rowHeights"].as_array().unwrap();
    assert!(
        row_heights.is_empty(),
        "No custom row heights should be set"
    );

    let default_height = workbook["sheets"][0]["defaultRowHeight"].as_f64().unwrap();
    assert!(
        default_height > 0.0,
        "Default row height should be positive"
    );
}

#[test]
fn test_row_height_zero() {
    // Test: ht="0" (hidden via height)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="2" ht="0" customHeight="1"><c r="A2"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let row_heights = workbook["sheets"][0]["rowHeights"].as_array().unwrap();
    let row_2 = row_heights.iter().find(|r| r["row"] == 1);
    assert!(row_2.is_some(), "Row 2 should have height defined");

    let height = row_2.unwrap()["height"].as_f64().unwrap();
    assert!(
        height < 1.0,
        "Zero height row should be ~0 pixels, got {}",
        height
    );
}

#[test]
fn test_row_height_very_tall() {
    // Test: ht="100"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1" ht="100" customHeight="1"><c r="A1"><v>Tall</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let row_heights = workbook["sheets"][0]["rowHeights"].as_array().unwrap();
    let row_1 = row_heights.iter().find(|r| r["row"] == 0);
    assert!(row_1.is_some(), "Row 1 should have height");

    // Height 100 points * 1.33 = ~133 pixels
    let height = row_1.unwrap()["height"].as_f64().unwrap();
    assert!(
        (height - 133.0).abs() < 1.0,
        "Very tall row should be ~133 pixels, got {}",
        height
    );
}

// =============================================================================
// HIDDEN ROWS/COLUMNS TESTS
// =============================================================================

#[test]
fn test_hidden_column() {
    // Test: col hidden="1"
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="2" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="B1"><v>Hidden</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let hidden_cols = workbook["sheets"][0]["hiddenCols"].as_array().unwrap();
    assert!(
        hidden_cols.contains(&serde_json::json!(1)),
        "Column B (index 1) should be hidden"
    );
}

#[test]
fn test_hidden_row() {
    // Test: row hidden="1"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Visible</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden</v></c></row>
            <row r="3"><c r="A3"><v>Visible</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let hidden_rows = workbook["sheets"][0]["hiddenRows"].as_array().unwrap();
    assert!(
        hidden_rows.contains(&serde_json::json!(1)),
        "Row 2 (index 1) should be hidden"
    );
    assert!(
        !hidden_rows.contains(&serde_json::json!(0)),
        "Row 1 should not be hidden"
    );
    assert!(
        !hidden_rows.contains(&serde_json::json!(2)),
        "Row 3 should not be hidden"
    );
}

#[test]
fn test_hidden_range() {
    // Test: Multiple consecutive hidden columns and rows
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="2" max="4" width="10" hidden="1"/>
        </cols>
        <sheetData>
            <row r="1"><c r="A1"><v>A</v></c></row>
            <row r="2" hidden="1"><c r="A2"><v>2</v></c></row>
            <row r="3" hidden="1"><c r="A3"><v>3</v></c></row>
            <row r="4" hidden="1"><c r="A4"><v>4</v></c></row>
            <row r="5"><c r="A5"><v>5</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let hidden_cols = workbook["sheets"][0]["hiddenCols"].as_array().unwrap();
    assert!(
        hidden_cols.contains(&serde_json::json!(1)),
        "Column B should be hidden"
    );
    assert!(
        hidden_cols.contains(&serde_json::json!(2)),
        "Column C should be hidden"
    );
    assert!(
        hidden_cols.contains(&serde_json::json!(3)),
        "Column D should be hidden"
    );

    let hidden_rows = workbook["sheets"][0]["hiddenRows"].as_array().unwrap();
    assert!(
        hidden_rows.contains(&serde_json::json!(1)),
        "Row 2 should be hidden"
    );
    assert!(
        hidden_rows.contains(&serde_json::json!(2)),
        "Row 3 should be hidden"
    );
    assert!(
        hidden_rows.contains(&serde_json::json!(3)),
        "Row 4 should be hidden"
    );
}

// =============================================================================
// MERGED CELLS TESTS
// =============================================================================

#[test]
fn test_merge_simple() {
    // Test: mergeCell ref="A1:B2"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Merged</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1, "Should have 1 merge");

    let merge = &merges[0];
    assert_eq!(merge["startRow"], 0);
    assert_eq!(merge["startCol"], 0);
    assert_eq!(merge["endRow"], 1);
    assert_eq!(merge["endCol"], 1);
}

#[test]
fn test_merge_wide() {
    // Test: mergeCell ref="A1:Z1" (26 columns wide)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Wide merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:Z1"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1);

    let merge = &merges[0];
    assert_eq!(merge["startRow"], 0);
    assert_eq!(merge["startCol"], 0);
    assert_eq!(merge["endRow"], 0);
    assert_eq!(merge["endCol"], 25); // Z is column 26 (0-indexed: 25)
}

#[test]
fn test_merge_tall() {
    // Test: mergeCell ref="A1:A100" (100 rows tall)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Tall merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:A100"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1);

    let merge = &merges[0];
    assert_eq!(merge["startRow"], 0);
    assert_eq!(merge["startCol"], 0);
    assert_eq!(merge["endRow"], 99); // Row 100 (0-indexed: 99)
    assert_eq!(merge["endCol"], 0);
}

#[test]
fn test_merge_large() {
    // Test: mergeCell ref="A1:D10"
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Large merge</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:D10"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1);

    let merge = &merges[0];
    assert_eq!(merge["startRow"], 0);
    assert_eq!(merge["startCol"], 0);
    assert_eq!(merge["endRow"], 9);
    assert_eq!(merge["endCol"], 3); // D is column 4 (0-indexed: 3)
}

#[test]
fn test_merge_multiple() {
    // Test: Several non-overlapping merges
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Merge 1</v></c><c r="D1"><v>Merge 2</v></c></row>
            <row r="5"><c r="A5"><v>Merge 3</v></c></row>
        </sheetData>
        <mergeCells count="3">
            <mergeCell ref="A1:B2"/>
            <mergeCell ref="D1:F1"/>
            <mergeCell ref="A5:C7"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 3, "Should have 3 merges");

    // Verify each merge exists (order may vary)
    let has_a1_b2 = merges
        .iter()
        .any(|m| m["startRow"] == 0 && m["startCol"] == 0 && m["endRow"] == 1 && m["endCol"] == 1);
    let has_d1_f1 = merges
        .iter()
        .any(|m| m["startRow"] == 0 && m["startCol"] == 3 && m["endRow"] == 0 && m["endCol"] == 5);
    let has_a5_c7 = merges
        .iter()
        .any(|m| m["startRow"] == 4 && m["startCol"] == 0 && m["endRow"] == 6 && m["endCol"] == 2);

    assert!(has_a1_b2, "Should have A1:B2 merge");
    assert!(has_d1_f1, "Should have D1:F1 merge");
    assert!(has_a5_c7, "Should have A5:C7 merge");
}

#[test]
fn test_merge_with_content() {
    // Test: Value in top-left cell only
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1">
                <c r="A1"><v>Content in top-left</v></c>
                <c r="B1"/>
            </row>
            <row r="2">
                <c r="A2"/>
                <c r="B2"/>
            </row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let merges = workbook["sheets"][0]["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1);

    // Find cell A1 and verify it has content
    let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
    let cell_a1 = cells.iter().find(|c| c["r"] == 0 && c["c"] == 0);
    assert!(cell_a1.is_some(), "Cell A1 should exist");
    assert_eq!(cell_a1.unwrap()["cell"]["v"], "Content in top-left");
}

// =============================================================================
// MULTIPLE SHEETS TESTS
// =============================================================================

#[test]
fn test_two_sheets() {
    // Test: workbook with 2 sheets
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="First Sheet" sheetId="1" r:id="rId1"/>
    <sheet name="Second Sheet" sheetId="2" r:id="rId2"/>
</sheets>
</workbook>"#;

    let sheet1 = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Sheet 1 Content</v></c></row>
        </sheetData>
    "#,
    );

    let sheet2 = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>Sheet 2 Content</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_xlsx(
        workbook_xml,
        &[("First Sheet", &sheet1), ("Second Sheet", &sheet2)],
    );
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 2, "Should have 2 sheets");
    assert_eq!(sheets[0]["name"], "First Sheet");
    assert_eq!(sheets[1]["name"], "Second Sheet");
}

#[test]
fn test_many_sheets() {
    // Test: workbook with 10 sheets
    let mut sheet_elements = String::new();
    let mut sheets_data: Vec<(String, String)> = Vec::new();

    for i in 1..=10 {
        sheet_elements.push_str(&format!(
            r#"<sheet name="Sheet{}" sheetId="{}" r:id="rId{}"/>"#,
            i, i, i
        ));
        let content = wrap_sheet(&format!(
            r#"
            <sheetData>
                <row r="1"><c r="A1"><v>Content {}</v></c></row>
            </sheetData>
        "#,
            i
        ));
        sheets_data.push((format!("Sheet{}", i), content));
    }

    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{}</sheets>
</workbook>"#,
        sheet_elements
    );

    let sheets_ref: Vec<(&str, &str)> = sheets_data
        .iter()
        .map(|(n, c)| (n.as_str(), c.as_str()))
        .collect();

    let xlsx = create_xlsx(&workbook_xml, &sheets_ref);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 10, "Should have 10 sheets");

    for i in 1..=10 {
        assert_eq!(sheets[i - 1]["name"], format!("Sheet{}", i));
    }
}

#[test]
fn test_sheet_names_special() {
    // Test: Various names including spaces, unicode
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Sales Data 2024" sheetId="1" r:id="rId1"/>
    <sheet name="Q1 Summary" sheetId="2" r:id="rId2"/>
    <sheet name="Sheet-With-Dashes" sheetId="3" r:id="rId3"/>
</sheets>
</workbook>"#;

    let sheet1 = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#);
    let sheet2 = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>"#);
    let sheet3 = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>"#);

    let xlsx = create_xlsx(
        workbook_xml,
        &[
            ("Sales Data 2024", &sheet1),
            ("Q1 Summary", &sheet2),
            ("Sheet-With-Dashes", &sheet3),
        ],
    );
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets[0]["name"], "Sales Data 2024");
    assert_eq!(sheets[1]["name"], "Q1 Summary");
    assert_eq!(sheets[2]["name"], "Sheet-With-Dashes");
}

#[test]
fn test_sheet_order() {
    // Test: Verify sheets are in correct order
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId1"/>
    <sheet name="Beta" sheetId="2" r:id="rId2"/>
    <sheet name="Gamma" sheetId="3" r:id="rId3"/>
    <sheet name="Delta" sheetId="4" r:id="rId4"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>X</v></c></row></sheetData>"#);

    let xlsx = create_xlsx(
        workbook_xml,
        &[
            ("Alpha", &sheet),
            ("Beta", &sheet),
            ("Gamma", &sheet),
            ("Delta", &sheet),
        ],
    );
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 4);

    // Verify order matches declaration order in workbook.xml
    assert_eq!(sheets[0]["name"], "Alpha");
    assert_eq!(sheets[1]["name"], "Beta");
    assert_eq!(sheets[2]["name"], "Gamma");
    assert_eq!(sheets[3]["name"], "Delta");
}

// =============================================================================
// SHEET VISIBILITY TESTS
// Note: These test the expected XML structure. The current parser may not
// fully support visibility attributes yet - these tests document expected behavior.
// =============================================================================

#[test]
fn test_sheet_visibility_visible() {
    // Test: state="visible" (default)
    // Currently the parser does not expose sheet visibility state.
    // This test documents the expected XML structure.
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Visible Sheet" sheetId="1" state="visible" r:id="rId1"/>
</sheets>
</workbook>"#;

    let sheet =
        wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>Visible</v></c></row></sheetData>"#);

    let xlsx = create_xlsx(workbook_xml, &[("Visible Sheet", &sheet)]);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    // Sheet should be parsed successfully
    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0]["name"], "Visible Sheet");

    // TODO: When visibility support is added, verify:
    // assert_eq!(sheets[0]["state"], "visible");
}

#[test]
fn test_sheet_visibility_hidden() {
    // Test: state="hidden"
    // Documents expected behavior for hidden sheets
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Visible" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>X</v></c></row></sheetData>"#);

    let xlsx = create_xlsx(workbook_xml, &[("Visible", &sheet), ("Hidden", &sheet)]);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 2);

    // TODO: When visibility support is added, verify:
    // assert_eq!(sheets[1]["state"], "hidden");
}

#[test]
fn test_sheet_visibility_very_hidden() {
    // Test: state="veryHidden"
    // Very hidden sheets cannot be unhidden through the UI
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Normal" sheetId="1" r:id="rId1"/>
    <sheet name="VeryHidden" sheetId="2" state="veryHidden" r:id="rId2"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(r#"<sheetData><row r="1"><c r="A1"><v>X</v></c></row></sheetData>"#);

    let xlsx = create_xlsx(workbook_xml, &[("Normal", &sheet), ("VeryHidden", &sheet)]);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 2);

    // TODO: When visibility support is added, verify:
    // assert_eq!(sheets[1]["state"], "veryHidden");
}

// =============================================================================
// SHEET TAB COLOR TESTS
// Note: These test the expected XML structure. The current parser may not
// fully support tab colors yet - these tests document expected behavior.
// =============================================================================

#[test]
fn test_tab_color_rgb() {
    // Test: tabColor rgb="FF00FF00"
    // Documents expected behavior for RGB tab colors
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Green Tab" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#;

    // Tab color is defined in the sheet's sheetPr element
    let sheet = wrap_sheet(
        r#"
        <sheetPr>
            <tabColor rgb="FF00FF00"/>
        </sheetPr>
        <sheetData><row r="1"><c r="A1"><v>Green</v></c></row></sheetData>
    "#,
    );

    let xlsx = create_xlsx(workbook_xml, &[("Green Tab", &sheet)]);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When tab color support is added, verify:
    // assert_eq!(sheets[0]["tabColor"], "#00FF00");
}

#[test]
fn test_tab_color_theme() {
    // Test: tabColor theme="4"
    // Documents expected behavior for theme-based tab colors
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
    <sheet name="Theme Tab" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#;

    let sheet = wrap_sheet(
        r#"
        <sheetPr>
            <tabColor theme="4"/>
        </sheetPr>
        <sheetData><row r="1"><c r="A1"><v>Theme</v></c></row></sheetData>
    "#,
    );

    let xlsx = create_xlsx(workbook_xml, &[("Theme Tab", &sheet)]);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When tab color support is added, verify theme color resolution
}

// =============================================================================
// FROZEN PANES TESTS
// Note: These test the expected XML structure. The current parser may not
// fully support frozen panes yet - these tests document expected behavior.
// =============================================================================

#[test]
fn test_frozen_rows() {
    // Test: pane ySplit="1" topLeftCell="A2" state="frozen"
    // Freeze the first row
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Header</v></c></row>
            <row r="2"><c r="A2"><v>Data</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When frozen pane support is added, verify:
    // assert_eq!(sheets[0]["frozenRows"], 1);
    // assert_eq!(sheets[0]["frozenCols"], 0);
}

#[test]
fn test_frozen_columns() {
    // Test: pane xSplit="1" topLeftCell="B1" state="frozen"
    // Freeze the first column
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Label</v></c><c r="B1"><v>Value</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When frozen pane support is added, verify:
    // assert_eq!(sheets[0]["frozenRows"], 0);
    // assert_eq!(sheets[0]["frozenCols"], 1);
}

#[test]
fn test_frozen_both() {
    // Test: xSplit and ySplit both set
    // Freeze first 2 rows and first column
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Corner</v></c><c r="B1"><v>Header 1</v></c></row>
            <row r="2"><c r="A2"><v>Label</v></c><c r="B2"><v>Header 2</v></c></row>
            <row r="3"><c r="A3"><v>Row Label</v></c><c r="B3"><v>Data</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When frozen pane support is added, verify:
    // assert_eq!(sheets[0]["frozenRows"], 2);
    // assert_eq!(sheets[0]["frozenCols"], 1);
}

#[test]
fn test_split_panes() {
    // Test: state="split" (different from frozen)
    // Split panes allow scrolling in both sections independently
    let sheet_xml = wrap_sheet(
        r#"
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <pane xSplit="2000" ySplit="1500" topLeftCell="C5" activePane="bottomRight" state="split"/>
            </sheetView>
        </sheetViews>
        <sheetData>
            <row r="1"><c r="A1"><v>Split</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheets = workbook["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);

    // TODO: When split pane support is added, verify:
    // Split panes use pixel positions, not row/column counts
    // assert_eq!(sheets[0]["splitState"], "split");
}

// =============================================================================
// DEFAULT DIMENSIONS TESTS
// =============================================================================

#[test]
fn test_sheet_format_pr_defaults() {
    // Test: sheetFormatPr with defaultColWidth and defaultRowHeight
    let sheet_xml = wrap_sheet(
        r#"
        <sheetFormatPr defaultColWidth="12.5" defaultRowHeight="18"/>
        <sheetData>
            <row r="1"><c r="A1"><v>Test</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    // Currently the parser uses hardcoded defaults.
    // This test documents expected behavior when sheetFormatPr is supported.
    let sheet = &workbook["sheets"][0];

    // Verify defaults are present (may be hardcoded or from sheetFormatPr)
    let default_col_width = sheet["defaultColWidth"].as_f64().unwrap();
    let default_row_height = sheet["defaultRowHeight"].as_f64().unwrap();

    assert!(
        default_col_width > 0.0,
        "Default column width should be positive"
    );
    assert!(
        default_row_height > 0.0,
        "Default row height should be positive"
    );

    // TODO: When sheetFormatPr parsing is added:
    // assert!((default_col_width - expected_pixel_width).abs() < 1.0);
    // assert!((default_row_height - expected_pixel_height).abs() < 1.0);
}

// =============================================================================
// COMBINED LAYOUT TESTS
// =============================================================================

#[test]
fn test_complex_layout() {
    // Test combining multiple layout features
    let sheet_xml = wrap_sheet(
        r#"
        <cols>
            <col min="1" max="1" width="25" customWidth="1"/>
            <col min="2" max="2" width="10" hidden="1"/>
            <col min="3" max="5" width="15" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1" ht="30" customHeight="1">
                <c r="A1"><v>Header</v></c>
                <c r="C1"><v>Merged Header</v></c>
            </row>
            <row r="2" hidden="1"><c r="A2"><v>Hidden Row</v></c></row>
            <row r="3"><c r="A3"><v>Data</v></c></row>
        </sheetData>
        <mergeCells count="1">
            <mergeCell ref="C1:E1"/>
        </mergeCells>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheet = &workbook["sheets"][0];

    // Verify column widths
    let col_widths = sheet["colWidths"].as_array().unwrap();
    assert!(
        col_widths.len() >= 5,
        "Should have column width definitions"
    );

    // Verify hidden column
    let hidden_cols = sheet["hiddenCols"].as_array().unwrap();
    assert!(
        hidden_cols.contains(&serde_json::json!(1)),
        "Column B should be hidden"
    );

    // Verify row height
    let row_heights = sheet["rowHeights"].as_array().unwrap();
    let row_1 = row_heights.iter().find(|r| r["row"] == 0);
    assert!(row_1.is_some(), "Row 1 should have custom height");

    // Verify hidden row
    let hidden_rows = sheet["hiddenRows"].as_array().unwrap();
    assert!(
        hidden_rows.contains(&serde_json::json!(1)),
        "Row 2 should be hidden"
    );

    // Verify merge
    let merges = sheet["merges"].as_array().unwrap();
    assert_eq!(merges.len(), 1);
    let merge = &merges[0];
    assert_eq!(merge["startCol"], 2); // C
    assert_eq!(merge["endCol"], 4); // E
}

#[test]
fn test_max_dimensions() {
    // Test that maxRow and maxCol are calculated correctly
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>A1</v></c></row>
            <row r="5"><c r="D5"><v>D5</v></c></row>
            <row r="10"><c r="J10"><v>J10</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheet = &workbook["sheets"][0];

    // maxRow should be 10 (1-indexed in source)
    let max_row = sheet["maxRow"].as_u64().unwrap();
    assert_eq!(max_row, 10, "maxRow should be 10");

    // maxCol should be 10 (J is column 10, stored as count)
    let max_col = sheet["maxCol"].as_u64().unwrap();
    assert_eq!(max_col, 10, "maxCol should be 10");
}

#[test]
fn test_sparse_data() {
    // Test with sparse data (gaps in rows and columns)
    let sheet_xml = wrap_sheet(
        r#"
        <sheetData>
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="100"><c r="Z100"><v>100</v></c></row>
        </sheetData>
    "#,
    );

    let xlsx = create_single_sheet_xlsx("Sheet1", &sheet_xml);
    let workbook: serde_json::Value =
        serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
            .expect("Failed to parse JSON");

    let sheet = &workbook["sheets"][0];

    // Only 2 cells should be in the sparse representation
    let cells = sheet["cells"].as_array().unwrap();
    assert_eq!(cells.len(), 2, "Should only have 2 cells");

    // Verify cell positions
    let cell_a1 = cells.iter().find(|c| c["r"] == 0 && c["c"] == 0);
    let cell_z100 = cells.iter().find(|c| c["r"] == 99 && c["c"] == 25);

    assert!(cell_a1.is_some(), "Cell A1 should exist");
    assert!(cell_z100.is_some(), "Cell Z100 should exist");
}

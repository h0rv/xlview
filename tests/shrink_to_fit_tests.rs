//! Tests for shrink_to_fit alignment styling
//!
//! Shrink to fit is an Excel alignment feature that reduces the font size
//! to fit content within a cell without wrapping. When shrinkToFit="1" is
//! specified in the alignment element, the text is scaled down to fit.
//!
//! According to ECMA-376:
//! - shrinkToFit is mutually exclusive with wrapText
//! - When both are set, wrapText takes precedence
//! - The attribute is stored in xl/styles.xml in <alignment> elements
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
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Test Helpers
// =============================================================================

/// Create an XLSX with a custom styles.xml that has shrinkToFit alignment
fn create_xlsx_with_shrink_to_fit(shrink_to_fit: bool, wrap_text: bool) -> Vec<u8> {
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

    // xl/styles.xml with shrinkToFit
    let shrink_attr = if shrink_to_fit {
        r#" shrinkToFit="1""#
    } else {
        ""
    };
    let wrap_attr = if wrap_text { r#" wrapText="1""# } else { "" };

    let styles_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="2">
  <xf fontId="0" fillId="0" borderId="0"/>
  <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
    <alignment{}{} horizontal="center"/>
  </xf>
</cellXfs>
</styleSheet>"#,
        shrink_attr, wrap_attr
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(styles_xml.as_bytes());

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
  <c r="A1" s="1" t="s"><v>0</v></c>
  <c r="B1" s="0"><v>42</v></c>
</row>
</sheetData>
</worksheet>"#,
    );

    // xl/sharedStrings.xml
    let _ = zip.start_file("xl/sharedStrings.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
<si><t>This is long text that needs shrinking</t></si>
</sst>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests: Shrink to Fit Basic Parsing
// =============================================================================

#[test]
fn test_shrink_to_fit_enabled() {
    let xlsx = create_xlsx_with_shrink_to_fit(true, false);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(!sheet.cells.is_empty(), "Sheet should have cells");

    // Find cell A1 which has style index 1 (with shrinkToFit)
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Cell A1 should exist");

    let cell = cell_a1.unwrap();
    assert!(cell.cell.s.is_some(), "Cell A1 should have a style");

    let style = cell.cell.s.as_ref().unwrap();
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "Cell A1 should have shrinkToFit=true"
    );
}

#[test]
fn test_shrink_to_fit_disabled() {
    let xlsx = create_xlsx_with_shrink_to_fit(false, false);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Cell A1 should exist");

    let cell = cell_a1.unwrap();
    // When shrinkToFit is not set, it should be None or Some(false)
    if let Some(ref style) = cell.cell.s {
        // If shrink_to_fit is Some, it should be false
        // If it's None, that's also acceptable (default behavior)
        if let Some(shrink) = style.shrink_to_fit {
            assert!(!shrink, "Cell A1 should not have shrinkToFit=true");
        }
    }
}

#[test]
fn test_wrap_text_takes_precedence_over_shrink_to_fit() {
    // When both wrapText and shrinkToFit are set, wrapText takes precedence
    // according to ECMA-376
    let xlsx = create_xlsx_with_shrink_to_fit(true, true);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell_a1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0);
    assert!(cell_a1.is_some(), "Cell A1 should exist");

    let cell = cell_a1.unwrap();
    let style = cell.cell.s.as_ref().expect("Cell should have style");

    // Both values should be parsed and stored
    assert_eq!(style.wrap, Some(true), "wrapText should be true");
    assert_eq!(
        style.shrink_to_fit,
        Some(true),
        "shrinkToFit should be true"
    );
    // Note: The renderer should handle the precedence - the parser stores both values
}

#[test]
fn test_cell_without_shrink_to_fit() {
    let xlsx = create_xlsx_with_shrink_to_fit(true, false);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Find cell B1 which has style index 0 (no alignment settings)
    let cell_b1 = sheet.cells.iter().find(|c| c.r == 0 && c.c == 1);
    assert!(cell_b1.is_some(), "Cell B1 should exist");

    let cell = cell_b1.unwrap();
    // Cell B1 uses default style (index 0), so shrink_to_fit should not be set
    if let Some(ref style) = cell.cell.s {
        if let Some(shrink) = style.shrink_to_fit {
            assert!(!shrink, "Default style should not have shrinkToFit enabled");
        }
    }
}

// =============================================================================
// Tests: Shrink to Fit with Various Alignments
// =============================================================================

/// Create an XLSX with shrinkToFit and various horizontal alignments
fn create_xlsx_with_shrink_and_alignment(h_align: &str) -> Vec<u8> {
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
    let styles_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="2">
  <xf fontId="0" fillId="0" borderId="0"/>
  <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
    <alignment shrinkToFit="1" horizontal="{}"/>
  </xf>
</cellXfs>
</styleSheet>"#,
        h_align
    );

    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(styles_xml.as_bytes());

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" s="1"><v>12345</v></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

#[test]
fn test_shrink_to_fit_with_left_align() {
    let xlsx = create_xlsx_with_shrink_and_alignment("left");
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0).unwrap();
    let style = cell.cell.s.as_ref().expect("Cell should have style");

    assert_eq!(style.shrink_to_fit, Some(true));
    assert!(
        style.align_h.is_some(),
        "Horizontal alignment should be set"
    );
}

#[test]
fn test_shrink_to_fit_with_center_align() {
    let xlsx = create_xlsx_with_shrink_and_alignment("center");
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0).unwrap();
    let style = cell.cell.s.as_ref().expect("Cell should have style");

    assert_eq!(style.shrink_to_fit, Some(true));
}

#[test]
fn test_shrink_to_fit_with_right_align() {
    let xlsx = create_xlsx_with_shrink_and_alignment("right");
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0).unwrap();
    let style = cell.cell.s.as_ref().expect("Cell should have style");

    assert_eq!(style.shrink_to_fit, Some(true));
}

// =============================================================================
// Tests: JSON Serialization
// =============================================================================

#[test]
fn test_shrink_to_fit_serialization() {
    let xlsx = create_xlsx_with_shrink_to_fit(true, false);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let json = serde_json::to_value(&workbook).expect("Failed to serialize");

    // Find the cell with shrinkToFit
    let cells = &json["sheets"][0]["cells"];
    assert!(cells.is_array(), "cells should be an array");

    // Find cell A1 (r=0, c=0)
    let cell_a1 = cells
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["r"].as_u64() == Some(0) && c["c"].as_u64() == Some(0));

    assert!(cell_a1.is_some(), "Cell A1 should exist in JSON");
    let cell_a1 = cell_a1.unwrap();

    let style = &cell_a1["cell"]["s"];
    assert!(style.is_object(), "Style should be an object");
    assert_eq!(
        style["shrinkToFit"], true,
        "shrinkToFit should be true in JSON"
    );
}

#[test]
fn test_shrink_to_fit_omitted_when_false() {
    let xlsx = create_xlsx_with_shrink_to_fit(false, false);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let json = serde_json::to_value(&workbook).expect("Failed to serialize");

    // Find cell A1
    let cells = &json["sheets"][0]["cells"];
    let cell_a1 = cells
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["r"].as_u64() == Some(0) && c["c"].as_u64() == Some(0));

    if let Some(cell_a1) = cell_a1 {
        let style = &cell_a1["cell"]["s"];
        if style.is_object() {
            // shrinkToFit should be omitted or null when not set/false
            let shrink_value = &style["shrinkToFit"];
            assert!(
                shrink_value.is_null() || shrink_value == &serde_json::Value::Bool(false),
                "shrinkToFit should be null or false, got: {:?}",
                shrink_value
            );
        }
    }
}

// =============================================================================
// Tests: Real XLSX File Parsing (if available)
// =============================================================================

#[test]
fn test_kitchen_sink_v2_shrink_to_fit_cells() {
    // This test checks if any cells in kitchen_sink_v2.xlsx have shrinkToFit
    // The file may or may not have such cells - this test documents what's present
    let path = "test/kitchen_sink_v2.xlsx";
    if !std::path::Path::new(path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read test file");
    let workbook = xlview::parser::parse(&data).expect("Failed to parse XLSX");

    let mut shrink_to_fit_count = 0;

    for sheet in &workbook.sheets {
        for cell_data in &sheet.cells {
            if let Some(ref style) = cell_data.cell.s {
                if style.shrink_to_fit == Some(true) {
                    shrink_to_fit_count += 1;
                    println!(
                        "Found shrinkToFit cell in sheet '{}' at ({}, {})",
                        sheet.name, cell_data.r, cell_data.c
                    );
                }
            }
        }
    }

    // Document the count (may be 0 if file doesn't have shrinkToFit cells)
    println!(
        "kitchen_sink_v2.xlsx has {} cells with shrinkToFit=true",
        shrink_to_fit_count
    );
}

#[test]
fn test_ms_cf_samples_shrink_to_fit_cells() {
    // This test checks ms_cf_samples.xlsx for shrinkToFit cells
    let path = "test/ms_cf_samples.xlsx";
    if !std::path::Path::new(path).exists() {
        println!("Skipping test: {} not found", path);
        return;
    }

    let data = fs::read(path).expect("Failed to read test file");
    let workbook = xlview::parser::parse(&data).expect("Failed to parse XLSX");

    let mut shrink_to_fit_count = 0;

    for sheet in &workbook.sheets {
        for cell_data in &sheet.cells {
            if let Some(ref style) = cell_data.cell.s {
                if style.shrink_to_fit == Some(true) {
                    shrink_to_fit_count += 1;
                    println!(
                        "Found shrinkToFit cell in sheet '{}' at ({}, {}): {:?}",
                        sheet.name, cell_data.r, cell_data.c, cell_data.cell.v
                    );
                }
            }
        }
    }

    println!(
        "ms_cf_samples.xlsx has {} cells with shrinkToFit=true",
        shrink_to_fit_count
    );
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_shrink_to_fit_value_zero() {
    // Test that shrinkToFit="0" is parsed as false
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Minimal XLSX setup
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

    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );

    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
    );

    // styles.xml with shrinkToFit="0"
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="2">
  <xf fontId="0" fillId="0" borderId="0"/>
  <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
    <alignment shrinkToFit="0"/>
  </xf>
</cellXfs>
</styleSheet>"#,
    );

    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" s="1"><v>Test</v></c></row>
</sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    let xlsx = cursor.into_inner();

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let cell = sheet.cells.iter().find(|c| c.r == 0 && c.c == 0).unwrap();

    if let Some(ref style) = cell.cell.s {
        // shrinkToFit="0" should result in false or None
        if let Some(shrink) = style.shrink_to_fit {
            assert!(!shrink, "shrinkToFit='0' should be false");
        }
    }
}

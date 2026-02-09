//! Integration tests for hyperlink detection and rendering preparation.
//!
//! These tests verify that cells with hyperlinks have the hyperlink data available
//! for rendering. The visual styling (blue #0563C1, underlined) is handled by
//! the rendering layer, but the parser must correctly populate hyperlink fields.
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
    clippy::clone_on_copy,
    clippy::needless_pass_by_value
)]

mod common;
mod fixtures;

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Helper Functions
// ============================================================================

/// Create an XLSX with hyperlinks attached to cells.
fn create_xlsx_with_hyperlinks_for_rendering(
    cells_and_hyperlinks: Vec<(&str, &str, Option<HyperlinkInfo>)>,
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Collect hyperlinks and relationship IDs
    let mut hyperlinks_xml = String::new();
    let mut rels_entries = Vec::new();
    let mut rid_counter = 1;

    for (cell_ref, _text, hyperlink_info) in &cells_and_hyperlinks {
        if let Some(info) = hyperlink_info {
            if info.is_external {
                // External hyperlink uses relationship
                hyperlinks_xml.push_str(&format!(
                    r#"<hyperlink ref="{}" r:id="rId{}"{}/>"#,
                    cell_ref,
                    rid_counter,
                    info.tooltip
                        .as_ref()
                        .map(|t| format!(r#" tooltip="{}""#, t))
                        .unwrap_or_default()
                ));
                rels_entries.push((rid_counter, info.target.clone()));
                rid_counter += 1;
            } else {
                // Internal hyperlink uses location attribute
                hyperlinks_xml.push_str(&format!(
                    r#"<hyperlink ref="{}" location="{}"{}/>"#,
                    cell_ref,
                    info.target,
                    info.tooltip
                        .as_ref()
                        .map(|t| format!(r#" tooltip="{}""#, t))
                        .unwrap_or_default()
                ));
            }
        }
    }

    let has_hyperlinks = !hyperlinks_xml.is_empty();
    if has_hyperlinks {
        hyperlinks_xml = format!("<hyperlinks>{}</hyperlinks>", hyperlinks_xml);
    }

    // Build relationships XML if we have external hyperlinks
    let rels_xml = if !rels_entries.is_empty() {
        let mut rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (rid, target) in &rels_entries {
            rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{}" TargetMode="External"/>"#,
                rid, target
            ));
        }
        rels.push_str("</Relationships>");
        Some(rels)
    } else {
        None
    };

    // Collect shared strings
    let shared_strings: Vec<&str> = cells_and_hyperlinks.iter().map(|(_, t, _)| *t).collect();

    // [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options.clone());
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#,
    );

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options.clone());
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options.clone());
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options.clone());
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#,
    );

    // xl/styles.xml (minimal)
    let _ = zip.start_file("xl/styles.xml", options.clone());
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#,
    );

    // xl/sharedStrings.xml
    let _ = zip.start_file("xl/sharedStrings.xml", options.clone());
    let mut sst = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">"#,
        shared_strings.len(),
        shared_strings.len()
    );
    for s in &shared_strings {
        sst.push_str(&format!("<si><t>{}</t></si>", s));
    }
    sst.push_str("</sst>");
    let _ = zip.write_all(sst.as_bytes());

    // xl/worksheets/_rels/sheet1.xml.rels (if we have external hyperlinks)
    if let Some(ref rels) = rels_xml {
        let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options.clone());
        let _ = zip.write_all(rels.as_bytes());
    }

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options.clone());
    let cells_xml: String = cells_and_hyperlinks
        .iter()
        .enumerate()
        .map(|(i, (cell_ref, _, _))| format!(r#"<c r="{}" t="s"><v>{}</v></c>"#, cell_ref, i))
        .collect::<Vec<_>>()
        .join("\n");

    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1">{}</row>
</sheetData>
{}</worksheet>"#,
        cells_xml, hyperlinks_xml
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Hyperlink information for test fixture creation.
struct HyperlinkInfo {
    target: String,
    is_external: bool,
    tooltip: Option<String>,
}

impl HyperlinkInfo {
    fn external(target: &str) -> Self {
        Self {
            target: target.to_string(),
            is_external: true,
            tooltip: None,
        }
    }

    fn internal(target: &str) -> Self {
        Self {
            target: target.to_string(),
            is_external: false,
            tooltip: None,
        }
    }

    #[allow(dead_code)]
    fn with_tooltip(mut self, tooltip: &str) -> Self {
        self.tooltip = Some(tooltip.to_string());
        self
    }
}

// ============================================================================
// Test 1: Cell with External URL Hyperlink
// ============================================================================

#[test]
fn test_cell_with_external_hyperlink() {
    // Create XLSX with a cell that has an external URL hyperlink
    let xlsx = create_xlsx_with_hyperlinks_for_rendering(vec![(
        "A1",
        "Click here",
        Some(HyperlinkInfo::external("https://example.com")),
    )]);

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    // Verify the sheet has hyperlinks at sheet level
    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 1, "Sheet should have 1 hyperlink");

    let hyperlink_def = &sheet.hyperlinks[0];
    assert_eq!(hyperlink_def.cell_ref, "A1");
    assert_eq!(hyperlink_def.hyperlink.target, "https://example.com");
    assert!(
        hyperlink_def.hyperlink.is_external,
        "Hyperlink should be marked as external"
    );

    // Verify the cell itself has the hyperlink attached
    let cell_data = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    assert!(
        cell_data.cell.hyperlink.is_some(),
        "Cell should have hyperlink field populated"
    );

    let cell_hyperlink = cell_data.cell.hyperlink.as_ref().unwrap();
    assert_eq!(cell_hyperlink.target, "https://example.com");
    assert!(cell_hyperlink.is_external);
}

// ============================================================================
// Test 2: Cell with Internal Sheet Reference Hyperlink
// ============================================================================

#[test]
fn test_cell_with_internal_hyperlink() {
    // Create XLSX with a cell that has an internal sheet reference hyperlink
    let xlsx = create_xlsx_with_hyperlinks_for_rendering(vec![(
        "A1",
        "Go to Sheet2",
        Some(HyperlinkInfo::internal("Sheet2!A1")),
    )]);

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 1, "Sheet should have 1 hyperlink");

    let hyperlink_def = &sheet.hyperlinks[0];
    assert_eq!(hyperlink_def.cell_ref, "A1");
    assert_eq!(hyperlink_def.hyperlink.target, "Sheet2!A1");
    assert!(
        !hyperlink_def.hyperlink.is_external,
        "Internal hyperlink should not be marked as external"
    );

    // Verify the cell has the hyperlink attached
    let cell_data = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    assert!(
        cell_data.cell.hyperlink.is_some(),
        "Cell should have hyperlink field populated for internal link"
    );

    let cell_hyperlink = cell_data.cell.hyperlink.as_ref().unwrap();
    assert_eq!(cell_hyperlink.target, "Sheet2!A1");
    assert!(!cell_hyperlink.is_external);
    assert_eq!(
        cell_hyperlink.location.as_deref(),
        Some("Sheet2!A1"),
        "Internal link should have location field set"
    );
}

// ============================================================================
// Test 3: Cell Without Hyperlink Has None
// ============================================================================

#[test]
fn test_cell_without_hyperlink_has_none() {
    // Create XLSX with a cell that has no hyperlink
    let xlsx = create_xlsx_with_hyperlinks_for_rendering(vec![("A1", "Plain text", None)]);

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Sheet should have no hyperlinks
    assert!(
        sheet.hyperlinks.is_empty(),
        "Sheet should have no hyperlinks"
    );

    // Cell should not have a hyperlink attached
    let cell_data = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");

    assert!(
        cell_data.cell.hyperlink.is_none(),
        "Cell without hyperlink should have None for hyperlink field"
    );
}

// ============================================================================
// Test 4: Multiple Cells with Different Hyperlinks
// ============================================================================

#[test]
fn test_multiple_cells_with_different_hyperlinks() {
    // Create XLSX with multiple cells having different types of hyperlinks
    let xlsx = create_xlsx_with_hyperlinks_for_rendering(vec![
        (
            "A1",
            "External Link",
            Some(HyperlinkInfo::external("https://example.com")),
        ),
        (
            "B1",
            "Internal Link",
            Some(HyperlinkInfo::internal("Sheet2!B5")),
        ),
        ("C1", "No Link", None),
        (
            "D1",
            "Another External",
            Some(HyperlinkInfo::external("https://anthropic.com")),
        ),
    ]);

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Sheet should have 3 hyperlinks (A1, B1, D1)
    assert_eq!(sheet.hyperlinks.len(), 3, "Sheet should have 3 hyperlinks");

    // Verify A1 - external hyperlink
    let a1_hyperlink = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "A1")
        .expect("A1 should have hyperlink");
    assert_eq!(a1_hyperlink.hyperlink.target, "https://example.com");
    assert!(a1_hyperlink.hyperlink.is_external);

    let a1_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 0)
        .expect("Cell A1 should exist");
    assert!(a1_cell.cell.hyperlink.is_some());
    assert_eq!(
        a1_cell.cell.hyperlink.as_ref().unwrap().target,
        "https://example.com"
    );

    // Verify B1 - internal hyperlink
    let b1_hyperlink = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "B1")
        .expect("B1 should have hyperlink");
    assert_eq!(b1_hyperlink.hyperlink.target, "Sheet2!B5");
    assert!(!b1_hyperlink.hyperlink.is_external);

    let b1_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 1)
        .expect("Cell B1 should exist");
    assert!(b1_cell.cell.hyperlink.is_some());
    assert!(!b1_cell.cell.hyperlink.as_ref().unwrap().is_external);

    // Verify C1 - no hyperlink
    let c1_has_hyperlink = sheet.hyperlinks.iter().any(|h| h.cell_ref == "C1");
    assert!(
        !c1_has_hyperlink,
        "C1 should not have hyperlink in sheet.hyperlinks"
    );

    let c1_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 2)
        .expect("Cell C1 should exist");
    assert!(
        c1_cell.cell.hyperlink.is_none(),
        "C1 should have None for hyperlink"
    );

    // Verify D1 - another external hyperlink
    let d1_hyperlink = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "D1")
        .expect("D1 should have hyperlink");
    assert_eq!(d1_hyperlink.hyperlink.target, "https://anthropic.com");
    assert!(d1_hyperlink.hyperlink.is_external);

    let d1_cell = sheet
        .cells
        .iter()
        .find(|c| c.r == 0 && c.c == 3)
        .expect("Cell D1 should exist");
    assert!(d1_cell.cell.hyperlink.is_some());
    assert_eq!(
        d1_cell.cell.hyperlink.as_ref().unwrap().target,
        "https://anthropic.com"
    );
}

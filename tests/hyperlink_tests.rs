//! Tests for hyperlink parsing in XLSX files.
//!
//! This module tests the parsing of both internal and external hyperlinks from XLSX files.
//! Hyperlinks can be:
//! - External URLs (http://, https://, mailto:, file://)
//! - Internal references to other sheets/cells (Sheet2!A1)
//! - Links with display text and tooltips
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
    clippy::redundant_clone
)]

mod common;
mod fixtures;

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Helper Functions for Creating XLSX with Hyperlinks
// ============================================================================

/// Create a minimal XLSX structure with hyperlinks.
fn create_xlsx_with_hyperlinks(
    hyperlinks_xml: &str,
    rels_xml: Option<&str>,
    shared_strings: &[&str],
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

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
    for s in shared_strings {
        sst.push_str(&format!("<si><t>{}</t></si>", s));
    }
    sst.push_str("</sst>");
    let _ = zip.write_all(sst.as_bytes());

    // xl/worksheets/_rels/sheet1.xml.rels (hyperlink targets) - optional
    if let Some(rels) = rels_xml {
        let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options.clone());
        let _ = zip.write_all(rels.as_bytes());
    }

    // xl/worksheets/sheet1.xml (with hyperlinks)
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options.clone());
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1">
{}</row>
</sheetData>
{}</worksheet>"#,
        (0..shared_strings.len())
            .map(|i| {
                let col = (b'A' + i as u8) as char;
                format!(r#"<c r="{}1" t="s"><v>{}</v></c>"#, col, i)
            })
            .collect::<Vec<_>>()
            .join("\n"),
        hyperlinks_xml
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// ============================================================================
// Test 1: External URL Hyperlinks (http://, https://)
// ============================================================================

#[test]
fn test_external_http_url_hyperlink() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Click Here"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 1);

    let link = &sheet.hyperlinks[0];
    assert_eq!(link.cell_ref, "A1");
    assert_eq!(link.hyperlink.target, "http://example.com");
    assert!(link.hyperlink.is_external);
}

#[test]
fn test_external_https_url_hyperlink() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secure.example.com/page" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Secure Link"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 1);

    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://secure.example.com/page");
    assert!(link.hyperlink.is_external);
}

// ============================================================================
// Test 2: Internal Sheet References (Sheet2!A1)
// ============================================================================

#[test]
fn test_internal_sheet_reference() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" location="Sheet2!A1"/>
</hyperlinks>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["Go to Sheet2"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 1);

    let link = &sheet.hyperlinks[0];
    assert_eq!(link.cell_ref, "A1");
    assert_eq!(link.hyperlink.target, "Sheet2!A1");
    assert!(!link.hyperlink.is_external);
    assert_eq!(link.hyperlink.location.as_deref(), Some("Sheet2!A1"));
}

#[test]
fn test_internal_sheet_reference_with_quotes() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" location="'My Sheet'!B5"/>
</hyperlinks>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["Go to My Sheet"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "'My Sheet'!B5");
    assert!(!link.hyperlink.is_external);
}

#[test]
fn test_internal_named_range_reference() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" location="MyNamedRange"/>
</hyperlinks>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["Go to Named Range"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "MyNamedRange");
    assert!(!link.hyperlink.is_external);
}

// ============================================================================
// Test 3: Email Hyperlinks (mailto:)
// ============================================================================

#[test]
fn test_mailto_hyperlink() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:test@example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Email Us"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "mailto:test@example.com");
    assert!(link.hyperlink.is_external);
}

#[test]
fn test_mailto_hyperlink_with_subject() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:support@example.com?subject=Help%20Request" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Contact Support"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(
        link.hyperlink.target,
        "mailto:support@example.com?subject=Help%20Request"
    );
    assert!(link.hyperlink.is_external);
}

// ============================================================================
// Test 4: File Hyperlinks
// ============================================================================

#[test]
fn test_file_hyperlink_relative() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../documents/report.pdf" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Open Report"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "../documents/report.pdf");
    assert!(link.hyperlink.is_external);
}

#[test]
fn test_file_hyperlink_absolute() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///C:/Users/Documents/file.xlsx" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Open File"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(
        link.hyperlink.target,
        "file:///C:/Users/Documents/file.xlsx"
    );
    assert!(link.hyperlink.is_external);
}

// ============================================================================
// Test 5: Hyperlinks with Display Text
// ============================================================================

#[test]
fn test_hyperlink_with_display_text() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" display="Click here to visit"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Link Text"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://example.com");
    // Display text is typically used for rendering, the cell value is separate
    assert!(link.hyperlink.is_external);
}

// ============================================================================
// Test 6: Hyperlinks with Tooltips
// ============================================================================

#[test]
fn test_hyperlink_with_tooltip() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" tooltip="Visit our website"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Website"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://example.com");
    assert_eq!(link.hyperlink.tooltip.as_deref(), Some("Visit our website"));
}

#[test]
fn test_hyperlink_with_display_and_tooltip() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" display="Anthropic" tooltip="Visit Anthropic's website"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://anthropic.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Link"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://anthropic.com");
    assert_eq!(
        link.hyperlink.tooltip.as_deref(),
        Some("Visit Anthropic's website")
    );
}

// ============================================================================
// Test 7: Multiple Hyperlinks in One Sheet
// ============================================================================

#[test]
fn test_multiple_hyperlinks_in_sheet() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" tooltip="External URL"/>
<hyperlink ref="B1" location="Sheet2!A1" tooltip="Internal link"/>
<hyperlink ref="C1" r:id="rId2" tooltip="Email link"/>
<hyperlink ref="D1" r:id="rId3"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:test@example.com" TargetMode="External"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://another.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(
        hyperlinks,
        Some(rels),
        &["Website", "Go to Sheet2", "Email", "Another"],
    );
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 4);

    // Verify each hyperlink
    let link_a1 = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "A1")
        .unwrap();
    assert_eq!(link_a1.hyperlink.target, "https://example.com");
    assert!(link_a1.hyperlink.is_external);
    assert_eq!(link_a1.hyperlink.tooltip.as_deref(), Some("External URL"));

    let link_b1 = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "B1")
        .unwrap();
    assert_eq!(link_b1.hyperlink.target, "Sheet2!A1");
    assert!(!link_b1.hyperlink.is_external);
    assert_eq!(link_b1.hyperlink.tooltip.as_deref(), Some("Internal link"));

    let link_c1 = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "C1")
        .unwrap();
    assert_eq!(link_c1.hyperlink.target, "mailto:test@example.com");
    assert!(link_c1.hyperlink.is_external);

    let link_d1 = sheet
        .hyperlinks
        .iter()
        .find(|h| h.cell_ref == "D1")
        .unwrap();
    assert_eq!(link_d1.hyperlink.target, "https://another.com");
    assert!(link_d1.hyperlink.is_external);
    assert!(link_d1.hyperlink.tooltip.is_none());
}

// ============================================================================
// Test 8: Hyperlinks with Special Characters in URL
// ============================================================================

#[test]
fn test_hyperlink_with_encoded_url() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/search?q=hello%20world&amp;page=1" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Search"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    // Note: XML entities like &amp; get decoded during parsing
    assert!(link.hyperlink.target.contains("example.com/search"));
    assert!(link.hyperlink.is_external);
}

#[test]
fn test_hyperlink_with_unicode_in_url() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/path/%E4%B8%AD%E6%96%87" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Chinese Path"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(
        link.hyperlink.target,
        "https://example.com/path/%E4%B8%AD%E6%96%87"
    );
}

#[test]
fn test_hyperlink_with_fragment() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/page#section-2" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Page Section"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://example.com/page#section-2");
}

// ============================================================================
// Additional Edge Cases
// ============================================================================

#[test]
fn test_no_hyperlinks() {
    let hyperlinks = "";
    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["Plain Text"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(sheet.hyperlinks.is_empty());
}

#[test]
fn test_empty_hyperlinks_element() {
    let hyperlinks = "<hyperlinks/>";
    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["Plain Text"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(sheet.hyperlinks.is_empty());
}

#[test]
fn test_hyperlink_missing_relationship() {
    // Hyperlink references rId99 which doesn't exist in rels
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId99"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Broken Link"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    // Hyperlink with missing relationship should be skipped
    assert!(
        sheet.hyperlinks.is_empty()
            || !sheet
                .hyperlinks
                .iter()
                .any(|h| h.cell_ref == "A1" && h.hyperlink.target == "https://example.com")
    );
}

#[test]
fn test_hyperlink_types_struct() {
    use xlview::types::{Hyperlink, HyperlinkDef};

    // Test creating Hyperlink structs directly
    let external = Hyperlink {
        target: "https://example.com".to_string(),
        location: None,
        tooltip: Some("Visit site".to_string()),
        is_external: true,
    };

    assert_eq!(external.target, "https://example.com");
    assert!(external.is_external);
    assert_eq!(external.tooltip.as_deref(), Some("Visit site"));
    assert!(external.location.is_none());

    let internal = Hyperlink {
        target: "Sheet2!A1".to_string(),
        location: Some("Sheet2!A1".to_string()),
        tooltip: None,
        is_external: false,
    };

    assert_eq!(internal.target, "Sheet2!A1");
    assert!(!internal.is_external);
    assert!(internal.tooltip.is_none());
    assert_eq!(internal.location.as_deref(), Some("Sheet2!A1"));

    let def = HyperlinkDef {
        cell_ref: "A1".to_string(),
        hyperlink: external.clone(),
    };

    assert_eq!(def.cell_ref, "A1");
    assert_eq!(def.hyperlink.target, "https://example.com");
}

#[test]
fn test_hyperlink_serialization_json() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" tooltip="Test tooltip"/>
<hyperlink ref="B1" location="Sheet2!A1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["External", "Internal"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    // Serialize to JSON
    let json = serde_json::to_value(&workbook).expect("Failed to serialize");

    let sheet = &json["sheets"][0];
    let hyperlinks_json = &sheet["hyperlinks"];

    assert!(hyperlinks_json.is_array());
    let arr = hyperlinks_json.as_array().unwrap();
    assert_eq!(arr.len(), 2);

    // Check structure of first hyperlink
    let first = &arr[0];
    assert!(first["cellRef"].is_string());
    assert!(first["hyperlink"]["target"].is_string());
    assert!(first["hyperlink"]["isExternal"].is_boolean());
}

#[test]
fn test_hyperlink_with_both_rid_and_location() {
    // External link with additional location/anchor
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1" location="section1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/page" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["Page with anchor"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.hyperlink.target, "https://example.com/page");
    assert!(link.hyperlink.is_external);
    assert_eq!(link.hyperlink.location.as_deref(), Some("section1"));
}

#[test]
fn test_hyperlink_on_various_cell_positions() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" location="Target1"/>
<hyperlink ref="Z1" location="Target2"/>
<hyperlink ref="AA1" location="Target3"/>
<hyperlink ref="A100" location="Target4"/>
</hyperlinks>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, None, &["A1", "Z1", "AA1", "A100"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.hyperlinks.len(), 4);

    // Verify cell references are preserved
    let refs: Vec<&str> = sheet
        .hyperlinks
        .iter()
        .map(|h| h.cell_ref.as_str())
        .collect();
    assert!(refs.contains(&"A1"));
    assert!(refs.contains(&"Z1"));
    assert!(refs.contains(&"AA1"));
    assert!(refs.contains(&"A100"));
}

#[test]
fn test_ftp_hyperlink() {
    let hyperlinks = r#"<hyperlinks>
<hyperlink ref="A1" r:id="rId1"/>
</hyperlinks>"#;

    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="ftp://ftp.example.com/files/data.zip" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_xlsx_with_hyperlinks(hyperlinks, Some(rels), &["FTP Download"]);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let link = &sheet.hyperlinks[0];
    assert_eq!(
        link.hyperlink.target,
        "ftp://ftp.example.com/files/data.zip"
    );
    assert!(link.hyperlink.is_external);
}

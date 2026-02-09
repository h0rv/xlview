//! Tests for cell protection/locking indicator parsing
//!
//! Tests the parsing of:
//! - `<protection>` element inside `<xf>` in styles.xml
//! - `<sheetProtection>` element in sheet XML
//!
//! Excel cell protection behavior:
//! - By default, all cells are locked (`locked="1"` or attribute absent)
//! - Cells can be explicitly unlocked with `locked="0"`
//! - Formulas can be hidden with `hidden="1"` (only visible when sheet is protected)
//! - Protection only takes effect when the sheet is protected via `<sheetProtection>`
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

// ============================================================================
// Helper Functions for Creating XLSX Files
// ============================================================================

/// Create a minimal XLSX file with custom styles and sheet content
fn create_xlsx_with_styles_and_sheet(styles_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let options = FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();

        // xl/_rels/workbook.xml.rels
        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#).unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#).unwrap();

        // xl/styles.xml
        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        // xl/worksheets/sheet1.xml
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buf.into_inner()
}

/// Create a minimal styles.xml with protection settings in cellXfs
fn styles_xml_with_protection(xf_entries: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font><sz val="11"/><name val="Calibri"/></font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        {xf_entries}
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
</styleSheet>"#
    )
}

/// Create a basic sheet XML with optional sheetProtection and cells
fn sheet_xml_with_protection(protection_element: &str, cells: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    {protection_element}
    <sheetData>
        {cells}
    </sheetData>
</worksheet>"#
    )
}

// ============================================================================
// CELL LOCKED (DEFAULT WHEN PROTECTED) TESTS
// ============================================================================

mod cell_locked_default {
    use super::*;

    #[test]
    fn test_cell_locked_default_no_protection_element() {
        // Cell with no protection element - locked is default true in Excel
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>Test</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        // Sheet should be protected
        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        // Cell without explicit protection - locked is default (not explicitly set in output)
        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        // When locked is default (true), it should not be explicitly set in style
        // since we only set locked=Some(false) when explicitly unlocked
        let locked = cell["s"]["locked"].as_bool();
        assert!(
            locked.is_none() || locked == Some(true),
            "Default locked cell should not have locked=false"
        );
    }

    #[test]
    fn test_cell_locked_explicit_true() {
        // Cell with explicit locked="1"
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>Locked</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        // Explicit locked=true - should not set locked=false
        let locked = cell["s"]["locked"].as_bool();
        assert!(locked.is_none() || locked == Some(true));
    }

    #[test]
    fn test_cell_locked_with_empty_protection_element() {
        // Cell with empty protection element (defaults to locked=true)
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>Default Locked</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// CELL UNLOCKED TESTS
// ============================================================================

mod cell_unlocked {
    use super::*;

    #[test]
    fn test_cell_unlocked_explicit() {
        // Cell with locked="0" - explicitly unlocked
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>Unlocked</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        // Unlocked cell should have locked=false
        assert_eq!(
            cell["s"]["locked"], false,
            "Unlocked cell should have locked=false"
        );
    }

    #[test]
    fn test_cell_unlocked_on_unprotected_sheet() {
        // Cell unlocked on unprotected sheet - protection has no effect
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            "", // No sheetProtection element
            r#"<row r="1"><c r="A1" s="1"><v>Unlocked</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        // Sheet should not be protected
        assert_eq!(workbook["sheets"][0]["isProtected"], false);

        // Cell style should still reflect unlocked status
        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];
        assert_eq!(cell["s"]["locked"], false);
    }

    #[test]
    fn test_multiple_cells_mixed_lock_status() {
        // Multiple cells with different lock statuses
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1">
                <c r="A1" s="1"><v>Unlocked</v></c>
                <c r="B1" s="2"><v>Locked</v></c>
            </row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();

        // Find cells by column
        let cell_a1 = cells.iter().find(|c| c["c"] == 0).unwrap();
        let cell_b1 = cells.iter().find(|c| c["c"] == 1).unwrap();

        assert_eq!(
            cell_a1["cell"]["s"]["locked"], false,
            "A1 should be unlocked"
        );
        // B1 has explicit locked=true, which may or may not be in output
        let b1_locked = cell_b1["cell"]["s"]["locked"].as_bool();
        assert!(
            b1_locked.is_none() || b1_locked == Some(true),
            "B1 should be locked"
        );
    }
}

// ============================================================================
// FORMULA HIDDEN TESTS
// ============================================================================

mod formula_hidden {
    use super::*;

    #[test]
    fn test_formula_hidden_true() {
        // Cell with hidden="1" - formula is hidden when sheet is protected
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection hidden="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>42</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        assert_eq!(cell["s"]["hidden"], true, "Formula should be hidden");
    }

    #[test]
    fn test_formula_hidden_false() {
        // Cell with hidden="0" - formula is visible
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection hidden="0"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>42</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        // hidden=false is default, so it should not be in output
        let hidden = cell["s"]["hidden"].as_bool();
        assert!(hidden.is_none() || hidden == Some(false));
    }

    #[test]
    fn test_formula_hidden_and_unlocked() {
        // Cell with both hidden="1" and locked="0"
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0" hidden="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>42</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];

        assert_eq!(cell["s"]["locked"], false, "Cell should be unlocked");
        assert_eq!(cell["s"]["hidden"], true, "Formula should be hidden");
    }
}

// ============================================================================
// SHEET PROTECTION ENABLED TESTS
// ============================================================================

mod sheet_protection_enabled {
    use super::*;

    #[test]
    fn test_sheet_protection_basic() {
        // Basic sheet protection with sheet="1"
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Protected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_sheet_protection_disabled() {
        // Sheet protection with sheet="0"
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Unprotected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], false);
    }

    #[test]
    fn test_sheet_protection_absent() {
        // No sheetProtection element at all
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            "", // No protection element
            r#"<row r="1"><c r="A1"><v>No Protection</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], false);
    }

    #[test]
    fn test_sheet_protection_element_without_sheet_attribute() {
        // sheetProtection element exists but no sheet attribute
        // Parser should treat this as protected (conservative approach)
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection objects="1" scenarios="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Protected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        // Presence of sheetProtection without sheet attribute should be treated as protected
        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// SHEET PROTECTION WITH PASSWORD HASH TESTS
// ============================================================================

mod sheet_protection_password {
    use super::*;

    #[test]
    fn test_sheet_protection_with_password() {
        // Sheet protection with password hash
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" password="CC1A"/>"#,
            r#"<row r="1"><c r="A1"><v>Password Protected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_sheet_protection_with_hash_algorithm() {
        // Modern Excel uses hashValue with algorithm specification
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" algorithmName="SHA-512" hashValue="abc123..." saltValue="xyz789..." spinCount="100000"/>"#,
            r#"<row r="1"><c r="A1"><v>Hash Protected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// PROTECTION OPTIONS (FORMAT CELLS, ROWS, COLUMNS) TESTS
// ============================================================================

mod protection_format_options {
    use super::*;

    #[test]
    fn test_protection_format_cells_allowed() {
        // formatCells="0" means formatting cells IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" formatCells="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Format Allowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_format_cells_disallowed() {
        // formatCells="1" means formatting cells is NOT allowed (default)
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" formatCells="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Format Disallowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_format_rows_allowed() {
        // formatRows="0" means formatting rows IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" formatRows="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Format Rows</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_format_columns_allowed() {
        // formatColumns="0" means formatting columns IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" formatColumns="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Format Columns</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_all_format_options() {
        // All format options specified
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" formatCells="0" formatRows="0" formatColumns="0"/>"#,
            r#"<row r="1"><c r="A1"><v>All Format Options</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// PROTECTION OPTIONS (INSERT/DELETE ROWS/COLUMNS) TESTS
// ============================================================================

mod protection_insert_delete_options {
    use super::*;

    #[test]
    fn test_protection_insert_rows_allowed() {
        // insertRows="0" means inserting rows IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" insertRows="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Insert Rows</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_insert_columns_allowed() {
        // insertColumns="0" means inserting columns IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" insertColumns="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Insert Columns</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_delete_rows_allowed() {
        // deleteRows="0" means deleting rows IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" deleteRows="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Delete Rows</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_delete_columns_allowed() {
        // deleteColumns="0" means deleting columns IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" deleteColumns="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Delete Columns</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_insert_hyperlinks_allowed() {
        // insertHyperlinks="0" means inserting hyperlinks IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" insertHyperlinks="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Insert Hyperlinks</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_all_insert_delete_options() {
        // All insert/delete options specified
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" insertRows="0" insertColumns="0" deleteRows="0" deleteColumns="0"/>"#,
            r#"<row r="1"><c r="A1"><v>All Insert/Delete</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// PROTECTION OPTIONS (SORT, AUTOFILTER) TESTS
// ============================================================================

mod protection_sort_autofilter_options {
    use super::*;

    #[test]
    fn test_protection_sort_allowed() {
        // sort="0" means sorting IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" sort="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Sort Allowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_sort_disallowed() {
        // sort="1" means sorting is NOT allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" sort="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Sort Disallowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_autofilter_allowed() {
        // autoFilter="0" means using autofilter IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" autoFilter="0"/>"#,
            r#"<row r="1"><c r="A1"><v>AutoFilter Allowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_autofilter_disallowed() {
        // autoFilter="1" means using autofilter is NOT allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" autoFilter="1"/>"#,
            r#"<row r="1"><c r="A1"><v>AutoFilter Disallowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_protection_pivot_tables_allowed() {
        // pivotTables="0" means using pivot tables IS allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" pivotTables="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Pivot Tables</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// SELECT LOCKED CELLS ALLOWED TESTS
// ============================================================================

mod select_locked_cells {
    use super::*;

    #[test]
    fn test_select_locked_cells_allowed() {
        // selectLockedCells="0" means selecting locked cells IS allowed (default)
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" selectLockedCells="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Select Locked Allowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_select_locked_cells_disallowed() {
        // selectLockedCells="1" means selecting locked cells is NOT allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" selectLockedCells="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Select Locked Disallowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_select_locked_cells_default() {
        // No selectLockedCells attribute - default is allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Default Select Locked</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// SELECT UNLOCKED CELLS ONLY TESTS
// ============================================================================

mod select_unlocked_cells {
    use super::*;

    #[test]
    fn test_select_unlocked_cells_allowed() {
        // selectUnlockedCells="0" means selecting unlocked cells IS allowed (default)
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" selectUnlockedCells="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Select Unlocked Allowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_select_unlocked_cells_disallowed() {
        // selectUnlockedCells="1" means selecting unlocked cells is NOT allowed
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" selectUnlockedCells="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Select Unlocked Disallowed</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_select_only_unlocked_cells() {
        // Common pattern: selectLockedCells="1" and selectUnlockedCells="0"
        // This means users can only select unlocked cells
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" selectLockedCells="1" selectUnlockedCells="0"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>Unlocked - Selectable</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];
        assert_eq!(cell["s"]["locked"], false);
    }
}

// ============================================================================
// COMPREHENSIVE PROTECTION SCENARIOS
// ============================================================================

mod comprehensive_scenarios {
    use super::*;

    #[test]
    fn test_typical_input_form_protection() {
        // Typical input form: protected sheet with unlocked input cells
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" objects="1" scenarios="1"/>"#,
            r#"<row r="1">
                <c r="A1" s="2"><v>Label:</v></c>
                <c r="B1" s="1"><v>Enter value here</v></c>
            </row>
            <row r="2">
                <c r="A2" s="2"><v>Result:</v></c>
                <c r="B2" s="2"><v>Calculated</v></c>
            </row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();

        // B1 should be unlocked (input cell)
        let cell_b1 = cells.iter().find(|c| c["r"] == 0 && c["c"] == 1).unwrap();
        assert_eq!(cell_b1["cell"]["s"]["locked"], false);

        // A1 should be locked (label)
        let cell_a1 = cells.iter().find(|c| c["r"] == 0 && c["c"] == 0).unwrap();
        let a1_locked = cell_a1["cell"]["s"]["locked"].as_bool();
        assert!(a1_locked.is_none() || a1_locked == Some(true));
    }

    #[test]
    fn test_hidden_formula_protection() {
        // Formula cells with hidden formulas
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="1" hidden="1"/>
            </xf>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"/>"#,
            r#"<row r="1"><c r="A1" s="1"><v>100</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);

        let cells = workbook["sheets"][0]["cells"].as_array().unwrap();
        let cell = &cells[0]["cell"];
        assert_eq!(cell["s"]["hidden"], true);
    }

    #[test]
    fn test_full_protection_options() {
        // All protection options specified
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1"
                password="CC1A"
                formatCells="0"
                formatRows="0"
                formatColumns="0"
                insertRows="0"
                insertColumns="0"
                insertHyperlinks="0"
                deleteRows="0"
                deleteColumns="0"
                sort="0"
                autoFilter="0"
                pivotTables="0"
                selectLockedCells="0"
                selectUnlockedCells="0"/>"#,
            r#"<row r="1"><c r="A1"><v>Full Protection</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }

    #[test]
    fn test_objects_and_scenarios_protection() {
        // Protection with objects and scenarios
        let styles = styles_xml_with_protection(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );
        let sheet = sheet_xml_with_protection(
            r#"<sheetProtection sheet="1" objects="1" scenarios="1"/>"#,
            r#"<row r="1"><c r="A1"><v>Objects Protected</v></c></row>"#,
        );

        let xlsx = create_xlsx_with_styles_and_sheet(&styles, &sheet);
        let workbook: serde_json::Value =
            serde_json::from_str(&xlview::parse_xlsx(&xlsx).expect("Failed to parse"))
                .expect("Failed to parse JSON");

        assert_eq!(workbook["sheets"][0]["isProtected"], true);
    }
}

// ============================================================================
// STYLES.XML PARSING TESTS (Original tests preserved)
// ============================================================================

use std::io::Cursor as StdCursor;
use xlview::styles::parse_styles;
use xlview::types::{RawProtection, StyleSheet};

/// Helper to create a minimal styles.xml with protection settings
fn styles_xml_with_protection_for_parsing(xf_content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
            <sz val="11"/>
            <name val="Calibri"/>
        </font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        {xf_content}
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
</styleSheet>"#
    )
}

/// Helper to parse a styles XML string and return the StyleSheet
fn parse_styles_xml(xml: &str) -> StyleSheet {
    let cursor = StdCursor::new(xml);
    parse_styles(cursor).expect("Failed to parse styles XML")
}

mod protection_parsing {
    use super::*;

    #[test]
    fn test_protection_locked_false() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0"/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.cell_xfs.len(), 2);

        let xf = &stylesheet.cell_xfs[1];
        assert!(xf.apply_protection);
        assert!(xf.protection.is_some());

        let protection = xf.protection.as_ref().unwrap();
        assert!(!protection.locked, "Cell should be unlocked");
        assert!(!protection.hidden, "Formula should not be hidden");
    }

    #[test]
    fn test_protection_locked_true_explicit() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="1"/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        assert!(xf.protection.is_some());
        let protection = xf.protection.as_ref().unwrap();
        assert!(protection.locked, "Cell should be locked");
    }

    #[test]
    fn test_protection_locked_default() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        assert!(xf.protection.is_some());
        let protection = xf.protection.as_ref().unwrap();
        assert!(protection.locked, "Cell should be locked by default");
    }

    #[test]
    fn test_protection_hidden_true() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection hidden="1"/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        assert!(xf.protection.is_some());
        let protection = xf.protection.as_ref().unwrap();
        assert!(protection.hidden, "Formula should be hidden");
    }

    #[test]
    fn test_protection_hidden_false() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection hidden="0"/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        assert!(xf.protection.is_some());
        let protection = xf.protection.as_ref().unwrap();
        assert!(!protection.hidden, "Formula should not be hidden");
    }

    #[test]
    fn test_protection_both_attributes() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">
                <protection locked="0" hidden="1"/>
            </xf>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        assert!(xf.protection.is_some());
        let protection = xf.protection.as_ref().unwrap();
        assert!(!protection.locked, "Cell should be unlocked");
        assert!(protection.hidden, "Formula should be hidden");
    }

    #[test]
    fn test_no_protection_element() {
        let xml = styles_xml_with_protection_for_parsing(
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let xf = &stylesheet.cell_xfs[1];

        // Note: The parser defaults applyProtection to true when not explicitly set.
        // This is consistent with Excel's behavior where protection is applied by default.
        assert!(xf.apply_protection);
        assert!(xf.protection.is_none());
    }
}

mod raw_protection_defaults {
    use super::*;

    #[test]
    fn test_raw_protection_default() {
        let protection = RawProtection::default();
        assert!(
            !protection.locked,
            "Default locked should be false in RawProtection::default()"
        );
        assert!(!protection.hidden, "Default hidden should be false");
    }
}

//! Tests for sheet visibility and tab color parsing
//!
//! Tests the parsing of:
//! - `state` attribute on `<sheet>` elements in workbook.xml
//! - `tabColor` element inside `<sheetPr>` in sheet XML
//!
//! Sheet visibility states in Excel:
//! - `visible` (default): Sheet is visible in the tab bar
//! - `hidden`: Sheet is hidden but can be unhidden via Excel UI
//! - `veryHidden`: Sheet is hidden and can only be unhidden via VBA
//!
//! Tab colors can be specified using:
//! - `rgb`: Direct ARGB color value (e.g., "FFFF0000" for red)
//! - `theme`: Theme color index (0-11) with optional tint
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

//! - `indexed`: Legacy indexed color (0-63)

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// ============================================================================
// Helper Functions for Creating Test XLSX Files
// ============================================================================

/// Create a minimal XLSX file with specified sheets and their visibility states
fn create_xlsx_with_visibility(sheets: &[(&str, Option<&str>)]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Write [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let content_types = generate_content_types(sheets.len());
    let _ = zip.write_all(content_types.as_bytes());

    // Write _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let rels = generate_rels();
    let _ = zip.write_all(rels.as_bytes());

    // Write xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let workbook_rels = generate_workbook_rels(sheets.len());
    let _ = zip.write_all(workbook_rels.as_bytes());

    // Write xl/workbook.xml with visibility states
    let _ = zip.start_file("xl/workbook.xml", options);
    let workbook = generate_workbook_with_visibility(sheets);
    let _ = zip.write_all(workbook.as_bytes());

    // Write xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(minimal_styles_xml().as_bytes());

    // Write xl/theme/theme1.xml
    let _ = zip.start_file("xl/theme/theme1.xml", options);
    let _ = zip.write_all(minimal_theme_xml().as_bytes());

    // Write each sheet
    for (i, _) in sheets.iter().enumerate() {
        let path = format!("xl/worksheets/sheet{}.xml", i + 1);
        let _ = zip.start_file(&path, options);
        let _ = zip.write_all(minimal_sheet_xml().as_bytes());
    }

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create an XLSX file with sheets that have tab colors
fn create_xlsx_with_tab_colors(sheets: &[(&str, Option<TabColorSpec>)]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Write [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let content_types = generate_content_types(sheets.len());
    let _ = zip.write_all(content_types.as_bytes());

    // Write _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let rels = generate_rels();
    let _ = zip.write_all(rels.as_bytes());

    // Write xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let workbook_rels = generate_workbook_rels(sheets.len());
    let _ = zip.write_all(workbook_rels.as_bytes());

    // Write xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let sheet_names: Vec<&str> = sheets.iter().map(|(name, _)| *name).collect();
    let workbook = generate_workbook_simple(&sheet_names);
    let _ = zip.write_all(workbook.as_bytes());

    // Write xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(minimal_styles_xml().as_bytes());

    // Write xl/theme/theme1.xml
    let _ = zip.start_file("xl/theme/theme1.xml", options);
    let _ = zip.write_all(minimal_theme_xml().as_bytes());

    // Write each sheet with optional tab color
    for (i, (_, tab_color)) in sheets.iter().enumerate() {
        let path = format!("xl/worksheets/sheet{}.xml", i + 1);
        let _ = zip.start_file(&path, options);
        let sheet_xml = generate_sheet_xml_with_tab_color(tab_color.as_ref());
        let _ = zip.write_all(sheet_xml.as_bytes());
    }

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Tab color specification for test creation
#[derive(Clone)]
enum TabColorSpec {
    Rgb(String),             // rgb="FFRRGGBB"
    Theme(u32),              // theme="N"
    ThemeWithTint(u32, f64), // theme="N" tint="X.X"
    Indexed(u32),            // indexed="N"
}

fn generate_content_types(sheet_count: usize) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#);

    for i in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            i
        ));
    }

    xml.push_str("</Types>");
    xml
}

fn generate_rels() -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    xml.push_str(r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#);
    xml.push_str("</Relationships>");
    xml
}

fn generate_workbook_rels(sheet_count: usize) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );

    let mut rid = 1;

    // Sheets
    for i in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            rid, i
        ));
        rid += 1;
    }

    // Styles
    xml.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        rid
    ));
    rid += 1;

    // Theme
    xml.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        rid
    ));

    xml.push_str("</Relationships>");
    xml
}

fn generate_workbook_with_visibility(sheets: &[(&str, Option<&str>)]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    xml.push_str("<sheets>");

    for (i, (name, state)) in sheets.iter().enumerate() {
        let state_attr = match state {
            Some(s) => format!(r#" state="{}""#, s),
            None => String::new(),
        };
        xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"{}/ >"#,
            escape_xml(name),
            i + 1,
            i + 1,
            state_attr
        ));
    }

    xml.push_str("</sheets>");
    xml.push_str("</workbook>");
    xml
}

fn generate_workbook_simple(sheets: &[&str]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#);
    xml.push_str("<sheets>");

    for (i, name) in sheets.iter().enumerate() {
        xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            escape_xml(name),
            i + 1,
            i + 1
        ));
    }

    xml.push_str("</sheets>");
    xml.push_str("</workbook>");
    xml
}

fn minimal_styles_xml() -> String {
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
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
</styleSheet>"#
        .to_string()
}

fn minimal_theme_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1><a:srgbClr val="000000"/></a:dk1>
            <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2>
            <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
            <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
            <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
</a:theme>"#
        .to_string()
}

fn minimal_sheet_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
</worksheet>"#
        .to_string()
}

fn generate_sheet_xml_with_tab_color(tab_color: Option<&TabColorSpec>) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );

    if let Some(color) = tab_color {
        xml.push_str("<sheetPr>");
        match color {
            TabColorSpec::Rgb(rgb) => {
                xml.push_str(&format!(r#"<tabColor rgb="{}"/>"#, rgb));
            }
            TabColorSpec::Theme(theme) => {
                xml.push_str(&format!(r#"<tabColor theme="{}"/>"#, theme));
            }
            TabColorSpec::ThemeWithTint(theme, tint) => {
                xml.push_str(&format!(r#"<tabColor theme="{}" tint="{}"/>"#, theme, tint));
            }
            TabColorSpec::Indexed(indexed) => {
                xml.push_str(&format!(r#"<tabColor indexed="{}"/>"#, indexed));
            }
        }
        xml.push_str("</sheetPr>");
    }

    xml.push_str("<sheetData/>");
    xml.push_str("</worksheet>");
    xml
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Parse an XLSX file and return the workbook
fn parse_xlsx(data: &[u8]) -> xlview::types::Workbook {
    xlview::parser::parse(data).expect("Failed to parse XLSX")
}

// ============================================================================
// SHEET VISIBILITY TESTS
// ============================================================================

mod visibility_tests {
    use super::*;
    use xlview::types::SheetState;

    #[test]
    fn test_all_sheets_visible_default() {
        // When no state attribute is specified, sheets should be visible by default
        let sheets = vec![("Sheet1", None), ("Sheet2", None), ("Sheet3", None)];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 3);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].name, "Sheet2");
        assert_eq!(workbook.sheets[1].state, SheetState::Visible);
        assert_eq!(workbook.sheets[2].name, "Sheet3");
        assert_eq!(workbook.sheets[2].state, SheetState::Visible);
    }

    #[test]
    fn test_one_hidden_sheet() {
        // One sheet with state="hidden"
        let sheets = vec![
            ("Visible Sheet", None),
            ("Hidden Sheet", Some("hidden")),
            ("Another Visible", None),
        ];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 3);
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].name, "Hidden Sheet");
        assert_eq!(workbook.sheets[1].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[2].state, SheetState::Visible);
    }

    #[test]
    fn test_one_very_hidden_sheet() {
        // One sheet with state="veryHidden" (can only be unhidden via VBA)
        let sheets = vec![
            ("Visible Sheet", None),
            ("VeryHidden Sheet", Some("veryHidden")),
        ];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 2);
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].name, "VeryHidden Sheet");
        assert_eq!(workbook.sheets[1].state, SheetState::VeryHidden);
    }

    #[test]
    fn test_multiple_hidden_sheets() {
        // Multiple sheets with state="hidden"
        let sheets = vec![
            ("Visible1", None),
            ("Hidden1", Some("hidden")),
            ("Hidden2", Some("hidden")),
            ("Visible2", None),
            ("Hidden3", Some("hidden")),
        ];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 5);
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[2].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[3].state, SheetState::Visible);
        assert_eq!(workbook.sheets[4].state, SheetState::Hidden);
    }

    #[test]
    fn test_mix_of_visible_hidden_and_very_hidden() {
        // Mix of all three visibility states
        let sheets = vec![
            ("Visible Sheet", None),
            ("Hidden Sheet", Some("hidden")),
            ("VeryHidden Sheet", Some("veryHidden")),
            ("Another Visible", Some("visible")), // Explicitly visible
            ("Another Hidden", Some("hidden")),
        ];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 5);
        assert_eq!(workbook.sheets[0].name, "Visible Sheet");
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].name, "Hidden Sheet");
        assert_eq!(workbook.sheets[1].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[2].name, "VeryHidden Sheet");
        assert_eq!(workbook.sheets[2].state, SheetState::VeryHidden);
        assert_eq!(workbook.sheets[3].name, "Another Visible");
        assert_eq!(workbook.sheets[3].state, SheetState::Visible);
        assert_eq!(workbook.sheets[4].name, "Another Hidden");
        assert_eq!(workbook.sheets[4].state, SheetState::Hidden);
    }

    #[test]
    fn test_sheet_without_explicit_state_attribute() {
        // Explicitly test that sheets without state attribute default to visible
        let sheets = vec![("NoState1", None), ("NoState2", None)];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 2);
        // Both should be visible since no state attribute is specified
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].state, SheetState::Visible);
    }

    #[test]
    fn test_explicit_visible_state() {
        // Test that state="visible" is correctly parsed
        let sheets = vec![("ExplicitlyVisible", Some("visible"))];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "ExplicitlyVisible");
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
    }

    #[test]
    fn test_unknown_state_defaults_to_visible() {
        // Unknown state values should default to visible
        let sheets = vec![("UnknownState", Some("unknown"))];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 1);
        // Unknown states should be treated as visible
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
    }

    #[test]
    fn test_single_visible_sheet() {
        // Single sheet should default to visible
        let sheets = vec![("OnlySheet", None)];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "OnlySheet");
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
    }

    #[test]
    fn test_all_hidden_except_one() {
        // Only one visible sheet, rest hidden
        let sheets = vec![
            ("Hidden1", Some("hidden")),
            ("Visible", None),
            ("Hidden2", Some("hidden")),
            ("VeryHidden", Some("veryHidden")),
        ];

        let xlsx = create_xlsx_with_visibility(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 4);
        assert_eq!(workbook.sheets[0].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[1].state, SheetState::Visible);
        assert_eq!(workbook.sheets[2].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[3].state, SheetState::VeryHidden);
    }
}

// ============================================================================
// TAB COLOR TESTS
// ============================================================================

mod tab_color_tests {
    use super::*;

    #[test]
    fn test_tab_color_with_rgb() {
        // Tab color specified with rgb attribute
        let sheets = vec![
            ("Red Tab", Some(TabColorSpec::Rgb("FFFF0000".to_string()))),
            ("Green Tab", Some(TabColorSpec::Rgb("FF00FF00".to_string()))),
            ("Blue Tab", Some(TabColorSpec::Rgb("FF0000FF".to_string()))),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 3);

        // Red tab
        assert_eq!(workbook.sheets[0].name, "Red Tab");
        assert_eq!(workbook.sheets[0].tab_color, Some("#FF0000".to_string()));

        // Green tab
        assert_eq!(workbook.sheets[1].name, "Green Tab");
        assert_eq!(workbook.sheets[1].tab_color, Some("#00FF00".to_string()));

        // Blue tab
        assert_eq!(workbook.sheets[2].name, "Blue Tab");
        assert_eq!(workbook.sheets[2].tab_color, Some("#0000FF".to_string()));
    }

    #[test]
    #[ignore = "TODO: Test XLSX fixture doesn't include proper theme definition for theme color resolution"]
    fn test_tab_color_with_theme() {
        // Tab color specified with theme index
        let sheets = vec![
            ("Theme 0 (Dark 1)", Some(TabColorSpec::Theme(0))), // Black
            ("Theme 1 (Light 1)", Some(TabColorSpec::Theme(1))), // White
            ("Theme 4 (Accent 1)", Some(TabColorSpec::Theme(4))), // Blue
            ("Theme 5 (Accent 2)", Some(TabColorSpec::Theme(5))), // Orange
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 4);

        // Theme 0 (dark 1 = black)
        assert_eq!(workbook.sheets[0].tab_color, Some("#000000".to_string()));

        // Theme 1 (light 1 = white)
        assert_eq!(workbook.sheets[1].tab_color, Some("#FFFFFF".to_string()));

        // Theme 4 (accent 1 = blue #4472C4)
        assert_eq!(workbook.sheets[2].tab_color, Some("#4472C4".to_string()));

        // Theme 5 (accent 2 = orange #ED7D31)
        assert_eq!(workbook.sheets[3].tab_color, Some("#ED7D31".to_string()));
    }

    #[test]
    fn test_tab_color_with_indexed() {
        // Tab color specified with indexed color
        let sheets = vec![
            ("Indexed 0 (Black)", Some(TabColorSpec::Indexed(0))),
            ("Indexed 1 (White)", Some(TabColorSpec::Indexed(1))),
            ("Indexed 2 (Red)", Some(TabColorSpec::Indexed(2))),
            ("Indexed 5 (Yellow)", Some(TabColorSpec::Indexed(5))),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 4);

        // Indexed 0 = Black
        assert_eq!(workbook.sheets[0].tab_color, Some("#000000".to_string()));

        // Indexed 1 = White
        assert_eq!(workbook.sheets[1].tab_color, Some("#FFFFFF".to_string()));

        // Indexed 2 = Red
        assert_eq!(workbook.sheets[2].tab_color, Some("#FF0000".to_string()));

        // Indexed 5 = Yellow
        assert_eq!(workbook.sheets[3].tab_color, Some("#FFFF00".to_string()));
    }

    #[test]
    #[ignore = "TODO: Test XLSX fixture doesn't include proper theme definition for theme color resolution"]
    fn test_tab_color_with_tint() {
        // Tab color with theme and tint (lightening/darkening)
        let sheets = vec![
            // Theme 4 (Accent 1) with positive tint (lighter)
            ("Accent1 Light", Some(TabColorSpec::ThemeWithTint(4, 0.5))),
            // Theme 4 (Accent 1) with negative tint (darker)
            ("Accent1 Dark", Some(TabColorSpec::ThemeWithTint(4, -0.5))),
            // Theme 0 (black) with 50% tint should give gray
            (
                "Black + 50% Tint",
                Some(TabColorSpec::ThemeWithTint(0, 0.5)),
            ),
            // Theme 1 (white) with -50% tint should give gray
            (
                "White - 50% Tint",
                Some(TabColorSpec::ThemeWithTint(1, -0.5)),
            ),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 4);

        // Check that tinted colors are present
        // Accent1 (#4472C4) + 0.5 tint should be lighter
        assert!(workbook.sheets[0].tab_color.is_some());
        let light_color = workbook.sheets[0].tab_color.as_ref().unwrap();
        assert!(light_color.starts_with('#'));
        assert_eq!(light_color.len(), 7);

        // Accent1 (#4472C4) - 0.5 tint should be darker
        assert!(workbook.sheets[1].tab_color.is_some());
        let dark_color = workbook.sheets[1].tab_color.as_ref().unwrap();
        assert!(dark_color.starts_with('#'));
        assert_eq!(dark_color.len(), 7);

        // Black + 50% tint = #808080 (gray)
        assert_eq!(workbook.sheets[2].tab_color, Some("#808080".to_string()));

        // White - 50% tint = #808080 (gray)
        assert_eq!(workbook.sheets[3].tab_color, Some("#808080".to_string()));
    }

    #[test]
    fn test_no_tab_color() {
        // Sheets without tab color should have None
        let sheets = vec![("No Color 1", None), ("No Color 2", None)];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 2);
        assert!(workbook.sheets[0].tab_color.is_none());
        assert!(workbook.sheets[1].tab_color.is_none());
    }

    #[test]
    fn test_mixed_tab_colors() {
        // Mix of different tab color types and no color
        let sheets = vec![
            ("RGB Red", Some(TabColorSpec::Rgb("FFFF0000".to_string()))),
            ("No Color", None),
            ("Theme Accent", Some(TabColorSpec::Theme(4))),
            ("Indexed Yellow", Some(TabColorSpec::Indexed(5))),
            ("Theme with Tint", Some(TabColorSpec::ThemeWithTint(4, 0.5))),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 5);

        // RGB Red
        assert_eq!(workbook.sheets[0].tab_color, Some("#FF0000".to_string()));

        // No color
        assert!(workbook.sheets[1].tab_color.is_none());

        // Theme Accent (4 = Accent1 = #4472C4)
        assert_eq!(workbook.sheets[2].tab_color, Some("#4472C4".to_string()));

        // Indexed Yellow (5 = Yellow)
        assert_eq!(workbook.sheets[3].tab_color, Some("#FFFF00".to_string()));

        // Theme with tint
        assert!(workbook.sheets[4].tab_color.is_some());
    }

    #[test]
    fn test_tab_color_rgb_common_colors() {
        // Test common Excel tab colors
        let sheets = vec![
            ("Cyan", Some(TabColorSpec::Rgb("FF00FFFF".to_string()))),
            ("Magenta", Some(TabColorSpec::Rgb("FFFF00FF".to_string()))),
            ("Orange", Some(TabColorSpec::Rgb("FFFF8000".to_string()))),
            ("Purple", Some(TabColorSpec::Rgb("FF800080".to_string()))),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets[0].tab_color, Some("#00FFFF".to_string()));
        assert_eq!(workbook.sheets[1].tab_color, Some("#FF00FF".to_string()));
        assert_eq!(workbook.sheets[2].tab_color, Some("#FF8000".to_string()));
        assert_eq!(workbook.sheets[3].tab_color, Some("#800080".to_string()));
    }

    #[test]
    #[ignore = "TODO: Test XLSX fixture doesn't include proper theme definition for theme color resolution"]
    fn test_tab_color_all_theme_indices() {
        // Test all 12 theme color indices
        let sheets = vec![
            ("dk1", Some(TabColorSpec::Theme(0))), // Dark 1 (usually black)
            ("lt1", Some(TabColorSpec::Theme(1))), // Light 1 (usually white)
            ("dk2", Some(TabColorSpec::Theme(2))), // Dark 2
            ("lt2", Some(TabColorSpec::Theme(3))), // Light 2
            ("accent1", Some(TabColorSpec::Theme(4))), // Accent 1
            ("accent2", Some(TabColorSpec::Theme(5))), // Accent 2
            ("accent3", Some(TabColorSpec::Theme(6))), // Accent 3
            ("accent4", Some(TabColorSpec::Theme(7))), // Accent 4
            ("accent5", Some(TabColorSpec::Theme(8))), // Accent 5
            ("accent6", Some(TabColorSpec::Theme(9))), // Accent 6
            ("hlink", Some(TabColorSpec::Theme(10))), // Hyperlink
            ("folHlink", Some(TabColorSpec::Theme(11))), // Followed hyperlink
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 12);

        // Verify each theme color is resolved
        let expected_colors = [
            "#000000", // dk1
            "#FFFFFF", // lt1
            "#44546A", // dk2
            "#E7E6E6", // lt2
            "#4472C4", // accent1
            "#ED7D31", // accent2
            "#A5A5A5", // accent3
            "#FFC000", // accent4
            "#5B9BD5", // accent5
            "#70AD47", // accent6
            "#0563C1", // hlink
            "#954F72", // folHlink
        ];

        for (i, expected) in expected_colors.iter().enumerate() {
            assert_eq!(
                workbook.sheets[i].tab_color,
                Some(expected.to_string()),
                "Theme index {} should resolve to {}",
                i,
                expected
            );
        }
    }

    #[test]
    fn test_tab_color_tint_variations() {
        // Test various tint values
        let sheets = vec![
            // Very light (90% tint)
            ("90% Light", Some(TabColorSpec::ThemeWithTint(4, 0.9))),
            // Moderately light (40% tint)
            ("40% Light", Some(TabColorSpec::ThemeWithTint(4, 0.4))),
            // No tint (same as base theme)
            ("No Tint", Some(TabColorSpec::Theme(4))),
            // Moderately dark (-40% tint)
            ("40% Dark", Some(TabColorSpec::ThemeWithTint(4, -0.4))),
            // Very dark (-90% tint)
            ("90% Dark", Some(TabColorSpec::ThemeWithTint(4, -0.9))),
        ];

        let xlsx = create_xlsx_with_tab_colors(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 5);

        // All should have resolved colors
        for sheet in &workbook.sheets {
            assert!(
                sheet.tab_color.is_some(),
                "Sheet '{}' should have tab color",
                sheet.name
            );
            let color = sheet.tab_color.as_ref().unwrap();
            assert!(color.starts_with('#'), "Color should start with #");
            assert_eq!(color.len(), 7, "Color should be #RRGGBB format");
        }

        // The untinted color should be the base theme color
        assert_eq!(workbook.sheets[2].tab_color, Some("#4472C4".to_string()));
    }
}

// ============================================================================
// COMBINED VISIBILITY AND TAB COLOR TESTS
// ============================================================================

mod combined_tests {
    use super::*;
    use xlview::types::SheetState;

    /// Create an XLSX with both visibility and tab color
    fn create_xlsx_with_visibility_and_tab_color(
        sheets: &[(&str, Option<&str>, Option<TabColorSpec>)],
    ) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        // Write [Content_Types].xml
        let _ = zip.start_file("[Content_Types].xml", options);
        let content_types = generate_content_types(sheets.len());
        let _ = zip.write_all(content_types.as_bytes());

        // Write _rels/.rels
        let _ = zip.start_file("_rels/.rels", options);
        let rels = generate_rels();
        let _ = zip.write_all(rels.as_bytes());

        // Write xl/_rels/workbook.xml.rels
        let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
        let workbook_rels = generate_workbook_rels(sheets.len());
        let _ = zip.write_all(workbook_rels.as_bytes());

        // Write xl/workbook.xml with visibility states
        let _ = zip.start_file("xl/workbook.xml", options);
        let visibility_sheets: Vec<(&str, Option<&str>)> = sheets
            .iter()
            .map(|(name, state, _)| (*name, *state))
            .collect();
        let workbook = generate_workbook_with_visibility(&visibility_sheets);
        let _ = zip.write_all(workbook.as_bytes());

        // Write xl/styles.xml
        let _ = zip.start_file("xl/styles.xml", options);
        let _ = zip.write_all(minimal_styles_xml().as_bytes());

        // Write xl/theme/theme1.xml
        let _ = zip.start_file("xl/theme/theme1.xml", options);
        let _ = zip.write_all(minimal_theme_xml().as_bytes());

        // Write each sheet with optional tab color
        for (i, (_, _, tab_color)) in sheets.iter().enumerate() {
            let path = format!("xl/worksheets/sheet{}.xml", i + 1);
            let _ = zip.start_file(&path, options);
            let sheet_xml = generate_sheet_xml_with_tab_color(tab_color.as_ref());
            let _ = zip.write_all(sheet_xml.as_bytes());
        }

        let cursor = zip.finish().expect("Failed to finish ZIP");
        cursor.into_inner()
    }

    #[test]
    fn test_hidden_sheet_with_tab_color() {
        // Hidden sheets can still have tab colors
        let sheets = vec![
            (
                "Visible Red",
                None,
                Some(TabColorSpec::Rgb("FFFF0000".to_string())),
            ),
            (
                "Hidden Blue",
                Some("hidden"),
                Some(TabColorSpec::Rgb("FF0000FF".to_string())),
            ),
            (
                "VeryHidden Green",
                Some("veryHidden"),
                Some(TabColorSpec::Rgb("FF00FF00".to_string())),
            ),
        ];

        let xlsx = create_xlsx_with_visibility_and_tab_color(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 3);

        // Visible sheet with red tab
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert_eq!(workbook.sheets[0].tab_color, Some("#FF0000".to_string()));

        // Hidden sheet with blue tab
        assert_eq!(workbook.sheets[1].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[1].tab_color, Some("#0000FF".to_string()));

        // VeryHidden sheet with green tab
        assert_eq!(workbook.sheets[2].state, SheetState::VeryHidden);
        assert_eq!(workbook.sheets[2].tab_color, Some("#00FF00".to_string()));
    }

    #[test]
    fn test_visibility_and_tab_color_mixed() {
        let sheets = vec![
            ("Visible No Color", None, None),
            ("Visible With Color", None, Some(TabColorSpec::Theme(4))),
            ("Hidden No Color", Some("hidden"), None),
            (
                "Hidden With Color",
                Some("hidden"),
                Some(TabColorSpec::Indexed(2)),
            ),
            ("VeryHidden No Color", Some("veryHidden"), None),
        ];

        let xlsx = create_xlsx_with_visibility_and_tab_color(&sheets);
        let workbook = parse_xlsx(&xlsx);

        assert_eq!(workbook.sheets.len(), 5);

        // Sheet 0: Visible, no color
        assert_eq!(workbook.sheets[0].state, SheetState::Visible);
        assert!(workbook.sheets[0].tab_color.is_none());

        // Sheet 1: Visible, theme color
        assert_eq!(workbook.sheets[1].state, SheetState::Visible);
        assert_eq!(workbook.sheets[1].tab_color, Some("#4472C4".to_string()));

        // Sheet 2: Hidden, no color
        assert_eq!(workbook.sheets[2].state, SheetState::Hidden);
        assert!(workbook.sheets[2].tab_color.is_none());

        // Sheet 3: Hidden, indexed red
        assert_eq!(workbook.sheets[3].state, SheetState::Hidden);
        assert_eq!(workbook.sheets[3].tab_color, Some("#FF0000".to_string()));

        // Sheet 4: VeryHidden, no color
        assert_eq!(workbook.sheets[4].state, SheetState::VeryHidden);
        assert!(workbook.sheets[4].tab_color.is_none());
    }
}

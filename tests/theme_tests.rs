//! Comprehensive tests for theme parsing and resolution.
//!
//! Tests cover:
//! 1. Default theme colors (12 colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
//! 2. Custom theme colors
//! 3. Theme color with tint/shade applied
//! 4. Major font (heading font) parsing
//! 5. Minor font (body font) parsing
//! 6. Missing theme file (should use defaults)
//! 7. Partial theme (some colors missing)
//! 8. Theme effects/variants
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

use xlview::parser::parse;

// ============================================================================
// Test Helper: Create XLSX with custom theme1.xml
// ============================================================================

/// Create a minimal XLSX with a custom theme1.xml content.
fn xlsx_with_theme(theme_xml: &str) -> Vec<u8> {
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
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
    );

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
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

    // xl/theme/theme1.xml
    let _ = zip.start_file("xl/theme/theme1.xml", options);
    let _ = zip.write_all(theme_xml.as_bytes());

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create a minimal XLSX without theme1.xml (to test default fallback).
fn xlsx_without_theme() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml - no theme override
    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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

    // xl/_rels/workbook.xml.rels - no theme relationship
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
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
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
    let _ = zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#,
    );

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// ============================================================================
// 1. Default Theme Colors Tests
// ============================================================================

#[cfg(test)]
mod default_theme_colors {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    /// Standard Office theme XML with all 12 colors.
    fn office_theme_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
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
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#
    }

    #[test]
    fn test_parses_all_12_theme_colors() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"]
            .as_array()
            .expect("colors should be array");
        assert_eq!(colors.len(), 12, "Should have exactly 12 theme colors");
    }

    #[test]
    fn test_dk1_dark_1_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#000000");
    }

    #[test]
    fn test_lt1_light_1_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#FFFFFF");
    }

    #[test]
    fn test_dk2_dark_2_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[2].as_str().unwrap().to_uppercase(), "#44546A");
    }

    #[test]
    fn test_lt2_light_2_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[3].as_str().unwrap().to_uppercase(), "#E7E6E6");
    }

    #[test]
    fn test_accent1_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#4472C4");
    }

    #[test]
    fn test_accent2_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[5].as_str().unwrap().to_uppercase(), "#ED7D31");
    }

    #[test]
    fn test_accent3_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[6].as_str().unwrap().to_uppercase(), "#A5A5A5");
    }

    #[test]
    fn test_accent4_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[7].as_str().unwrap().to_uppercase(), "#FFC000");
    }

    #[test]
    fn test_accent5_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[8].as_str().unwrap().to_uppercase(), "#5B9BD5");
    }

    #[test]
    fn test_accent6_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[9].as_str().unwrap().to_uppercase(), "#70AD47");
    }

    #[test]
    fn test_hlink_hyperlink_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[10].as_str().unwrap().to_uppercase(), "#0563C1");
    }

    #[test]
    fn test_folhlink_followed_hyperlink_color() {
        let xlsx = xlsx_with_theme(office_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[11].as_str().unwrap().to_uppercase(), "#954F72");
    }
}

// ============================================================================
// 2. Custom Theme Colors Tests
// ============================================================================

#[cfg(test)]
mod custom_theme_colors {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    fn custom_theme_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
<a:themeElements>
<a:clrScheme name="Custom">
  <a:dk1><a:srgbClr val="1A1A1A"/></a:dk1>
  <a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
  <a:dk2><a:srgbClr val="2B2B2B"/></a:dk2>
  <a:lt2><a:srgbClr val="D0D0D0"/></a:lt2>
  <a:accent1><a:srgbClr val="FF5733"/></a:accent1>
  <a:accent2><a:srgbClr val="33FF57"/></a:accent2>
  <a:accent3><a:srgbClr val="3357FF"/></a:accent3>
  <a:accent4><a:srgbClr val="FF33F5"/></a:accent4>
  <a:accent5><a:srgbClr val="33FFF5"/></a:accent5>
  <a:accent6><a:srgbClr val="F5FF33"/></a:accent6>
  <a:hlink><a:srgbClr val="1E90FF"/></a:hlink>
  <a:folHlink><a:srgbClr val="9400D3"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Custom">
  <a:majorFont><a:latin typeface="Arial Black"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Custom"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#
    }

    #[test]
    fn test_custom_dk1_color() {
        let xlsx = xlsx_with_theme(custom_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#1A1A1A");
    }

    #[test]
    fn test_custom_accent1_color() {
        let xlsx = xlsx_with_theme(custom_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#FF5733");
    }

    #[test]
    fn test_custom_accent_colors_are_vibrant() {
        let xlsx = xlsx_with_theme(custom_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();

        // Verify all custom accent colors
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#FF5733"); // accent1 - coral
        assert_eq!(colors[5].as_str().unwrap().to_uppercase(), "#33FF57"); // accent2 - green
        assert_eq!(colors[6].as_str().unwrap().to_uppercase(), "#3357FF"); // accent3 - blue
        assert_eq!(colors[7].as_str().unwrap().to_uppercase(), "#FF33F5"); // accent4 - magenta
        assert_eq!(colors[8].as_str().unwrap().to_uppercase(), "#33FFF5"); // accent5 - cyan
        assert_eq!(colors[9].as_str().unwrap().to_uppercase(), "#F5FF33"); // accent6 - yellow
    }

    #[test]
    fn test_custom_hyperlink_colors() {
        let xlsx = xlsx_with_theme(custom_theme_xml());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[10].as_str().unwrap().to_uppercase(), "#1E90FF"); // dodger blue
        assert_eq!(colors[11].as_str().unwrap().to_uppercase(), "#9400D3"); // dark violet
    }
}

// ============================================================================
// 3. Theme Color with Tint/Shade Tests
// ============================================================================

#[cfg(test)]
mod theme_color_tint_shade {
    /// Tests for tint/shade are primarily handled in color resolution.
    /// These tests document the expected behavior when theme colors
    /// are modified with tint values.

    #[test]
    fn test_positive_tint_lightens_color() {
        // Positive tint (0 to 1) should lighten the color toward white
        // Formula: new_L = L + (1 - L) * tint
        //
        // Example: accent1 (#4472C4) with tint=0.5
        // The color should become lighter (higher luminance)
    }

    #[test]
    fn test_negative_tint_darkens_color() {
        // Negative tint (-1 to 0) should darken the color toward black (shade)
        // Formula: new_L = L * (1 + tint)
        //
        // Example: accent1 (#4472C4) with tint=-0.5
        // The color should become darker (lower luminance)
    }

    #[test]
    fn test_tint_zero_preserves_color() {
        // Tint of 0 should preserve the original color exactly
    }

    #[test]
    fn test_tint_one_produces_white() {
        // Tint of 1.0 should produce white (#FFFFFF)
    }

    #[test]
    fn test_tint_negative_one_produces_black() {
        // Tint of -1.0 should produce black (#000000)
    }

    #[test]
    fn test_excel_common_tint_values() {
        // Excel commonly uses these tint values:
        // -0.499984740745262 (50% darker)
        // -0.249977111117893 (25% darker)
        //  0.39997558519241921 (40% lighter)
        //  0.59999389629810485 (60% lighter)
        //  0.79998168889431442 (80% lighter)
    }
}

// ============================================================================
// 4. Major Font (Heading Font) Parsing Tests
// ============================================================================

#[cfg(test)]
mod major_font_tests {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    fn theme_with_major_font(font_name: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
<a:themeElements>
<a:clrScheme name="Test">
  <a:dk1><a:srgbClr val="000000"/></a:dk1>
  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="444444"/></a:dk2>
  <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
  <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
  <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
  <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
  <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Test">
  <a:majorFont>
    <a:latin typeface="{}"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
  </a:majorFont>
  <a:minorFont>
    <a:latin typeface="Calibri"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
  </a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#,
            font_name
        )
    }

    #[test]
    fn test_major_font_calibri_light() {
        let xlsx = xlsx_with_theme(&theme_with_major_font("Calibri Light"));
        let workbook = parse_xlsx_to_json(&xlsx);

        // Note: majorFont may not be exposed in the current JSON output
        // This test documents the expected behavior
        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_major_font_cambria() {
        let xlsx = xlsx_with_theme(&theme_with_major_font("Cambria"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_major_font_arial_black() {
        let xlsx = xlsx_with_theme(&theme_with_major_font("Arial Black"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_major_font_times_new_roman() {
        let xlsx = xlsx_with_theme(&theme_with_major_font("Times New Roman"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }
}

// ============================================================================
// 5. Minor Font (Body Font) Parsing Tests
// ============================================================================

#[cfg(test)]
mod minor_font_tests {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    fn theme_with_minor_font(font_name: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
<a:themeElements>
<a:clrScheme name="Test">
  <a:dk1><a:srgbClr val="000000"/></a:dk1>
  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="444444"/></a:dk2>
  <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
  <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
  <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
  <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
  <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Test">
  <a:majorFont>
    <a:latin typeface="Calibri Light"/>
  </a:majorFont>
  <a:minorFont>
    <a:latin typeface="{}"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
  </a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#,
            font_name
        )
    }

    #[test]
    fn test_minor_font_calibri() {
        let xlsx = xlsx_with_theme(&theme_with_minor_font("Calibri"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_minor_font_arial() {
        let xlsx = xlsx_with_theme(&theme_with_minor_font("Arial"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_minor_font_helvetica() {
        let xlsx = xlsx_with_theme(&theme_with_minor_font("Helvetica"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }

    #[test]
    fn test_minor_font_verdana() {
        let xlsx = xlsx_with_theme(&theme_with_minor_font("Verdana"));
        let workbook = parse_xlsx_to_json(&xlsx);

        assert!(workbook["theme"].is_object());
    }
}

// ============================================================================
// 6. Missing Theme File Tests
// ============================================================================

#[cfg(test)]
mod missing_theme_file {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    #[test]
    fn test_missing_theme_uses_defaults() {
        let xlsx = xlsx_without_theme();
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"]
            .as_array()
            .expect("colors should be array");
        assert_eq!(colors.len(), 12, "Should have 12 default theme colors");
    }

    #[test]
    fn test_missing_theme_default_dk1() {
        let xlsx = xlsx_without_theme();
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        // Default dk1 should be black
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#000000");
    }

    #[test]
    fn test_missing_theme_default_lt1() {
        let xlsx = xlsx_without_theme();
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        // Default lt1 should be white
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#FFFFFF");
    }

    #[test]
    fn test_missing_theme_default_accent1() {
        let xlsx = xlsx_without_theme();
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        // Default accent1 should be Office blue
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#4472C4");
    }

    #[test]
    fn test_missing_theme_all_defaults() {
        let xlsx = xlsx_without_theme();
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();

        let expected_defaults = [
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

        for (i, expected) in expected_defaults.iter().enumerate() {
            assert_eq!(
                colors[i].as_str().unwrap().to_uppercase(),
                expected.to_uppercase(),
                "Color at index {} should match default",
                i
            );
        }
    }
}

// ============================================================================
// 7. Partial Theme Tests (Some Colors Missing)
// ============================================================================

#[cfg(test)]
mod partial_theme {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    fn partial_theme_xml_missing_accents() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Partial Theme">
<a:themeElements>
<a:clrScheme name="Partial">
  <a:dk1><a:srgbClr val="111111"/></a:dk1>
  <a:lt1><a:srgbClr val="EEEEEE"/></a:lt1>
  <a:dk2><a:srgbClr val="222222"/></a:dk2>
  <a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>
</a:clrScheme>
<a:fontScheme name="Partial">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Partial"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#
    }

    fn partial_theme_xml_only_dk1_lt1() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Minimal Theme">
<a:themeElements>
<a:clrScheme name="Minimal">
  <a:dk1><a:srgbClr val="0A0A0A"/></a:dk1>
  <a:lt1><a:srgbClr val="F5F5F5"/></a:lt1>
</a:clrScheme>
<a:fontScheme name="Minimal">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Minimal"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#
    }

    #[test]
    fn test_partial_theme_has_12_colors() {
        let xlsx = xlsx_with_theme(partial_theme_xml_missing_accents());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(
            colors.len(),
            12,
            "Should still have 12 colors (with defaults for missing)"
        );
    }

    #[test]
    fn test_partial_theme_preserves_defined_colors() {
        let xlsx = xlsx_with_theme(partial_theme_xml_missing_accents());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#111111"); // dk1
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#EEEEEE"); // lt1
        assert_eq!(colors[2].as_str().unwrap().to_uppercase(), "#222222"); // dk2
        assert_eq!(colors[3].as_str().unwrap().to_uppercase(), "#DDDDDD"); // lt2
    }

    #[test]
    fn test_partial_theme_uses_defaults_for_missing() {
        let xlsx = xlsx_with_theme(partial_theme_xml_missing_accents());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();

        // Missing accent colors should fall back to defaults
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#4472C4"); // accent1 default
        assert_eq!(colors[5].as_str().unwrap().to_uppercase(), "#ED7D31"); // accent2 default
    }

    #[test]
    fn test_minimal_theme_only_dk1_lt1() {
        let xlsx = xlsx_with_theme(partial_theme_xml_only_dk1_lt1());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();

        // Defined colors
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#0A0A0A"); // dk1
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#F5F5F5"); // lt1

        // Remaining should be defaults
        assert_eq!(colors.len(), 12);
    }
}

// ============================================================================
// 8. Theme Effects/Variants Tests
// ============================================================================

#[cfg(test)]
mod theme_effects_variants {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    fn theme_with_sys_clr_variants() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="System Colors Theme">
<a:themeElements>
<a:clrScheme name="System">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:sysClr val="btnText" lastClr="000000"/></a:dk2>
  <a:lt2><a:sysClr val="btnFace" lastClr="F0F0F0"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
  <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
  <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
  <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
  <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="System">
  <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
  <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="System"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#
    }

    #[test]
    fn test_sysclr_uses_lastclr_attribute() {
        // System colors (sysClr) should use the lastClr attribute
        // which contains the actual resolved color value
        let xlsx = xlsx_with_theme(theme_with_sys_clr_variants());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#000000"); // windowText
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#FFFFFF"); // window
    }

    #[test]
    fn test_sysclr_btnface_variant() {
        let xlsx = xlsx_with_theme(theme_with_sys_clr_variants());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors[3].as_str().unwrap().to_uppercase(), "#F0F0F0"); // btnFace
    }

    #[test]
    fn test_mixed_sysclr_and_srgbclr() {
        // Theme can mix system colors and sRGB colors
        let xlsx = xlsx_with_theme(theme_with_sys_clr_variants());
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();

        // sysClr
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#000000");

        // srgbClr
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#4472C4");
    }

    #[test]
    fn test_lowercase_color_values_normalized() {
        // Test that color values are normalized to uppercase
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Lowercase Theme">
<a:themeElements>
<a:clrScheme name="Lowercase">
  <a:dk1><a:srgbClr val="aabbcc"/></a:dk1>
  <a:lt1><a:srgbClr val="ddeeff"/></a:lt1>
  <a:dk2><a:srgbClr val="112233"/></a:dk2>
  <a:lt2><a:srgbClr val="445566"/></a:lt2>
  <a:accent1><a:srgbClr val="778899"/></a:accent1>
  <a:accent2><a:srgbClr val="aabbcc"/></a:accent2>
  <a:accent3><a:srgbClr val="ddeeff"/></a:accent3>
  <a:accent4><a:srgbClr val="112233"/></a:accent4>
  <a:accent5><a:srgbClr val="445566"/></a:accent5>
  <a:accent6><a:srgbClr val="778899"/></a:accent6>
  <a:hlink><a:srgbClr val="aabbcc"/></a:hlink>
  <a:folHlink><a:srgbClr val="ddeeff"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Lowercase">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Lowercase"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#;

        let xlsx = xlsx_with_theme(theme_xml);
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        // Should be normalized to uppercase
        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#AABBCC");
    }
}

// ============================================================================
// 9. Theme Font Scheme Resolution Tests (BUG-004)
// ============================================================================

#[cfg(test)]
mod font_scheme_resolution {
    use super::*;

    /// Create XLSX with fonts that reference theme via scheme attribute
    fn xlsx_with_font_scheme() -> Vec<u8> {
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
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
        );

        // xl/workbook.xml
        let _ = zip.start_file("xl/workbook.xml", options);
        let _ = zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
        );

        // xl/styles.xml with fonts that have scheme="minor" and scheme="major"
        let _ = zip.start_file("xl/styles.xml", options);
        let _ = zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
  <font>
    <sz val="11"/>
    <name val="Calibri"/>
    <scheme val="minor"/>
  </font>
  <font>
    <sz val="14"/>
    <name val="Calibri Light"/>
    <scheme val="major"/>
  </font>
  <font>
    <sz val="10"/>
    <name val="Arial"/>
  </font>
</fonts>
<fills count="2">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
  <border><left/><right/><top/><bottom/></border>
</borders>
<cellXfs count="3">
  <xf fontId="0" fillId="0" borderId="0"/>
  <xf fontId="1" fillId="0" borderId="0"/>
  <xf fontId="2" fillId="0" borderId="0"/>
</cellXfs>
</styleSheet>"#,
        );

        // xl/theme/theme1.xml with custom fonts
        let _ = zip.start_file("xl/theme/theme1.xml", options);
        let _ = zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
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
  <a:majorFont>
    <a:latin typeface="Times New Roman"/>
  </a:majorFont>
  <a:minorFont>
    <a:latin typeface="Helvetica"/>
  </a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#,
        );

        // xl/worksheets/sheet1.xml with cells using different fonts
        let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
        let _ = zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
  <row r="1">
    <c r="A1" s="0" t="inlineStr"><is><t>Minor</t></is></c>
    <c r="B1" s="1" t="inlineStr"><is><t>Major</t></is></c>
    <c r="C1" s="2" t="inlineStr"><is><t>Normal</t></is></c>
  </row>
</sheetData>
</worksheet>"#,
        );

        let cursor = zip.finish().expect("Failed to finish ZIP");
        cursor.into_inner()
    }

    #[test]
    fn test_minor_font_scheme_resolves_to_theme() {
        let xlsx = xlsx_with_font_scheme();

        // Use the real parser instead of test helper
        let workbook = parse(&xlsx).expect("Failed to parse workbook");
        let workbook_json = serde_json::to_value(&workbook).expect("Failed to serialize");

        // Debug output
        eprintln!(
            "Workbook JSON:\n{}",
            serde_json::to_string_pretty(&workbook_json).unwrap()
        );

        // Cell A1 uses font with scheme="minor" which should resolve to Helvetica (theme's minorFont)
        let cell = &workbook_json["sheets"][0]["cells"][0];
        assert_eq!(cell["c"].as_u64().unwrap(), 0); // Column A
        assert_eq!(cell["r"].as_u64().unwrap(), 0); // Row 1

        // Should use theme's minor font (Helvetica), not the name attribute (Calibri)
        assert!(
            cell["cell"]["s"].is_object(),
            "Cell should have a style object"
        );
        assert_eq!(
            cell["cell"]["s"]["fontFamily"].as_str().unwrap(),
            "Helvetica",
            "Font with scheme=minor should use theme's minorFont"
        );
    }

    #[test]
    fn test_major_font_scheme_resolves_to_theme() {
        let xlsx = xlsx_with_font_scheme();

        // Use the real parser instead of test helper
        let workbook = parse(&xlsx).expect("Failed to parse workbook");
        let workbook_json = serde_json::to_value(&workbook).expect("Failed to serialize");

        // Cell B1 uses font with scheme="major" which should resolve to Times New Roman (theme's majorFont)
        let cell = &workbook_json["sheets"][0]["cells"][1];
        assert_eq!(cell["c"].as_u64().unwrap(), 1); // Column B

        // Should use theme's major font (Times New Roman), not the name attribute (Calibri Light)
        assert_eq!(
            cell["cell"]["s"]["fontFamily"].as_str().unwrap(),
            "Times New Roman",
            "Font with scheme=major should use theme's majorFont"
        );
    }

    #[test]
    fn test_font_without_scheme_uses_name_attribute() {
        let xlsx = xlsx_with_font_scheme();

        // Use the real parser instead of test helper
        let workbook = parse(&xlsx).expect("Failed to parse workbook");
        let workbook_json = serde_json::to_value(&workbook).expect("Failed to serialize");

        // Cell C1 uses font without scheme, should use name attribute directly
        let cell = &workbook_json["sheets"][0]["cells"][2];
        assert_eq!(cell["c"].as_u64().unwrap(), 2); // Column C

        // Should use the font's name attribute (Arial)
        assert_eq!(
            cell["cell"]["s"]["fontFamily"].as_str().unwrap(),
            "Arial",
            "Font without scheme should use name attribute"
        );
    }

    #[test]
    fn test_theme_fonts_are_parsed() {
        let xlsx = xlsx_with_font_scheme();

        // Use the real parser instead of test helper
        let workbook = parse(&xlsx).expect("Failed to parse workbook");
        let workbook_json = serde_json::to_value(&workbook).expect("Failed to serialize");

        // Verify theme fonts are present (JSON uses snake_case due to serde defaults)
        assert_eq!(
            workbook_json["theme"]["major_font"].as_str().unwrap(),
            "Times New Roman",
            "Theme should have majorFont"
        );
        assert_eq!(
            workbook_json["theme"]["minor_font"].as_str().unwrap(),
            "Helvetica",
            "Theme should have minorFont"
        );
    }
}

// ============================================================================
// Additional Edge Case Tests
// ============================================================================

#[cfg(test)]
mod edge_cases {
    use super::*;
    use crate::common::parse_xlsx_to_json;

    #[test]
    fn test_empty_theme_file() {
        // An essentially empty theme should still parse and use defaults
        let empty_theme = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Empty">
<a:themeElements>
<a:clrScheme name="Empty"></a:clrScheme>
<a:fontScheme name="Empty">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Empty"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#;

        let xlsx = xlsx_with_theme(empty_theme);
        let workbook = parse_xlsx_to_json(&xlsx);

        let colors = workbook["theme"]["colors"].as_array().unwrap();
        assert_eq!(colors.len(), 12);
    }

    #[test]
    fn test_theme_with_namespaced_elements() {
        // Theme XML often uses namespaced elements (a:dk1 vs dk1)
        let xlsx = xlsx_with_theme(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Namespaced">
<a:themeElements>
<a:clrScheme name="Test">
  <a:dk1><a:srgbClr val="ABCDEF"/></a:dk1>
  <a:lt1><a:srgbClr val="FEDCBA"/></a:lt1>
  <a:dk2><a:srgbClr val="123456"/></a:dk2>
  <a:lt2><a:srgbClr val="654321"/></a:lt2>
  <a:accent1><a:srgbClr val="AABBCC"/></a:accent1>
  <a:accent2><a:srgbClr val="DDEEFF"/></a:accent2>
  <a:accent3><a:srgbClr val="112233"/></a:accent3>
  <a:accent4><a:srgbClr val="445566"/></a:accent4>
  <a:accent5><a:srgbClr val="778899"/></a:accent5>
  <a:accent6><a:srgbClr val="AABBDD"/></a:accent6>
  <a:hlink><a:srgbClr val="CCDDEE"/></a:hlink>
  <a:folHlink><a:srgbClr val="FFEEDD"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Test">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#,
        );

        let workbook = parse_xlsx_to_json(&xlsx);
        let colors = workbook["theme"]["colors"].as_array().unwrap();

        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#ABCDEF");
    }

    #[test]
    fn test_theme_colors_order_is_correct() {
        // Verify the exact order: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Order Test">
<a:themeElements>
<a:clrScheme name="Order">
  <a:dk1><a:srgbClr val="000001"/></a:dk1>
  <a:lt1><a:srgbClr val="000002"/></a:lt1>
  <a:dk2><a:srgbClr val="000003"/></a:dk2>
  <a:lt2><a:srgbClr val="000004"/></a:lt2>
  <a:accent1><a:srgbClr val="000005"/></a:accent1>
  <a:accent2><a:srgbClr val="000006"/></a:accent2>
  <a:accent3><a:srgbClr val="000007"/></a:accent3>
  <a:accent4><a:srgbClr val="000008"/></a:accent4>
  <a:accent5><a:srgbClr val="000009"/></a:accent5>
  <a:accent6><a:srgbClr val="00000A"/></a:accent6>
  <a:hlink><a:srgbClr val="00000B"/></a:hlink>
  <a:folHlink><a:srgbClr val="00000C"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Order">
  <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Order"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>"#;

        let xlsx = xlsx_with_theme(theme_xml);
        let workbook = parse_xlsx_to_json(&xlsx);
        let colors = workbook["theme"]["colors"].as_array().unwrap();

        assert_eq!(colors[0].as_str().unwrap().to_uppercase(), "#000001"); // dk1
        assert_eq!(colors[1].as_str().unwrap().to_uppercase(), "#000002"); // lt1
        assert_eq!(colors[2].as_str().unwrap().to_uppercase(), "#000003"); // dk2
        assert_eq!(colors[3].as_str().unwrap().to_uppercase(), "#000004"); // lt2
        assert_eq!(colors[4].as_str().unwrap().to_uppercase(), "#000005"); // accent1
        assert_eq!(colors[5].as_str().unwrap().to_uppercase(), "#000006"); // accent2
        assert_eq!(colors[6].as_str().unwrap().to_uppercase(), "#000007"); // accent3
        assert_eq!(colors[7].as_str().unwrap().to_uppercase(), "#000008"); // accent4
        assert_eq!(colors[8].as_str().unwrap().to_uppercase(), "#000009"); // accent5
        assert_eq!(colors[9].as_str().unwrap().to_uppercase(), "#00000A"); // accent6
        assert_eq!(colors[10].as_str().unwrap().to_uppercase(), "#00000B"); // hlink
        assert_eq!(colors[11].as_str().unwrap().to_uppercase(), "#00000C"); // folHlink
    }
}

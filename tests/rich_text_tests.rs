//! Rich Text Tests for xlview
//!
//! Tests for parsing rich text content in XLSX files.
//! Rich text allows multiple formats within a single cell, using <r> (run) elements
//! with <rPr> (run properties) to define formatting for each segment.
//!
//! IMPORTANT DATA MODEL NOTES:
//! The current Cell.v is Option<String>, so rich text is flattened to plain text.
//! Future enhancements could add:
//! - A `rich_text: Option<Vec<RichTextRun>>` field
//! - Convert to HTML/annotated string
//! - Or keep flattened for v1 and add rich text support later
//!
//! These tests document expected behavior and serve as a specification.
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

/// Helper to create a minimal XLSX file for testing
fn create_test_xlsx(shared_strings_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buffer);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#).unwrap();

        // xl/workbook.xml
        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </sheets>
</workbook>"#).unwrap();

        // xl/sharedStrings.xml
        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        // xl/worksheets/sheet1.xml
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    buffer.into_inner()
}

/// Helper to create XLSX with inline strings in the sheet
fn create_test_xlsx_inline(sheet_xml: &str) -> Vec<u8> {
    let empty_shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>"#;
    create_test_xlsx(empty_shared_strings, sheet_xml)
}

// ============================================================================
// SHARED STRINGS RICH TEXT TESTS
// ============================================================================

mod shared_strings_rich_text {
    use super::*;

    /// Test 1: Simple rich text - Part bold, part normal
    /// Expected: Text content "Bold Normal" should be extracted (concatenated)
    #[test]
    fn test_simple_rich_text_bold_and_normal() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><t> Normal</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].cells.len(), 1);

        let cell = &workbook.sheets[0].cells[0].cell;
        // Rich text should be concatenated to plain text
        assert_eq!(cell.v.as_deref(), Some("Bold Normal"));
    }

    /// Test 2: Multiple formats - Bold, italic, and colored runs
    /// Expected: Text "Bold Italic Red" concatenated
    #[test]
    fn test_multiple_format_runs() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold </t></r>
    <r><rPr><i/></rPr><t>Italic </t></r>
    <r><rPr><color rgb="FFFF0000"/></rPr><t>Red</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Bold Italic Red"));
    }

    /// Test 3: Font size changes within cell
    /// Expected: Text "Small Large" concatenated
    #[test]
    fn test_font_size_changes() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><sz val="8"/></rPr><t>Small </t></r>
    <r><rPr><sz val="14"/></rPr><t>Large</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Small Large"));
    }

    /// Test 4: Different font families
    /// Expected: Text "Arial Courier" concatenated
    #[test]
    fn test_font_family_changes() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><rFont val="Arial"/></rPr><t>Arial </t></r>
    <r><rPr><rFont val="Courier New"/></rPr><t>Courier</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Arial Courier"));
    }

    /// Test 5: Partial underline in rich text
    /// Expected: Text "Normal Underlined" concatenated
    #[test]
    fn test_partial_underline() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Normal </t></r>
    <r><rPr><u/></rPr><t>Underlined</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Normal Underlined"));
    }

    /// Test 6: Partial strikethrough in rich text
    /// Expected: Text "Normal Strikethrough" concatenated
    #[test]
    fn test_partial_strikethrough() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>Normal </t></r>
    <r><rPr><strike/></rPr><t>Strikethrough</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Normal Strikethrough"));
    }

    /// Test 7: Subscript text using vertAlign
    /// Expected: Text "H2O" concatenated (subscript 2)
    #[test]
    fn test_subscript() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>H</t></r>
    <r><rPr><vertAlign val="subscript"/></rPr><t>2</t></r>
    <r><t>O</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("H2O"));
    }

    /// Test 8: Superscript text using vertAlign
    /// Expected: Text "x2" concatenated (superscript 2)
    #[test]
    fn test_superscript() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><t>x</t></r>
    <r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("x2"));
    }

    /// Test 9: Complex combination - Multiple properties in one run
    /// Expected: Text "Bold+Italic+Red+Size" concatenated
    #[test]
    fn test_complex_combination() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr>
        <b/>
        <i/>
        <color rgb="FFFF0000"/>
        <sz val="14"/>
        <rFont val="Arial"/>
        <u/>
      </rPr>
      <t>Bold+Italic+Red+Size</t>
    </r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Bold+Italic+Red+Size"));
    }

    /// Test: Mixed plain and rich text in shared strings table
    /// The shared strings table can contain both <si><t>plain</t></si> and <si><r>...</r></si>
    #[test]
    fn test_mixed_plain_and_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Plain text</t></si>
  <si>
    <r><rPr><b/></rPr><t>Rich</t></r>
    <r><t> text</t></r>
  </si>
  <si><t>Another plain</t></si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        assert_eq!(workbook.sheets[0].cells.len(), 3);

        // Find cells by column
        let cell_a = workbook.sheets[0].cells.iter().find(|c| c.c == 0).unwrap();
        let cell_b = workbook.sheets[0].cells.iter().find(|c| c.c == 1).unwrap();
        let cell_c = workbook.sheets[0].cells.iter().find(|c| c.c == 2).unwrap();

        assert_eq!(cell_a.cell.v.as_deref(), Some("Plain text"));
        assert_eq!(cell_b.cell.v.as_deref(), Some("Rich text"));
        assert_eq!(cell_c.cell.v.as_deref(), Some("Another plain"));
    }

    /// Test: Empty runs should be handled gracefully
    #[test]
    fn test_empty_rich_text_runs() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t></t></r>
    <r><t>Content</t></r>
    <r><rPr><i/></rPr><t></t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Content"));
    }

    /// Test: Whitespace preservation in rich text runs
    /// XML space handling: <t xml:space="preserve"> to preserve leading/trailing spaces
    #[test]
    fn test_whitespace_preservation() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t xml:space="preserve">Bold </t></r>
    <r><t xml:space="preserve"> and </t></r>
    <r><rPr><i/></rPr><t xml:space="preserve"> Italic</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        // Depending on implementation, spaces may or may not be preserved
        // This documents expected behavior
        let value = cell.v.as_deref().unwrap_or("");
        assert!(
            value.contains("Bold") && value.contains("and") && value.contains("Italic"),
            "Expected all text parts, got: {}",
            value
        );
    }

    /// Test: Theme colors in rich text run properties
    #[test]
    fn test_theme_color_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><color theme="4"/></rPr><t>Accent1 Color</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Accent1 Color"));
    }

    /// Test: Double underline in rich text
    #[test]
    fn test_double_underline_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><u val="double"/></rPr><t>Double Underline</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Double Underline"));
    }

    /// Test: Rich text with character set (charset) and font family type
    #[test]
    fn test_rich_text_with_charset() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r>
      <rPr>
        <rFont val="Calibri"/>
        <charset val="1"/>
        <family val="2"/>
        <scheme val="minor"/>
      </rPr>
      <t>With Charset</t>
    </r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("With Charset"));
    }
}

// ============================================================================
// INLINE STRING RICH TEXT TESTS
// ============================================================================

mod inline_string_rich_text {
    use super::*;

    /// Test 10: Inline rich text in cell (t="inlineStr")
    /// Cells can have rich text directly embedded, not referencing shared strings
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_rich_text() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><rPr><b/></rPr><t>Bold</t></r>
          <r><t> Normal</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].cells.len(), 1);

        let cell = &workbook.sheets[0].cells[0].cell;
        // Inline rich text should be concatenated to plain text
        assert_eq!(cell.v.as_deref(), Some("Bold Normal"));
    }

    /// Test: Inline string with plain text (no rich text runs)
    #[test]
    fn test_inline_plain_text() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Plain inline string</t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Plain inline string"));
    }

    /// Test: Inline rich text with multiple formatting runs
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_multiple_runs() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><rPr><b/></rPr><t>Bold </t></r>
          <r><rPr><i/></rPr><t>Italic </t></r>
          <r><rPr><u/></rPr><t>Underline</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Bold Italic Underline"));
    }

    /// Test: Inline rich text with subscript/superscript
    #[test]
    #[ignore = "TODO: Inline rich text run concatenation not yet implemented"]
    fn test_inline_sub_superscript() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><t>E=mc</t></r>
          <r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("E=mc2"));
    }

    /// Test: Empty inline string
    #[test]
    fn test_empty_inline_string() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t></t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx_inline(sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        // Empty cell might not be included, or might have empty value
        if !workbook.sheets[0].cells.is_empty() {
            let cell = &workbook.sheets[0].cells[0].cell;
            assert!(
                cell.v.is_none() || cell.v.as_deref() == Some(""),
                "Expected empty value, got: {:?}",
                cell.v
            );
        }
    }
}

// ============================================================================
// EDGE CASES AND ERROR HANDLING
// ============================================================================

mod edge_cases {
    use super::*;

    /// Test: Malformed rich text (missing <t> element)
    #[test]
    fn test_rich_text_missing_text_element() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr></r>
    <r><t>Only this</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        assert_eq!(cell.v.as_deref(), Some("Only this"));
    }

    /// Test: Rich text with nested elements (like phonetic reading <rPh>)
    /// Japanese/Chinese Excel files may have phonetic guides
    #[test]
    fn test_rich_text_with_phonetic_reading() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Main Text</t>
    <rPh sb="0" eb="4">
      <t>phonetic</t>
    </rPh>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        // Main text should be extracted, phonetic reading may or may not be included
        let value = cell.v.as_deref().unwrap_or("");
        assert!(
            value.contains("Main Text"),
            "Expected main text, got: {}",
            value
        );
    }

    /// Test: Very long rich text string (many runs)
    #[test]
    fn test_many_rich_text_runs() {
        // Build a shared string with 10 alternating bold/italic runs
        let runs: String = (0..10)
            .map(|i| {
                if i % 2 == 0 {
                    format!(r#"<r><rPr><b/></rPr><t>Run{} </t></r>"#, i)
                } else {
                    format!(r#"<r><rPr><i/></rPr><t>Run{} </t></r>"#, i)
                }
            })
            .collect();

        let shared_strings = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    {}
  </si>
</sst>"#,
            runs
        );

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(&shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        let value = cell.v.as_deref().unwrap_or("");

        // Should contain all run texts
        for i in 0..10 {
            assert!(
                value.contains(&format!("Run{}", i)),
                "Missing Run{} in: {}",
                i,
                value
            );
        }
    }

    /// Test: Unicode text in rich text runs
    #[test]
    fn test_unicode_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Hello </t></r>
    <r><t>世界 </t></r>
    <r><rPr><i/></rPr><t>Привет</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        let value = cell.v.as_deref().unwrap_or("");

        assert!(value.contains("Hello"), "Missing English text");
        assert!(value.contains("世界"), "Missing Chinese text");
        assert!(value.contains("Привет"), "Missing Russian text");
    }

    /// Test: Special XML characters in rich text
    #[test]
    fn test_xml_entities_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>&lt;bold&gt;</t></r>
    <r><t> &amp; </t></r>
    <r><rPr><i/></rPr><t>&quot;quoted&quot;</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        let value = cell.v.as_deref().unwrap_or("");

        // Entities should be unescaped
        assert!(value.contains("<bold>"), "Expected unescaped <bold>");
        assert!(value.contains("&"), "Expected unescaped &");
        assert!(value.contains("\"quoted\""), "Expected unescaped quotes");
    }

    /// Test: Newlines within rich text runs
    #[test]
    fn test_newlines_in_rich_text() {
        let shared_strings = r#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Line1
Line2</t></r>
    <r><t>
Line3</t></r>
  </si>
</sst>"#;

        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let xlsx_data = create_test_xlsx(shared_strings, sheet);
        let workbook = xlview::parser::parse(&xlsx_data).expect("Failed to parse XLSX");

        let cell = &workbook.sheets[0].cells[0].cell;
        let value = cell.v.as_deref().unwrap_or("");

        assert!(value.contains("Line1"), "Missing Line1");
        assert!(value.contains("Line2"), "Missing Line2");
        assert!(value.contains("Line3"), "Missing Line3");
    }
}

// ============================================================================
// FUTURE RICH TEXT DATA MODEL TESTS
// ============================================================================
// These tests document what a future rich text implementation should support.
// Currently marked as ignore since the data model doesn't support rich text runs.

mod future_rich_text_model {
    #[allow(dead_code)]
    /// Future: A rich text run with formatting properties
    #[derive(Debug, Clone, PartialEq)]
    struct RichTextRun {
        text: String,
        bold: bool,
        italic: bool,
        underline: bool,
        strikethrough: bool,
        font_size: Option<f64>,
        font_family: Option<String>,
        color: Option<String>,
        vert_align: Option<VerticalAlign>,
    }

    #[allow(dead_code)]
    #[derive(Debug, Clone, PartialEq)]
    enum VerticalAlign {
        Baseline,
        Subscript,
        Superscript,
    }

    /// Document what the future data model should look like
    #[test]
    #[ignore = "Future enhancement: rich text run data model"]
    fn test_future_rich_text_runs_model() {
        // When rich text is fully supported, a cell should provide:
        // 1. Plain text (concatenated) - for backward compatibility
        // 2. Rich text runs with individual formatting

        // Example of what the future API might look like:
        // let cell = ...;
        // assert_eq!(cell.plain_text(), "Bold Normal");
        // let runs = cell.rich_text_runs().unwrap();
        // assert_eq!(runs.len(), 2);
        // assert!(runs[0].bold);
        // assert_eq!(runs[0].text, "Bold");
        // assert!(!runs[1].bold);
        // assert_eq!(runs[1].text, " Normal");
    }

    /// Document expected HTML conversion for rich text
    #[test]
    #[ignore = "Future enhancement: rich text to HTML conversion"]
    fn test_future_rich_text_to_html() {
        // Future: Rich text should be convertible to HTML for rendering
        // Example output:
        // "<span style=\"font-weight:bold\">Bold</span><span> Normal</span>"

        // let cell = ...;
        // let html = cell.to_html();
        // assert!(html.contains("<span"));
        // assert!(html.contains("font-weight:bold"));
    }

    /// Document expected behavior for nested formatting
    #[test]
    #[ignore = "Future enhancement: rich text run property resolution"]
    fn test_future_nested_run_properties() {
        // Rich text runs can have multiple properties
        // The parser should correctly capture all of them

        // <r>
        //   <rPr>
        //     <b/>
        //     <i/>
        //     <color rgb="FFFF0000"/>
        //     <sz val="14"/>
        //   </rPr>
        //   <t>Formatted</t>
        // </r>

        // let run = ...;
        // assert!(run.bold);
        // assert!(run.italic);
        // assert_eq!(run.color, Some("#FF0000"));
        // assert_eq!(run.font_size, Some(14.0));
    }
}

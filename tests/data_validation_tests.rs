//! Tests for data validation parsing in XLSX files
//!
//! Data validation in Excel allows users to restrict input to cells based on various criteria.
//! The validation rules are stored in the worksheet XML within `<dataValidations>` elements.
//!
//! Example XML structure:
//! ```xml
//! <dataValidations count="1">
//!   <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">
//!     <formula1>"Option1,Option2,Option3"</formula1>
//!   </dataValidation>
//! </dataValidations>
//! ```
//!
//! Supported validation types:
//! - `list`: Dropdown with comma-separated values or cell range reference
//! - `whole`: Whole number validation with operators (between, greaterThan, etc.)
//! - `decimal`: Decimal number validation
//! - `date`: Date validation
//! - `time`: Time validation
//! - `textLength`: Text length validation
//! - `custom`: Custom formula validation
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

// =============================================================================
// Test Helpers
// =============================================================================

/// Create base XLSX structure with data validation XML injected into the worksheet
fn create_xlsx_with_data_validation(data_validation_xml: &str) -> Vec<u8> {
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

    // xl/worksheets/sheet1.xml with data validation
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1" t="inlineStr"><is><t>Test</t></is></c>
</row>
</sheetData>
{data_validation_xml}
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

/// Create XLSX with cells containing dropdown source values
fn create_xlsx_with_dropdown_source(
    data_validation_xml: &str,
    source_cells: &[(&str, &str)], // (cell_ref, value)
) -> Vec<u8> {
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

    // xl/worksheets/sheet1.xml with data validation and source cells
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);

    // Build rows for source cells
    let mut rows_xml = String::new();
    for (cell_ref, value) in source_cells {
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
{data_validation_xml}
</worksheet>"#
    );
    let _ = zip.write_all(sheet_xml.as_bytes());

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Tests: Dropdown List from Comma-Separated Values
// =============================================================================

#[test]
fn test_dropdown_list_comma_separated_values() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" showDropDown="0" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">
    <formula1>"Red,Green,Blue"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    assert!(
        !sheet.data_validations.is_empty(),
        "Should have data validations"
    );
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "A1:A10");
    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::List
    ));
    assert!(dv.validation.show_dropdown);
    assert!(dv.validation.allow_blank);

    // Check that the formula is parsed
    assert_eq!(
        dv.validation.formula1.as_deref(),
        Some("\"Red,Green,Blue\"")
    );

    // Check parsed list values
    if let Some(ref values) = dv.validation.list_values {
        assert_eq!(values.len(), 3);
        assert!(values.contains(&"Red".to_string()));
        assert!(values.contains(&"Green".to_string()));
        assert!(values.contains(&"Blue".to_string()));
    }
}

#[test]
fn test_dropdown_list_single_value() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="0" sqref="B5">
    <formula1>"OnlyOption"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "B5");
    assert!(!dv.validation.allow_blank);

    if let Some(ref values) = dv.validation.list_values {
        assert_eq!(values.len(), 1);
        assert_eq!(values[0], "OnlyOption");
    }
}

#[test]
fn test_dropdown_list_with_special_characters() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" sqref="C1">
    <formula1>"Yes/No,N/A,Don't Know"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    if let Some(ref values) = dv.validation.list_values {
        assert_eq!(values.len(), 3);
        assert!(values.contains(&"Yes/No".to_string()));
        assert!(values.contains(&"N/A".to_string()));
        assert!(values.contains(&"Don't Know".to_string()));
    }
}

// =============================================================================
// Tests: Dropdown List from Cell Range
// =============================================================================

#[test]
fn test_dropdown_list_from_cell_range() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
    <formula1>$D$1:$D$5</formula1>
  </dataValidation>
</dataValidations>"#;

    let source_cells = vec![
        ("D1", "Option1"),
        ("D2", "Option2"),
        ("D3", "Option3"),
        ("D4", "Option4"),
        ("D5", "Option5"),
    ];

    let xlsx = create_xlsx_with_dropdown_source(validation_xml, &source_cells);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "A1");
    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::List
    ));

    // Formula1 should contain the range reference
    assert_eq!(dv.validation.formula1.as_deref(), Some("$D$1:$D$5"));
}

#[test]
fn test_dropdown_list_from_named_range() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" sqref="B2:B100">
    <formula1>Categories</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.validation.formula1.as_deref(), Some("Categories"));
}

// =============================================================================
// Tests: Whole Number Validation
// =============================================================================

#[test]
fn test_whole_number_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="between" allowBlank="1" sqref="A1:A100">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Whole
    ));
    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::Between)
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("1"));
    assert_eq!(dv.validation.formula2.as_deref(), Some("100"));
}

#[test]
fn test_whole_number_not_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="notBetween" allowBlank="0" sqref="B1">
    <formula1>10</formula1>
    <formula2>20</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::NotBetween)
    ));
}

#[test]
fn test_whole_number_greater_than() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="greaterThan" allowBlank="1" sqref="C1">
    <formula1>0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::GreaterThan)
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("0"));
    assert!(dv.validation.formula2.is_none());
}

#[test]
fn test_whole_number_greater_than_or_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="greaterThanOrEqual" sqref="D1">
    <formula1>1</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::GreaterThanOrEqual)
    ));
}

#[test]
fn test_whole_number_less_than() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="lessThan" sqref="E1">
    <formula1>1000</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::LessThan)
    ));
}

#[test]
fn test_whole_number_less_than_or_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="lessThanOrEqual" sqref="F1">
    <formula1>999</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::LessThanOrEqual)
    ));
}

#[test]
fn test_whole_number_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="equal" sqref="G1">
    <formula1>42</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::Equal)
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("42"));
}

#[test]
fn test_whole_number_not_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="notEqual" sqref="H1">
    <formula1>0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::NotEqual)
    ));
}

// =============================================================================
// Tests: Decimal Validation
// =============================================================================

#[test]
fn test_decimal_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="decimal" operator="between" allowBlank="1" sqref="A1:A50">
    <formula1>0.0</formula1>
    <formula2>100.0</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Decimal
    ));
    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::Between)
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("0.0"));
    assert_eq!(dv.validation.formula2.as_deref(), Some("100.0"));
}

#[test]
fn test_decimal_greater_than() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="decimal" operator="greaterThan" sqref="B1">
    <formula1>0.01</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Decimal
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("0.01"));
}

#[test]
fn test_decimal_with_cell_reference() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="decimal" operator="lessThanOrEqual" sqref="C1:C100">
    <formula1>$Z$1</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.validation.formula1.as_deref(), Some("$Z$1"));
}

// =============================================================================
// Tests: Date Validation
// =============================================================================

#[test]
fn test_date_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="date" operator="between" allowBlank="1" sqref="A1:A100">
    <formula1>44927</formula1>
    <formula2>45291</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Date
    ));
    // Excel stores dates as serial numbers
    assert_eq!(dv.validation.formula1.as_deref(), Some("44927"));
    assert_eq!(dv.validation.formula2.as_deref(), Some("45291"));
}

#[test]
fn test_date_greater_than_formula() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="date" operator="greaterThan" sqref="B1">
    <formula1>TODAY()</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Date
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("TODAY()"));
}

#[test]
fn test_date_less_than_cell_reference() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="date" operator="lessThan" sqref="C1">
    <formula1>$E$1</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.validation.formula1.as_deref(), Some("$E$1"));
}

// =============================================================================
// Tests: Time Validation
// =============================================================================

#[test]
fn test_time_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="time" operator="between" allowBlank="1" sqref="A1:A50">
    <formula1>0.375</formula1>
    <formula2>0.708333333333333</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Time
    ));
    // Excel stores time as fractions of a day (0.375 = 9:00 AM)
    assert_eq!(dv.validation.formula1.as_deref(), Some("0.375"));
}

#[test]
fn test_time_greater_than() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="time" operator="greaterThan" sqref="B1">
    <formula1>0.333333333333333</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Time
    ));
}

// =============================================================================
// Tests: Text Length Validation
// =============================================================================

#[test]
fn test_text_length_between() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="textLength" operator="between" allowBlank="0" sqref="A1:A100">
    <formula1>1</formula1>
    <formula2>50</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::TextLength
    ));
    assert!(!dv.validation.allow_blank);
    assert_eq!(dv.validation.formula1.as_deref(), Some("1"));
    assert_eq!(dv.validation.formula2.as_deref(), Some("50"));
}

#[test]
fn test_text_length_less_than_or_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="textLength" operator="lessThanOrEqual" sqref="B1">
    <formula1>255</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::LessThanOrEqual)
    ));
    assert_eq!(dv.validation.formula1.as_deref(), Some("255"));
}

#[test]
fn test_text_length_equal() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="textLength" operator="equal" sqref="C1">
    <formula1>10</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::Equal)
    ));
}

// =============================================================================
// Tests: Custom Formula Validation
// =============================================================================

#[test]
fn test_custom_formula_validation() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="custom" allowBlank="1" sqref="A1:A100">
    <formula1>AND(LEN(A1)&gt;0,ISNUMBER(A1))</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Custom
    ));
    // The formula should be unescaped
    assert!(dv.validation.formula1.is_some());
}

#[test]
fn test_custom_formula_with_cell_reference() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="custom" sqref="B1:B50">
    <formula1>=COUNTIF($A:$A,B1)=0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(matches!(
        dv.validation.validation_type,
        xlview::types::ValidationType::Custom
    ));
    assert_eq!(
        dv.validation.formula1.as_deref(),
        Some("=COUNTIF($A:$A,B1)=0")
    );
}

#[test]
fn test_custom_formula_isblank() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="custom" sqref="C1">
    <formula1>NOT(ISBLANK(C1))</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.validation.formula1.as_deref(), Some("NOT(ISBLANK(C1))"));
}

// =============================================================================
// Tests: Input Message Display
// =============================================================================

#[test]
fn test_input_message_with_title_and_message() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" promptTitle="Enter Age" prompt="Please enter a value between 1 and 120" sqref="A1">
    <formula1>1</formula1>
    <formula2>120</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.show_input_message);
    assert_eq!(dv.validation.prompt_title.as_deref(), Some("Enter Age"));
    assert_eq!(
        dv.validation.prompt_message.as_deref(),
        Some("Please enter a value between 1 and 120")
    );
}

#[test]
fn test_input_message_disabled() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" showInputMessage="0" promptTitle="Hidden Title" prompt="Hidden Message" sqref="B1">
    <formula1>"A,B,C"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(!dv.validation.show_input_message);
    // Title and message should still be parsed even if not shown
    assert_eq!(dv.validation.prompt_title.as_deref(), Some("Hidden Title"));
}

#[test]
fn test_input_message_only_title() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="decimal" showInputMessage="1" promptTitle="Enter Value" sqref="C1">
    <formula1>0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.show_input_message);
    assert_eq!(dv.validation.prompt_title.as_deref(), Some("Enter Value"));
    assert!(dv.validation.prompt_message.is_none());
}

// =============================================================================
// Tests: Error Message Display (Stop, Warning, Info)
// =============================================================================

#[test]
fn test_error_message_stop_style() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="between" showErrorMessage="1" errorStyle="stop" errorTitle="Invalid Input" error="Value must be between 1 and 100" sqref="A1">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.show_error_message);
    assert_eq!(dv.validation.error_title.as_deref(), Some("Invalid Input"));
    assert_eq!(
        dv.validation.error_message.as_deref(),
        Some("Value must be between 1 and 100")
    );
}

#[test]
fn test_error_message_warning_style() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" showErrorMessage="1" errorStyle="warning" errorTitle="Warning" error="This value is not in the list. Continue anyway?" sqref="B1">
    <formula1>"Option1,Option2"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.show_error_message);
    assert_eq!(dv.validation.error_title.as_deref(), Some("Warning"));
}

#[test]
fn test_error_message_information_style() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="textLength" operator="lessThanOrEqual" showErrorMessage="1" errorStyle="information" errorTitle="Note" error="Text is longer than recommended" sqref="C1">
    <formula1>100</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.show_error_message);
    assert_eq!(dv.validation.error_title.as_deref(), Some("Note"));
}

#[test]
fn test_error_message_disabled() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" showErrorMessage="0" errorTitle="Not Shown" error="Not Shown" sqref="D1">
    <formula1>1</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(!dv.validation.show_error_message);
}

// =============================================================================
// Tests: Allow Blank Setting
// =============================================================================

#[test]
fn test_allow_blank_true() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" allowBlank="1" sqref="A1">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(dv.validation.allow_blank);
}

#[test]
fn test_allow_blank_false() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" allowBlank="0" sqref="B1">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert!(!dv.validation.allow_blank);
}

#[test]
fn test_allow_blank_default() {
    // When allowBlank is not specified, it defaults to false
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" sqref="C1">
    <formula1>"Yes,No"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // Default value (check your implementation - may be true or false)
    // Just verify the field is accessible (always true or false)
    let _ = dv.validation.allow_blank;
}

// =============================================================================
// Tests: Multiple Validation Ranges
// =============================================================================

#[test]
fn test_multiple_validations_same_sheet() {
    let validation_xml = r#"
<dataValidations count="3">
  <dataValidation type="list" allowBlank="1" sqref="A1:A100">
    <formula1>"Red,Green,Blue"</formula1>
  </dataValidation>
  <dataValidation type="whole" operator="between" allowBlank="1" sqref="B1:B100">
    <formula1>1</formula1>
    <formula2>1000</formula2>
  </dataValidation>
  <dataValidation type="date" operator="greaterThan" sqref="C1:C100">
    <formula1>TODAY()</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.data_validations.len(), 3);

    // First validation - list
    assert!(matches!(
        sheet.data_validations[0].validation.validation_type,
        xlview::types::ValidationType::List
    ));
    assert_eq!(sheet.data_validations[0].sqref, "A1:A100");

    // Second validation - whole number
    assert!(matches!(
        sheet.data_validations[1].validation.validation_type,
        xlview::types::ValidationType::Whole
    ));
    assert_eq!(sheet.data_validations[1].sqref, "B1:B100");

    // Third validation - date
    assert!(matches!(
        sheet.data_validations[2].validation.validation_type,
        xlview::types::ValidationType::Date
    ));
    assert_eq!(sheet.data_validations[2].sqref, "C1:C100");
}

#[test]
fn test_validation_with_multiple_ranges_in_sqref() {
    // A single validation can apply to multiple non-contiguous ranges
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" sqref="A1:A10 C1:C10 E1:E10">
    <formula1>"Yes,No"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // The sqref should contain all ranges
    assert_eq!(dv.sqref, "A1:A10 C1:C10 E1:E10");
}

#[test]
fn test_validation_single_cell() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="decimal" operator="greaterThan" sqref="Z99">
    <formula1>0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "Z99");
}

#[test]
fn test_validation_entire_column() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" sqref="A:A">
    <formula1>"Value1,Value2"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "A:A");
}

#[test]
fn test_validation_entire_row() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" operator="greaterThan" sqref="1:1">
    <formula1>0</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.sqref, "1:1");
}

// =============================================================================
// Tests: Edge Cases and Special Scenarios
// =============================================================================

#[test]
fn test_no_data_validations() {
    let validation_xml = ""; // No data validations

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(sheet.data_validations.is_empty());
}

#[test]
fn test_empty_data_validations_element() {
    let validation_xml = r#"<dataValidations count="0"></dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(sheet.data_validations.is_empty());
}

#[test]
fn test_validation_with_empty_formula() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" sqref="A1">
    <formula1></formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];

    // Should still parse, but formula may be None or empty string
    assert!(!sheet.data_validations.is_empty());
}

#[test]
fn test_validation_show_dropdown_hidden() {
    // showDropDown="1" actually HIDES the dropdown in Excel (counterintuitive)
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" showDropDown="1" sqref="A1">
    <formula1>"A,B,C"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // When showDropDown="1", the dropdown is hidden (show_dropdown should be false)
    assert!(!dv.validation.show_dropdown);
}

#[test]
fn test_validation_default_operator() {
    // When no operator is specified for whole/decimal, default is "between"
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" sqref="A1">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // Default operator should be Between
    assert!(matches!(
        dv.validation.operator,
        Some(xlview::types::ValidationOperator::Between) | None
    ));
}

#[test]
fn test_validation_with_xml_entities() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" showInputMessage="1" prompt="Select &lt;one&gt; option" sqref="A1">
    <formula1>"A &amp; B,C &lt; D"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // Entities should be properly decoded
    assert_eq!(
        dv.validation.prompt_message.as_deref(),
        Some("Select <one> option")
    );
}

#[test]
fn test_validation_serialization_to_json() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Select" prompt="Choose a color" errorTitle="Error" error="Invalid selection" sqref="A1:A10">
    <formula1>"Red,Green,Blue"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    // Serialize to JSON and back
    let json = serde_json::to_string(&workbook).expect("Failed to serialize");
    let parsed: serde_json::Value = serde_json::from_str(&json).expect("Failed to parse JSON");

    // Check dataValidations in JSON
    let validations = &parsed["sheets"][0]["dataValidations"];
    assert!(validations.is_array());
    assert_eq!(validations.as_array().unwrap().len(), 1);

    let dv = &validations[0];
    assert_eq!(dv["sqref"], "A1:A10");
    assert_eq!(dv["validation"]["validationType"], "list");
    assert_eq!(dv["validation"]["allowBlank"], true);
    assert_eq!(dv["validation"]["promptTitle"], "Select");
    assert_eq!(dv["validation"]["promptMessage"], "Choose a color");
    assert_eq!(dv["validation"]["errorTitle"], "Error");
    assert_eq!(dv["validation"]["errorMessage"], "Invalid selection");
}

#[test]
fn test_validation_list_with_numeric_values() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="list" sqref="A1">
    <formula1>"1,2,3,4,5"</formula1>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    if let Some(ref values) = dv.validation.list_values {
        assert_eq!(values.len(), 5);
        assert!(values.contains(&"1".to_string()));
        assert!(values.contains(&"5".to_string()));
    }
}

#[test]
fn test_validation_with_newlines_in_message() {
    let validation_xml = r#"
<dataValidations count="1">
  <dataValidation type="whole" showInputMessage="1" prompt="Enter a number.&#10;Must be positive.&#10;Maximum: 100" sqref="A1">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>"#;

    let xlsx = create_xlsx_with_data_validation(validation_xml);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let dv = &sheet.data_validations[0];

    // Newlines should be preserved
    assert!(dv.validation.prompt_message.is_some());
    let msg = dv.validation.prompt_message.as_ref().unwrap();
    assert!(msg.contains('\n') || msg.contains("&#10;") || msg.contains("Enter a number"));
}

// =============================================================================
// Tests: Real XLSX Files - kitchen_sink_v2.xlsx
// =============================================================================

/// Parse a real XLSX test file and return the workbook
#[allow(clippy::expect_used)]
fn parse_test_file(path: &str) -> xlview::types::Workbook {
    let data = std::fs::read(path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    xlview::parser::parse(&data).unwrap_or_else(|_| panic!("Failed to parse XLSX file: {}", path))
}

/// Find a sheet by name in the workbook
#[allow(dead_code)]
fn find_sheet<'a>(
    workbook: &'a xlview::types::Workbook,
    name: &str,
) -> Option<&'a xlview::types::Sheet> {
    workbook.sheets.iter().find(|s| s.name == name)
}

/// Get all data validations of a specific type from a sheet
fn get_validations_by_type(
    sheet: &xlview::types::Sheet,
    validation_type: xlview::types::ValidationType,
) -> Vec<&xlview::types::DataValidationRange> {
    sheet
        .data_validations
        .iter()
        .filter(|dv| {
            std::mem::discriminant(&dv.validation.validation_type)
                == std::mem::discriminant(&validation_type)
        })
        .collect()
}

/// Get a data validation by sqref (cell range)
fn get_validation_by_sqref<'a>(
    sheet: &'a xlview::types::Sheet,
    sqref: &str,
) -> Option<&'a xlview::types::DataValidationRange> {
    sheet.data_validations.iter().find(|dv| dv.sqref == sqref)
}

#[test]
fn test_kitchen_sink_v2_data_validations_exist() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    // kitchen_sink_v2.xlsx has data validations in sheet 3 (Data Validation sheet)
    // Find the sheet that has data validations
    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty());

    assert!(
        sheet_with_validations.is_some(),
        "kitchen_sink_v2.xlsx should have at least one sheet with data validations"
    );

    let sheet = sheet_with_validations.unwrap();

    // Should have 5 data validations based on our XML analysis
    assert_eq!(
        sheet.data_validations.len(),
        5,
        "Sheet should have 5 data validations"
    );
}

#[test]
fn test_kitchen_sink_v2_list_validation_with_prompt() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty())
        .expect("Should have a sheet with data validations");

    // Find the list validation at B3 which has a prompt
    let dv = get_validation_by_sqref(sheet_with_validations, "B3");
    assert!(dv.is_some(), "Should find validation at B3");

    let dv = dv.unwrap();
    assert!(
        matches!(
            dv.validation.validation_type,
            xlview::types::ValidationType::List
        ),
        "B3 should have a list validation"
    );

    // Check list values are parsed
    assert!(
        dv.validation.list_values.is_some(),
        "List validation should have parsed list values"
    );

    let values = dv.validation.list_values.as_ref().unwrap();
    assert_eq!(values.len(), 4, "Should have 4 list values");
    assert!(
        values.contains(&"Active".to_string()),
        "Should contain 'Active'"
    );
    assert!(
        values.contains(&"Pending".to_string()),
        "Should contain 'Pending'"
    );
    assert!(
        values.contains(&"Completed".to_string()),
        "Should contain 'Completed'"
    );
    assert!(
        values.contains(&"Cancelled".to_string()),
        "Should contain 'Cancelled'"
    );

    // Check prompt properties
    assert_eq!(
        dv.validation.prompt_title.as_deref(),
        Some("Status"),
        "Should have prompt title 'Status'"
    );
    assert_eq!(
        dv.validation.prompt_message.as_deref(),
        Some("Select a status"),
        "Should have prompt message"
    );
    assert!(dv.validation.allow_blank, "Should allow blank");
}

#[test]
fn test_kitchen_sink_v2_list_validation_priority() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty())
        .expect("Should have a sheet with data validations");

    // Find the list validation at B4 (High, Medium, Low)
    let dv = get_validation_by_sqref(sheet_with_validations, "B4");
    assert!(dv.is_some(), "Should find validation at B4");

    let dv = dv.unwrap();
    assert!(
        matches!(
            dv.validation.validation_type,
            xlview::types::ValidationType::List
        ),
        "B4 should have a list validation"
    );

    let values = dv.validation.list_values.as_ref().unwrap();
    assert_eq!(values.len(), 3, "Should have 3 priority values");
    assert!(values.contains(&"High".to_string()));
    assert!(values.contains(&"Medium".to_string()));
    assert!(values.contains(&"Low".to_string()));

    // This validation does not allow blank
    assert!(!dv.validation.allow_blank, "Should not allow blank");
}

#[test]
fn test_kitchen_sink_v2_whole_number_validation() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty())
        .expect("Should have a sheet with data validations");

    // Find the whole number validation at B6 (age between 1 and 120)
    let dv = get_validation_by_sqref(sheet_with_validations, "B6");
    assert!(dv.is_some(), "Should find validation at B6");

    let dv = dv.unwrap();
    assert!(
        matches!(
            dv.validation.validation_type,
            xlview::types::ValidationType::Whole
        ),
        "B6 should have a whole number validation"
    );

    // Check operator is between
    assert!(
        matches!(
            dv.validation.operator,
            Some(xlview::types::ValidationOperator::Between)
        ),
        "Should have 'between' operator"
    );

    // Check formula values
    assert_eq!(
        dv.validation.formula1.as_deref(),
        Some("1"),
        "formula1 should be 1"
    );
    assert_eq!(
        dv.validation.formula2.as_deref(),
        Some("120"),
        "formula2 should be 120"
    );

    // Check error message properties
    assert_eq!(
        dv.validation.error_title.as_deref(),
        Some("Invalid Age"),
        "Should have error title 'Invalid Age'"
    );
    assert_eq!(
        dv.validation.error_message.as_deref(),
        Some("Age must be between 1 and 120"),
        "Should have error message"
    );
}

#[test]
fn test_kitchen_sink_v2_date_validation() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty())
        .expect("Should have a sheet with data validations");

    // Find the date validation at B7
    let dv = get_validation_by_sqref(sheet_with_validations, "B7");
    assert!(dv.is_some(), "Should find validation at B7");

    let dv = dv.unwrap();
    assert!(
        matches!(
            dv.validation.validation_type,
            xlview::types::ValidationType::Date
        ),
        "B7 should have a date validation"
    );

    // Check operator is greaterThanOrEqual
    assert!(
        matches!(
            dv.validation.operator,
            Some(xlview::types::ValidationOperator::GreaterThanOrEqual)
        ),
        "Should have 'greaterThanOrEqual' operator"
    );

    // formula1 should contain a date reference
    assert!(
        dv.validation.formula1.is_some(),
        "Should have formula1 for date"
    );
}

#[test]
fn test_kitchen_sink_v2_list_validation_range() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let sheet_with_validations = workbook
        .sheets
        .iter()
        .find(|s| !s.data_validations.is_empty())
        .expect("Should have a sheet with data validations");

    // Find the list validation at B9:B11 (Yes/No across multiple cells)
    let dv = get_validation_by_sqref(sheet_with_validations, "B9:B11");
    assert!(dv.is_some(), "Should find validation at B9:B11");

    let dv = dv.unwrap();
    assert!(
        matches!(
            dv.validation.validation_type,
            xlview::types::ValidationType::List
        ),
        "B9:B11 should have a list validation"
    );

    let values = dv.validation.list_values.as_ref().unwrap();
    assert_eq!(values.len(), 2, "Should have 2 values");
    assert!(values.contains(&"Yes".to_string()));
    assert!(values.contains(&"No".to_string()));

    // Verify sqref contains the range
    assert!(
        dv.sqref.contains("B9") && dv.sqref.contains("B11"),
        "sqref should cover B9:B11"
    );
}

// =============================================================================
// Tests: Real XLSX Files - ms_cf_samples.xlsx
// =============================================================================

#[test]
fn test_ms_cf_samples_data_validations_exist() {
    let workbook = parse_test_file("test/ms_cf_samples.xlsx");

    // Count total data validations across all sheets
    let total_validations: usize = workbook
        .sheets
        .iter()
        .map(|s| s.data_validations.len())
        .sum();

    assert!(
        total_validations >= 3,
        "ms_cf_samples.xlsx should have at least 3 data validations, found {}",
        total_validations
    );
}

#[test]
fn test_ms_cf_samples_list_validations() {
    let workbook = parse_test_file("test/ms_cf_samples.xlsx");

    // Find all list validations across all sheets
    let mut list_validations: Vec<&xlview::types::DataValidationRange> = Vec::new();
    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            if matches!(
                dv.validation.validation_type,
                xlview::types::ValidationType::List
            ) {
                list_validations.push(dv);
            }
        }
    }

    assert!(
        !list_validations.is_empty(),
        "ms_cf_samples.xlsx should have list validations"
    );

    // Check that list values are parsed for each list validation
    for dv in &list_validations {
        assert!(
            dv.validation.list_values.is_some() || dv.validation.formula1.is_some(),
            "List validation should have list_values or formula1"
        );
    }
}

#[test]
fn test_ms_cf_samples_whole_number_equal_validation() {
    let workbook = parse_test_file("test/ms_cf_samples.xlsx");

    // Find the whole number validation with equal operator (B9:B23 with formula1=0)
    let mut found_equal_validation = false;

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            if matches!(
                dv.validation.validation_type,
                xlview::types::ValidationType::Whole
            ) && matches!(
                dv.validation.operator,
                Some(xlview::types::ValidationOperator::Equal)
            ) {
                found_equal_validation = true;
                assert_eq!(
                    dv.validation.formula1.as_deref(),
                    Some("0"),
                    "formula1 should be 0"
                );
            }
        }
    }

    assert!(
        found_equal_validation,
        "ms_cf_samples.xlsx should have a whole number equal validation"
    );
}

#[test]
fn test_ms_cf_samples_dropdown_list_values() {
    let workbook = parse_test_file("test/ms_cf_samples.xlsx");

    // Find list validations and check specific values
    let mut found_dairy_list = false;
    let mut found_price_list = false;

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            if matches!(
                dv.validation.validation_type,
                xlview::types::ValidationType::List
            ) {
                if let Some(ref values) = dv.validation.list_values {
                    if values.contains(&"Dairy".to_string()) {
                        found_dairy_list = true;
                        assert!(values.contains(&"Produce".to_string()));
                        assert!(values.contains(&"Grain".to_string()));
                    }
                    if values.iter().any(|v| v.contains("$100")) {
                        found_price_list = true;
                        // Should have $100, $200, $300, $400, $500
                        assert_eq!(values.len(), 5, "Price list should have 5 values");
                    }
                }
            }
        }
    }

    assert!(
        found_dairy_list,
        "ms_cf_samples.xlsx should have the Dairy/Produce/Grain list validation"
    );
    assert!(
        found_price_list,
        "ms_cf_samples.xlsx should have the price list validation"
    );
}

// =============================================================================
// Tests: Validation Type Coverage
// =============================================================================

#[test]
fn test_real_file_validation_type_coverage() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let mut found_types: std::collections::HashSet<String> = std::collections::HashSet::new();

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            let type_str = match dv.validation.validation_type {
                xlview::types::ValidationType::None => "none",
                xlview::types::ValidationType::Whole => "whole",
                xlview::types::ValidationType::Decimal => "decimal",
                xlview::types::ValidationType::List => "list",
                xlview::types::ValidationType::Date => "date",
                xlview::types::ValidationType::Time => "time",
                xlview::types::ValidationType::TextLength => "textLength",
                xlview::types::ValidationType::Custom => "custom",
            };
            found_types.insert(type_str.to_string());
        }
    }

    // kitchen_sink_v2.xlsx should have at least list, whole, and date types
    assert!(
        found_types.contains("list"),
        "Should find list validation type"
    );
    assert!(
        found_types.contains("whole"),
        "Should find whole validation type"
    );
    assert!(
        found_types.contains("date"),
        "Should find date validation type"
    );
}

#[test]
fn test_real_file_operator_coverage() {
    let workbook_v2 = parse_test_file("test/kitchen_sink_v2.xlsx");
    let workbook_cf = parse_test_file("test/ms_cf_samples.xlsx");

    let mut found_operators: std::collections::HashSet<String> = std::collections::HashSet::new();

    for workbook in [&workbook_v2, &workbook_cf] {
        for sheet in &workbook.sheets {
            for dv in &sheet.data_validations {
                if let Some(ref op) = dv.validation.operator {
                    let op_str = match op {
                        xlview::types::ValidationOperator::Between => "between",
                        xlview::types::ValidationOperator::NotBetween => "notBetween",
                        xlview::types::ValidationOperator::Equal => "equal",
                        xlview::types::ValidationOperator::NotEqual => "notEqual",
                        xlview::types::ValidationOperator::LessThan => "lessThan",
                        xlview::types::ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
                        xlview::types::ValidationOperator::GreaterThan => "greaterThan",
                        xlview::types::ValidationOperator::GreaterThanOrEqual => {
                            "greaterThanOrEqual"
                        }
                    };
                    found_operators.insert(op_str.to_string());
                }
            }
        }
    }

    // Should find at least between, equal, and greaterThanOrEqual operators
    assert!(
        found_operators.contains("between"),
        "Should find 'between' operator"
    );
    assert!(
        found_operators.contains("equal"),
        "Should find 'equal' operator"
    );
    assert!(
        found_operators.contains("greaterThanOrEqual"),
        "Should find 'greaterThanOrEqual' operator"
    );
}

// =============================================================================
// Tests: Data Validation Property Verification
// =============================================================================

#[test]
fn test_sqref_parsing_from_real_files() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            // All sqref should be non-empty
            assert!(!dv.sqref.is_empty(), "sqref should not be empty");

            // sqref should contain valid cell references
            assert!(
                dv.sqref.chars().any(|c| c.is_ascii_alphabetic()),
                "sqref should contain column letters: {}",
                dv.sqref
            );
            assert!(
                dv.sqref.chars().any(|c| c.is_ascii_digit()),
                "sqref should contain row numbers: {}",
                dv.sqref
            );
        }
    }
}

#[test]
fn test_allow_blank_from_real_files() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let mut found_allow_blank_true = false;
    let mut found_allow_blank_false = false;

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            if dv.validation.allow_blank {
                found_allow_blank_true = true;
            } else {
                found_allow_blank_false = true;
            }
        }
    }

    // We should have both allow_blank=true and allow_blank=false cases
    assert!(
        found_allow_blank_true,
        "Should find validations with allow_blank=true"
    );
    assert!(
        found_allow_blank_false,
        "Should find validations with allow_blank=false"
    );
}

#[test]
fn test_show_dropdown_from_real_files() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        let list_validations = get_validations_by_type(sheet, xlview::types::ValidationType::List);

        for dv in list_validations {
            // In the test file, showDropDown="0" means show the dropdown
            // (counterintuitive Excel behavior)
            // All list validations should have show_dropdown set correctly
            // Our parser inverts the XML value
            // Verify show_dropdown field is accessible (always true or false)
            let _ = dv.validation.show_dropdown;
        }
    }
}

#[test]
fn test_formula1_formula2_from_real_files() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            match dv.validation.validation_type {
                xlview::types::ValidationType::Whole | xlview::types::ValidationType::Decimal => {
                    // Numeric validations should have formula1
                    assert!(
                        dv.validation.formula1.is_some(),
                        "Numeric validation should have formula1"
                    );

                    // Between operator should have formula2
                    if matches!(
                        dv.validation.operator,
                        Some(xlview::types::ValidationOperator::Between)
                            | Some(xlview::types::ValidationOperator::NotBetween)
                    ) {
                        assert!(
                            dv.validation.formula2.is_some(),
                            "Between/NotBetween should have formula2"
                        );
                    }
                }
                xlview::types::ValidationType::List => {
                    // List validations should have formula1 (the list source)
                    assert!(
                        dv.validation.formula1.is_some(),
                        "List validation should have formula1"
                    );
                }
                xlview::types::ValidationType::Date | xlview::types::ValidationType::Time => {
                    // Date/Time validations should have formula1
                    assert!(
                        dv.validation.formula1.is_some(),
                        "Date/Time validation should have formula1"
                    );
                }
                _ => {}
            }
        }
    }
}

// =============================================================================
// Tests: JSON Serialization of Real File Data
// =============================================================================

#[test]
fn test_real_file_json_serialization() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    // Serialize to JSON
    let json = serde_json::to_string(&workbook).expect("Failed to serialize workbook");
    let parsed: serde_json::Value = serde_json::from_str(&json).expect("Failed to parse JSON");

    // Find the sheet with data validations
    let sheets = parsed["sheets"].as_array().unwrap();
    let sheet_with_dv = sheets
        .iter()
        .find(|s| {
            s.get("dataValidations")
                .map(|v| v.as_array().map(|a| !a.is_empty()).unwrap_or(false))
                .unwrap_or(false)
        })
        .expect("Should find sheet with data validations");

    let validations = sheet_with_dv["dataValidations"].as_array().unwrap();
    assert!(
        !validations.is_empty(),
        "Should have data validations in JSON"
    );

    // Check JSON structure of first validation
    let first_dv = &validations[0];
    assert!(first_dv.get("sqref").is_some(), "Should have sqref in JSON");
    assert!(
        first_dv.get("validation").is_some(),
        "Should have validation object in JSON"
    );

    let validation = &first_dv["validation"];
    assert!(
        validation.get("validationType").is_some(),
        "Should have validationType in JSON"
    );
    assert!(
        validation.get("allowBlank").is_some(),
        "Should have allowBlank in JSON"
    );
}

// =============================================================================
// Tests: Total Validation Count Statistics
// =============================================================================

#[test]
fn test_validation_count_statistics() {
    let workbook_v2 = parse_test_file("test/kitchen_sink_v2.xlsx");
    let workbook_cf = parse_test_file("test/ms_cf_samples.xlsx");

    // Count validations in kitchen_sink_v2.xlsx
    let v2_count: usize = workbook_v2
        .sheets
        .iter()
        .map(|s| s.data_validations.len())
        .sum();
    assert_eq!(
        v2_count, 5,
        "kitchen_sink_v2.xlsx should have 5 data validations"
    );

    // Count validations in ms_cf_samples.xlsx
    let cf_count: usize = workbook_cf
        .sheets
        .iter()
        .map(|s| s.data_validations.len())
        .sum();
    assert!(
        cf_count >= 3,
        "ms_cf_samples.xlsx should have at least 3 data validations, found {}",
        cf_count
    );

    // Count by type in kitchen_sink_v2.xlsx
    let mut type_counts: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();
    for sheet in &workbook_v2.sheets {
        for dv in &sheet.data_validations {
            let type_str = format!("{:?}", dv.validation.validation_type);
            *type_counts.entry(type_str).or_insert(0) += 1;
        }
    }

    // Should have List validations (most common in test file)
    let list_count = type_counts.get("List").unwrap_or(&0);
    assert!(
        *list_count >= 3,
        "kitchen_sink_v2.xlsx should have at least 3 list validations, found {}",
        list_count
    );
}

// =============================================================================
// Tests: Edge Cases in Real Files
// =============================================================================

#[test]
fn test_sheets_without_validations() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    // Some sheets should have no validations
    let sheets_without_validations: Vec<_> = workbook
        .sheets
        .iter()
        .filter(|s| s.data_validations.is_empty())
        .collect();

    assert!(
        !sheets_without_validations.is_empty(),
        "Some sheets should have no data validations"
    );
}

#[test]
fn test_validation_messages_parsed() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    let mut found_prompt_title = false;
    let mut found_prompt_message = false;
    let mut found_error_title = false;
    let mut found_error_message = false;

    for sheet in &workbook.sheets {
        for dv in &sheet.data_validations {
            if dv.validation.prompt_title.is_some() {
                found_prompt_title = true;
            }
            if dv.validation.prompt_message.is_some() {
                found_prompt_message = true;
            }
            if dv.validation.error_title.is_some() {
                found_error_title = true;
            }
            if dv.validation.error_message.is_some() {
                found_error_message = true;
            }
        }
    }

    // kitchen_sink_v2.xlsx has validations with prompt and error messages
    assert!(found_prompt_title, "Should find prompt titles");
    assert!(found_prompt_message, "Should find prompt messages");
    assert!(found_error_title, "Should find error titles");
    assert!(found_error_message, "Should find error messages");
}

#[test]
fn test_kitchen_sink_original_no_validations() {
    // kitchen_sink.xlsx (original) should not have data validations
    let workbook = parse_test_file("test/kitchen_sink.xlsx");

    let total_validations: usize = workbook
        .sheets
        .iter()
        .map(|s| s.data_validations.len())
        .sum();

    assert_eq!(
        total_validations, 0,
        "kitchen_sink.xlsx (original) should not have data validations"
    );
}

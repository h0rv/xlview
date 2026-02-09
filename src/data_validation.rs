//! Data validation parsing module
//! This module handles parsing of data validation rules from XLSX files.

use crate::types::{DataValidation, DataValidationRange, ValidationOperator, ValidationType};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

/// Parse data validation element
///
/// Parses a `<dataValidation>` element and its children from XLSX sheet XML.
///
/// # Example XML
/// ```xml
/// <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="A1:A100"
///     showInputMessage="1" showErrorMessage="1" errorTitle="Error" error="Invalid value"
///     promptTitle="Select" prompt="Choose an option">
///   <formula1>"Option1,Option2,Option3"</formula1>
///   <formula2>100</formula2>
/// </dataValidation>
/// ```
///
/// # Supported attributes
/// - `sqref`: Cell range(s) the validation applies to (e.g., "A1:A100" or "A1:A100 B1:B100")
/// - `type`: Validation type (whole, decimal, list, date, time, textLength, custom)
/// - `operator`: Comparison operator (between, notBetween, equal, notEqual, lessThan, etc.)
/// - `allowBlank`: Whether blank values are allowed ("1" = true)
/// - `showDropDown`: Whether to hide dropdown for list type ("1" = hide, counterintuitive)
/// - `showInputMessage`: Whether to show input prompt ("1" = true)
/// - `showErrorMessage`: Whether to show error message on invalid input ("1" = true)
/// - `errorTitle`: Title for error dialog
/// - `error`: Error message text
/// - `promptTitle`: Title for input prompt
/// - `prompt`: Input prompt message
///
/// # Child elements
/// - `<formula1>`: First formula/value for validation (required for most types)
/// - `<formula2>`: Second formula/value (for between/notBetween operators)
pub fn parse_data_validation<R: BufRead>(
    e: &BytesStart,
    xml: &mut Reader<R>,
) -> Option<DataValidationRange> {
    let mut sqref = String::new();
    let mut validation_type = ValidationType::None;
    let mut operator: Option<ValidationOperator> = None;
    let mut allow_blank = false;
    let mut show_dropdown = true;
    let mut show_input_message = false;
    let mut show_error_message = false;
    let mut error_title: Option<String> = None;
    let mut error_message: Option<String> = None;
    let mut prompt_title: Option<String> = None;
    let mut prompt_message: Option<String> = None;

    // Parse attributes from the opening tag
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"sqref" => {
                sqref = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"type" => {
                validation_type =
                    parse_validation_type(std::str::from_utf8(&attr.value).unwrap_or(""));
            }
            b"operator" => {
                operator = Some(parse_validation_operator(
                    std::str::from_utf8(&attr.value).unwrap_or(""),
                ));
            }
            b"allowBlank" => {
                allow_blank = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"showDropDown" => {
                // Note: in XLSX, showDropDown="1" means HIDE the dropdown (counterintuitive)
                show_dropdown = std::str::from_utf8(&attr.value).unwrap_or("0") != "1";
            }
            b"showInputMessage" => {
                show_input_message = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"showErrorMessage" => {
                show_error_message = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"errorTitle" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("");
                if !val.is_empty() {
                    error_title = Some(val.to_string());
                }
            }
            b"error" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("");
                if !val.is_empty() {
                    error_message = Some(val.to_string());
                }
            }
            b"promptTitle" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("");
                if !val.is_empty() {
                    prompt_title = Some(val.to_string());
                }
            }
            b"prompt" => {
                if let Ok(val) = attr.unescape_value() {
                    if !val.is_empty() {
                        prompt_message = Some(val.to_string());
                    }
                }
            }
            _ => {}
        }
    }

    // sqref is required
    if sqref.is_empty() {
        return None;
    }

    // Parse child elements (formula1, formula2)
    let mut formula1: Option<String> = None;
    let mut formula2: Option<String> = None;
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "formula1" => {
                        formula1 = read_element_text(xml);
                    }
                    "formula2" => {
                        formula2 = read_element_text(xml);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref inner)) => {
                if inner.local_name().as_ref() == b"dataValidation" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // For list type, parse the list values from formula1
    let list_values = if matches!(validation_type, ValidationType::List) {
        formula1.as_ref().and_then(|f| parse_list_values(f))
    } else {
        None
    };

    Some(DataValidationRange {
        sqref,
        validation: DataValidation {
            validation_type,
            operator,
            formula1,
            formula2,
            allow_blank,
            show_dropdown,
            show_input_message,
            show_error_message,
            error_title,
            error_message,
            prompt_title,
            prompt_message,
            list_values,
        },
    })
}

/// Parse validation type from string
fn parse_validation_type(s: &str) -> ValidationType {
    match s {
        "whole" => ValidationType::Whole,
        "decimal" => ValidationType::Decimal,
        "list" => ValidationType::List,
        "date" => ValidationType::Date,
        "time" => ValidationType::Time,
        "textLength" => ValidationType::TextLength,
        "custom" => ValidationType::Custom,
        _ => ValidationType::None,
    }
}

/// Parse validation operator from string
fn parse_validation_operator(s: &str) -> ValidationOperator {
    match s {
        "between" => ValidationOperator::Between,
        "notBetween" => ValidationOperator::NotBetween,
        "equal" => ValidationOperator::Equal,
        "notEqual" => ValidationOperator::NotEqual,
        "lessThan" => ValidationOperator::LessThan,
        "lessThanOrEqual" => ValidationOperator::LessThanOrEqual,
        "greaterThan" => ValidationOperator::GreaterThan,
        "greaterThanOrEqual" => ValidationOperator::GreaterThanOrEqual,
        _ => ValidationOperator::Between, // default
    }
}

/// Read text content of current element
fn read_element_text<R: BufRead>(xml: &mut Reader<R>) -> Option<String> {
    let mut text_buf = Vec::new();
    match xml.read_event_into(&mut text_buf) {
        Ok(Event::Text(text)) => text.unescape().ok().map(|s| s.to_string()),
        Ok(Event::CData(cdata)) => std::str::from_utf8(&cdata).ok().map(|s| s.to_string()),
        _ => None,
    }
}

/// Parse list values from formula1
///
/// List values in XLSX can be:
/// - Inline quoted comma-separated: `"Option1,Option2,Option3"`
/// - Cell reference: `$A$1:$A$10` or `Sheet1!$A$1:$A$10`
///
/// This function only parses inline quoted values. Cell references are left as-is
/// in the formula1 field for the consumer to resolve.
fn parse_list_values(formula: &str) -> Option<Vec<String>> {
    let trimmed = formula.trim();

    // Check if it's a quoted inline list
    if trimmed.starts_with('"') && trimmed.ends_with('"') {
        // Remove surrounding quotes and split by comma
        let inner = &trimmed[1..trimmed.len() - 1];
        let values: Vec<String> = inner
            .split(',')
            .map(|v| v.trim().to_string())
            .filter(|v| !v.is_empty())
            .collect();

        if values.is_empty() {
            None
        } else {
            Some(values)
        }
    } else {
        // It's a cell reference or other formula - don't parse as list values
        None
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_validation_type() {
        assert!(matches!(
            parse_validation_type("whole"),
            ValidationType::Whole
        ));
        assert!(matches!(
            parse_validation_type("decimal"),
            ValidationType::Decimal
        ));
        assert!(matches!(
            parse_validation_type("list"),
            ValidationType::List
        ));
        assert!(matches!(
            parse_validation_type("date"),
            ValidationType::Date
        ));
        assert!(matches!(
            parse_validation_type("time"),
            ValidationType::Time
        ));
        assert!(matches!(
            parse_validation_type("textLength"),
            ValidationType::TextLength
        ));
        assert!(matches!(
            parse_validation_type("custom"),
            ValidationType::Custom
        ));
        assert!(matches!(
            parse_validation_type("unknown"),
            ValidationType::None
        ));
        assert!(matches!(parse_validation_type(""), ValidationType::None));
    }

    #[test]
    fn test_parse_validation_operator() {
        assert!(matches!(
            parse_validation_operator("between"),
            ValidationOperator::Between
        ));
        assert!(matches!(
            parse_validation_operator("notBetween"),
            ValidationOperator::NotBetween
        ));
        assert!(matches!(
            parse_validation_operator("equal"),
            ValidationOperator::Equal
        ));
        assert!(matches!(
            parse_validation_operator("notEqual"),
            ValidationOperator::NotEqual
        ));
        assert!(matches!(
            parse_validation_operator("lessThan"),
            ValidationOperator::LessThan
        ));
        assert!(matches!(
            parse_validation_operator("lessThanOrEqual"),
            ValidationOperator::LessThanOrEqual
        ));
        assert!(matches!(
            parse_validation_operator("greaterThan"),
            ValidationOperator::GreaterThan
        ));
        assert!(matches!(
            parse_validation_operator("greaterThanOrEqual"),
            ValidationOperator::GreaterThanOrEqual
        ));
        // Default to Between for unknown
        assert!(matches!(
            parse_validation_operator("unknown"),
            ValidationOperator::Between
        ));
    }

    #[test]
    fn test_parse_list_values_quoted() {
        let result = parse_list_values("\"Option1,Option2,Option3\"");
        assert!(result.is_some());
        let values = result.unwrap();
        assert_eq!(values.len(), 3);
        assert_eq!(values[0], "Option1");
        assert_eq!(values[1], "Option2");
        assert_eq!(values[2], "Option3");
    }

    #[test]
    fn test_parse_list_values_with_spaces() {
        let result = parse_list_values("\" Yes , No , Maybe \"");
        assert!(result.is_some());
        let values = result.unwrap();
        assert_eq!(values.len(), 3);
        assert_eq!(values[0], "Yes");
        assert_eq!(values[1], "No");
        assert_eq!(values[2], "Maybe");
    }

    #[test]
    fn test_parse_list_values_cell_reference() {
        // Cell references should not be parsed as list values
        let result = parse_list_values("$A$1:$A$10");
        assert!(result.is_none());

        let result = parse_list_values("Sheet1!$A$1:$A$10");
        assert!(result.is_none());
    }

    #[test]
    fn test_parse_list_values_empty() {
        let result = parse_list_values("\"\"");
        assert!(result.is_none());
    }

    #[test]
    fn test_parse_data_validation_list() {
        let xml_str = r#"<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="A1:A100" showInputMessage="1" promptTitle="Select" prompt="Choose one">
            <formula1>"Red,Green,Blue"</formula1>
        </dataValidation>"#;

        let mut reader = Reader::from_str(xml_str);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"dataValidation" {
                        let result = parse_data_validation(e, &mut reader);
                        assert!(result.is_some());
                        let dv = result.unwrap();

                        assert_eq!(dv.sqref, "A1:A100");
                        assert!(matches!(
                            dv.validation.validation_type,
                            ValidationType::List
                        ));
                        assert!(dv.validation.allow_blank);
                        assert!(dv.validation.show_dropdown);
                        assert!(dv.validation.show_input_message);
                        assert_eq!(dv.validation.prompt_title, Some("Select".to_string()));
                        assert_eq!(dv.validation.prompt_message, Some("Choose one".to_string()));
                        assert_eq!(
                            dv.validation.formula1,
                            Some("\"Red,Green,Blue\"".to_string())
                        );

                        let list_values = dv.validation.list_values.unwrap();
                        assert_eq!(list_values.len(), 3);
                        assert_eq!(list_values[0], "Red");
                        assert_eq!(list_values[1], "Green");
                        assert_eq!(list_values[2], "Blue");
                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn test_parse_data_validation_whole_number() {
        let xml_str = r#"<dataValidation type="whole" operator="between" allowBlank="0" showErrorMessage="1" errorTitle="Error" error="Value must be between 1 and 100" sqref="B1:B50">
            <formula1>1</formula1>
            <formula2>100</formula2>
        </dataValidation>"#;

        let mut reader = Reader::from_str(xml_str);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"dataValidation" {
                        let result = parse_data_validation(e, &mut reader);
                        assert!(result.is_some());
                        let dv = result.unwrap();

                        assert_eq!(dv.sqref, "B1:B50");
                        assert!(matches!(
                            dv.validation.validation_type,
                            ValidationType::Whole
                        ));
                        assert!(matches!(
                            dv.validation.operator,
                            Some(ValidationOperator::Between)
                        ));
                        assert!(!dv.validation.allow_blank);
                        assert!(dv.validation.show_error_message);
                        assert_eq!(dv.validation.error_title, Some("Error".to_string()));
                        assert_eq!(
                            dv.validation.error_message,
                            Some("Value must be between 1 and 100".to_string())
                        );
                        assert_eq!(dv.validation.formula1, Some("1".to_string()));
                        assert_eq!(dv.validation.formula2, Some("100".to_string()));
                        assert!(dv.validation.list_values.is_none());
                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    #[test]
    fn test_parse_data_validation_empty_sqref() {
        let xml_str = r#"<dataValidation type="list" sqref="">
            <formula1>"A,B,C"</formula1>
        </dataValidation>"#;

        let mut reader = Reader::from_str(xml_str);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.local_name().as_ref() == b"dataValidation" {
                        let result = parse_data_validation(e, &mut reader);
                        // Should return None because sqref is empty
                        assert!(result.is_none());
                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
}

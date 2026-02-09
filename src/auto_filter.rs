//! Auto-filter parsing module
//! This module handles parsing of auto-filter settings from XLSX files.

use crate::cell_ref::parse_cell_ref_or_default;
use crate::types::{AutoFilter, CustomFilter, CustomFilterOperator, FilterColumn, FilterType};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

/// Parse autoFilter element
///
/// # XML Structure
/// ```xml
/// <autoFilter ref="A1:D10">
///   <filterColumn colId="0" hiddenButton="0" showButton="1">
///     <filters><filter val="Value1"/><filter val="Value2"/></filters>
///   </filterColumn>
///   <filterColumn colId="1">
///     <customFilters and="1">
///       <customFilter operator="greaterThan" val="100"/>
///     </customFilters>
///   </filterColumn>
///   <filterColumn colId="2">
///     <colorFilter dxfId="0" cellColor="1"/>
///   </filterColumn>
///   <filterColumn colId="3">
///     <iconFilter iconSet="3Arrows" iconId="0"/>
///   </filterColumn>
///   <filterColumn colId="4">
///     <dynamicFilter type="aboveAverage"/>
///   </filterColumn>
///   <filterColumn colId="5">
///     <top10 top="1" percent="0" val="10"/>
///   </filterColumn>
/// </autoFilter>
/// ```
pub fn parse_auto_filter<R: BufRead>(e: &BytesStart, xml: &mut Reader<R>) -> Option<AutoFilter> {
    // Parse the ref attribute for the filter range
    let mut range = String::new();
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"ref" {
            range = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
        }
    }

    if range.is_empty() {
        return None;
    }

    // Parse the range to get start/end coordinates
    let (start_ref, end_ref) = range
        .split_once(':')
        .unwrap_or((range.as_str(), range.as_str()));
    let (start_col, start_row) = parse_cell_ref_or_default(start_ref);
    let (end_col, end_row) = parse_cell_ref_or_default(end_ref);

    let mut filter_columns = Vec::new();
    let mut buf = Vec::new();

    // Parse filterColumn children
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                if name == "filterColumn" {
                    if let Some(filter_col) = parse_filter_column(inner, xml) {
                        filter_columns.push(filter_col);
                    }
                }
            }
            Ok(Event::Empty(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                if name == "filterColumn" {
                    // Empty filterColumn (no filter criteria, just the column definition)
                    let mut col_id: u32 = 0;
                    let mut show_button: Option<bool> = None;

                    for attr in inner.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"colId" => {
                                col_id = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0);
                            }
                            b"showButton" => {
                                show_button =
                                    Some(std::str::from_utf8(&attr.value).unwrap_or("1") == "1");
                            }
                            b"hiddenButton" => {
                                // hiddenButton="1" means button is hidden (showButton=false)
                                if std::str::from_utf8(&attr.value).unwrap_or("0") == "1" {
                                    show_button = Some(false);
                                }
                            }
                            _ => {}
                        }
                    }

                    filter_columns.push(FilterColumn {
                        col_id,
                        has_filter: false,
                        filter_type: FilterType::None,
                        show_button,
                        values: Vec::new(),
                        custom_filters: Vec::new(),
                        custom_filters_and: None,
                        dxf_id: None,
                        cell_color: None,
                        icon_set: None,
                        icon_id: None,
                        dynamic_type: None,
                        top: None,
                        percent: None,
                        top10_val: None,
                    });
                }
            }
            Ok(Event::End(ref inner)) => {
                if inner.local_name().as_ref() == b"autoFilter" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Some(AutoFilter {
        range,
        start_row,
        start_col,
        end_row,
        end_col,
        filter_columns,
    })
}

/// Parse a filterColumn element and its children
fn parse_filter_column<R: BufRead>(e: &BytesStart, xml: &mut Reader<R>) -> Option<FilterColumn> {
    let mut col_id: u32 = 0;
    let mut show_button: Option<bool> = None;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"colId" => {
                col_id = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            b"showButton" => {
                show_button = Some(std::str::from_utf8(&attr.value).unwrap_or("1") == "1");
            }
            b"hiddenButton" => {
                // hiddenButton="1" means button is hidden (showButton=false)
                if std::str::from_utf8(&attr.value).unwrap_or("0") == "1" {
                    show_button = Some(false);
                }
            }
            _ => {}
        }
    }

    let mut filter_type = FilterType::None;
    let mut values: Vec<String> = Vec::new();
    let mut custom_filters: Vec<CustomFilter> = Vec::new();
    let mut custom_filters_and: Option<bool> = None;
    let mut dxf_id: Option<u32> = None;
    let mut cell_color: Option<bool> = None;
    let mut icon_set: Option<u32> = None;
    let mut icon_id: Option<u32> = None;
    let mut dynamic_type: Option<String> = None;
    let mut top: Option<bool> = None;
    let mut percent: Option<bool> = None;
    let mut top10_val: Option<f64> = None;
    let mut has_filter = false;

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "filters" => {
                        // Parse <filters> element containing <filter val="..."/> children
                        filter_type = FilterType::Values;
                        has_filter = true;
                        values = parse_filters_element(xml);
                    }
                    "customFilters" => {
                        // Parse <customFilters> element
                        filter_type = FilterType::Custom;
                        has_filter = true;

                        // Check for "and" attribute
                        for attr in inner.attributes().flatten() {
                            if attr.key.as_ref() == b"and" {
                                custom_filters_and =
                                    Some(std::str::from_utf8(&attr.value).unwrap_or("0") == "1");
                            }
                        }

                        custom_filters = parse_custom_filters_element(xml);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "colorFilter" => {
                        // <colorFilter dxfId="0" cellColor="1"/>
                        filter_type = FilterType::Color;
                        has_filter = true;

                        for attr in inner.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"dxfId" => {
                                    dxf_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"cellColor" => {
                                    cell_color = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or("1") == "1",
                                    );
                                }
                                _ => {}
                            }
                        }
                    }
                    "iconFilter" => {
                        // <iconFilter iconSet="3Arrows" iconId="0"/>
                        filter_type = FilterType::Icon;
                        has_filter = true;

                        for attr in inner.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"iconSet" => {
                                    // iconSet can be a string like "3Arrows" or a number
                                    // We'll store the numeric value if parseable
                                    icon_set = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"iconId" => {
                                    icon_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }
                    }
                    "dynamicFilter" => {
                        // <dynamicFilter type="aboveAverage"/>
                        filter_type = FilterType::Dynamic;
                        has_filter = true;

                        for attr in inner.attributes().flatten() {
                            if attr.key.as_ref() == b"type" {
                                dynamic_type = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                    }
                    "top10" => {
                        // <top10 top="1" percent="0" val="10"/>
                        filter_type = FilterType::Top10;
                        has_filter = true;

                        for attr in inner.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"top" => {
                                    top = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or("1") == "1",
                                    );
                                }
                                b"percent" => {
                                    percent = Some(
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
                                    );
                                }
                                b"val" => {
                                    top10_val = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }
                    }
                    "filters" => {
                        // Empty <filters/> element (blank filter - filter for blanks only)
                        filter_type = FilterType::Values;
                        has_filter = true;
                        // Check for blank attribute
                        for attr in inner.attributes().flatten() {
                            if attr.key.as_ref() == b"blank"
                                && std::str::from_utf8(&attr.value).unwrap_or("0") == "1"
                            {
                                // This filters for blank cells
                                values.push(String::new());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref inner)) => {
                if inner.local_name().as_ref() == b"filterColumn" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Some(FilterColumn {
        col_id,
        has_filter,
        filter_type,
        show_button,
        values,
        custom_filters,
        custom_filters_and,
        dxf_id,
        cell_color,
        icon_set,
        icon_id,
        dynamic_type,
        top,
        percent,
        top10_val,
    })
}

/// Parse <filters> element containing <filter val="..."/> children
fn parse_filters_element<R: BufRead>(xml: &mut Reader<R>) -> Vec<String> {
    let mut values = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "filter" {
                    // <filter val="Value1"/>
                    for attr in inner.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                values.push(val.to_string());
                            }
                        }
                    }
                } else if name == "dateGroupItem" {
                    // <dateGroupItem year="2024" month="1" dateTimeGrouping="month"/>
                    // For date filters, we can construct a string representation
                    let mut year: Option<String> = None;
                    let mut month: Option<String> = None;
                    let mut day: Option<String> = None;

                    for attr in inner.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"year" => {
                                year = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                            b"month" => {
                                month = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                            b"day" => {
                                day = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                            _ => {}
                        }
                    }

                    // Build a date string representation
                    let date_str = match (year, month, day) {
                        (Some(y), Some(m), Some(d)) => format!("{y}-{m}-{d}"),
                        (Some(y), Some(m), None) => format!("{y}-{m}"),
                        (Some(y), None, None) => y,
                        _ => String::new(),
                    };

                    if !date_str.is_empty() {
                        values.push(date_str);
                    }
                }
            }
            Ok(Event::Start(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "filter" {
                    // <filter val="Value1"/> might also appear as a start element
                    for attr in inner.attributes().flatten() {
                        if attr.key.as_ref() == b"val" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                values.push(val.to_string());
                            }
                        }
                    }
                }
            }
            Ok(Event::End(ref inner)) => {
                if inner.local_name().as_ref() == b"filters" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    values
}

/// Parse <customFilters> element containing <customFilter operator="..." val="..."/> children
fn parse_custom_filters_element<R: BufRead>(xml: &mut Reader<R>) -> Vec<CustomFilter> {
    let mut filters = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref inner) | Event::Start(ref inner)) => {
                let local_name = inner.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "customFilter" {
                    // <customFilter operator="greaterThan" val="100"/>
                    let mut operator = CustomFilterOperator::Equal;
                    let mut val = String::new();

                    for attr in inner.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"operator" => {
                                let op_str = std::str::from_utf8(&attr.value).unwrap_or("equal");
                                operator = match op_str {
                                    "equal" => CustomFilterOperator::Equal,
                                    "notEqual" => CustomFilterOperator::NotEqual,
                                    "greaterThan" => CustomFilterOperator::GreaterThan,
                                    "greaterThanOrEqual" => {
                                        CustomFilterOperator::GreaterThanOrEqual
                                    }
                                    "lessThan" => CustomFilterOperator::LessThan,
                                    "lessThanOrEqual" => CustomFilterOperator::LessThanOrEqual,
                                    _ => CustomFilterOperator::Equal,
                                };
                            }
                            b"val" => {
                                val = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            _ => {}
                        }
                    }

                    filters.push(CustomFilter { operator, val });
                }
            }
            Ok(Event::End(ref inner)) => {
                if inner.local_name().as_ref() == b"customFilters" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    filters
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
    use std::io::Cursor;

    #[test]
    fn test_parse_auto_filter_basic() {
        let xml_str = r#"<autoFilter ref="A1:D10"/>"#;
        let mut reader = Reader::from_reader(Cursor::new(xml_str));
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let result = parse_auto_filter(e, &mut reader);
            assert!(result.is_some());
            let af = result.unwrap();
            assert_eq!(af.range, "A1:D10");
            assert_eq!(af.start_col, 0);
            assert_eq!(af.start_row, 0);
            assert_eq!(af.end_col, 3);
            assert_eq!(af.end_row, 9);
        }
    }

    #[test]
    fn test_parse_auto_filter_with_values() {
        let xml_str = r#"<autoFilter ref="A1:D10">
            <filterColumn colId="0">
                <filters>
                    <filter val="Value1"/>
                    <filter val="Value2"/>
                </filters>
            </filterColumn>
        </autoFilter>"#;

        let mut reader = Reader::from_reader(Cursor::new(xml_str));
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Start(ref e)) = reader.read_event_into(&mut buf) {
            let result = parse_auto_filter(e, &mut reader);
            assert!(result.is_some());
            let af = result.unwrap();
            assert_eq!(af.filter_columns.len(), 1);
            assert_eq!(af.filter_columns[0].col_id, 0);
            assert_eq!(af.filter_columns[0].filter_type, FilterType::Values);
            assert_eq!(af.filter_columns[0].values, vec!["Value1", "Value2"]);
        }
    }

    #[test]
    fn test_parse_auto_filter_with_custom_filters() {
        let xml_str = r#"<autoFilter ref="A1:D10">
            <filterColumn colId="1">
                <customFilters and="1">
                    <customFilter operator="greaterThan" val="100"/>
                    <customFilter operator="lessThan" val="500"/>
                </customFilters>
            </filterColumn>
        </autoFilter>"#;

        let mut reader = Reader::from_reader(Cursor::new(xml_str));
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Start(ref e)) = reader.read_event_into(&mut buf) {
            let result = parse_auto_filter(e, &mut reader);
            assert!(result.is_some());
            let af = result.unwrap();
            assert_eq!(af.filter_columns.len(), 1);
            assert_eq!(af.filter_columns[0].col_id, 1);
            assert_eq!(af.filter_columns[0].filter_type, FilterType::Custom);
            assert_eq!(af.filter_columns[0].custom_filters_and, Some(true));
            assert_eq!(af.filter_columns[0].custom_filters.len(), 2);
            assert_eq!(
                af.filter_columns[0].custom_filters[0].operator,
                CustomFilterOperator::GreaterThan
            );
            assert_eq!(af.filter_columns[0].custom_filters[0].val, "100");
        }
    }

    #[test]
    fn test_parse_auto_filter_with_color_filter() {
        let xml_str = r#"<autoFilter ref="A1:D10">
            <filterColumn colId="2">
                <colorFilter dxfId="0" cellColor="1"/>
            </filterColumn>
        </autoFilter>"#;

        let mut reader = Reader::from_reader(Cursor::new(xml_str));
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Start(ref e)) = reader.read_event_into(&mut buf) {
            let result = parse_auto_filter(e, &mut reader);
            assert!(result.is_some());
            let af = result.unwrap();
            assert_eq!(af.filter_columns.len(), 1);
            assert_eq!(af.filter_columns[0].filter_type, FilterType::Color);
            assert_eq!(af.filter_columns[0].dxf_id, Some(0));
            assert_eq!(af.filter_columns[0].cell_color, Some(true));
        }
    }

    #[test]
    fn test_parse_auto_filter_with_top10() {
        let xml_str = r#"<autoFilter ref="A1:D10">
            <filterColumn colId="3">
                <top10 top="1" percent="0" val="10"/>
            </filterColumn>
        </autoFilter>"#;

        let mut reader = Reader::from_reader(Cursor::new(xml_str));
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Start(ref e)) = reader.read_event_into(&mut buf) {
            let result = parse_auto_filter(e, &mut reader);
            assert!(result.is_some());
            let af = result.unwrap();
            assert_eq!(af.filter_columns.len(), 1);
            assert_eq!(af.filter_columns[0].filter_type, FilterType::Top10);
            assert_eq!(af.filter_columns[0].top, Some(true));
            assert_eq!(af.filter_columns[0].percent, Some(false));
            assert_eq!(af.filter_columns[0].top10_val, Some(10.0));
        }
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref_or_default("A1"), (0, 0));
        assert_eq!(parse_cell_ref_or_default("B2"), (1, 1));
        assert_eq!(parse_cell_ref_or_default("Z26"), (25, 25));
        assert_eq!(parse_cell_ref_or_default("AA1"), (26, 0));
        assert_eq!(parse_cell_ref_or_default("AZ100"), (51, 99));
    }
}

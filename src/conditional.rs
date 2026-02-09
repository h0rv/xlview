//! Conditional formatting parsing module
//! This module handles parsing of conditional formatting rules from XLSX files.

use crate::color::resolve_color;
use crate::types::{
    CFRule, CFRuleType, CFValueObject, ColorScale, ConditionalFormatting, DataBar, IconSet,
};
use crate::xml_helpers::parse_color_attrs;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

/// Parse conditional formatting rules from a `<conditionalFormatting>` element
///
/// # Arguments
/// * `start_element` - The opening `<conditionalFormatting>` tag with attributes
/// * `xml` - The XML reader positioned after the start tag
/// * `theme_colors` - Theme colors for resolving color references
///
/// # Returns
/// The parsed conditional formatting, or None if parsing fails
///
/// # XML Format
/// ```xml
/// <conditionalFormatting sqref="A1:A10">
///   <cfRule type="colorScale" priority="1">
///     <colorScale>
///       <cfvo type="min"/>
///       <cfvo type="max"/>
///       <color rgb="FFF8696B"/>
///       <color rgb="FF63BE7B"/>
///     </colorScale>
///   </cfRule>
///   <cfRule type="dataBar" priority="2">
///     <dataBar>
///       <cfvo type="min"/>
///       <cfvo type="max"/>
///       <color rgb="FF638EC6"/>
///     </dataBar>
///   </cfRule>
///   <cfRule type="iconSet" priority="3">
///     <iconSet iconSet="3Arrows">
///       <cfvo type="percent" val="33"/>
///       <cfvo type="percent" val="67"/>
///     </iconSet>
///   </cfRule>
/// </conditionalFormatting>
/// ```
pub fn parse_conditional_formatting<R: BufRead>(
    start_element: &BytesStart,
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<ConditionalFormatting> {
    // Parse sqref attribute from the conditionalFormatting element
    let mut sqref = String::new();
    for attr in start_element.attributes().flatten() {
        if attr.key.as_ref() == b"sqref" {
            sqref = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
        }
    }

    if sqref.is_empty() {
        return None;
    }

    let mut rules = Vec::new();
    let mut buf = Vec::new();

    // Parse child elements (cfRule elements)
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "cfRule" {
                    if let Some(rule) = parse_cf_rule(e, xml, theme_colors, indexed_colors) {
                        rules.push(rule);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                // Handle self-closing cfRule elements (e.g., <cfRule type="top10" .../>)
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "cfRule" {
                    if let Some(rule) = parse_cf_rule_empty(e) {
                        rules.push(rule);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"conditionalFormatting" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if rules.is_empty() {
        return None;
    }

    Some(ConditionalFormatting { sqref, rules })
}

/// Parse a self-closing cfRule element (no child elements like colorScale, dataBar, etc.)
/// Used for rule types like top10, aboveAverage, timePeriod, duplicateValues, uniqueValues, containsBlanks, notContainsBlanks
fn parse_cf_rule_empty(element: &BytesStart) -> Option<CFRule> {
    let mut rule_type_str = String::new();
    let mut priority: u32 = 0;
    let mut operator: Option<String> = None;
    let mut dxf_id: Option<u32> = None;

    // Top10 rule attributes
    let mut rank: Option<u32> = None;
    let mut percent: Option<bool> = None;
    let mut bottom: Option<bool> = None;

    // AboveAverage rule attributes
    let mut above_average: Option<bool> = None;
    let mut equal_average: Option<bool> = None;
    let mut std_dev: Option<u32> = None;

    // TimePeriod rule attributes
    let mut time_period: Option<String> = None;

    // Parse attributes from cfRule
    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"type" => {
                rule_type_str = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"priority" => {
                priority = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            b"operator" => {
                operator = std::str::from_utf8(&attr.value)
                    .ok()
                    .map(ToString::to_string);
            }
            b"dxfId" => {
                dxf_id = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            // Top10 attributes
            b"rank" => {
                rank = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"percent" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                percent = Some(val == "1" || val == "true");
            }
            b"bottom" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                bottom = Some(val == "1" || val == "true");
            }
            // AboveAverage attributes
            b"aboveAverage" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("1");
                // Note: aboveAverage defaults to true when not present, but when present "0" means below average
                above_average = Some(val != "0" && val != "false");
            }
            b"equalAverage" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                equal_average = Some(val == "1" || val == "true");
            }
            b"stdDev" => {
                std_dev = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            // TimePeriod attributes
            b"timePeriod" => {
                time_period = std::str::from_utf8(&attr.value)
                    .ok()
                    .map(ToString::to_string);
            }
            _ => {}
        }
    }

    Some(CFRule {
        rule_type: CFRuleType::from_str_val(&rule_type_str),
        priority,
        color_scale: None,
        data_bar: None,
        icon_set: None,
        formula: None,
        operator,
        dxf_id,
        rank,
        percent,
        bottom,
        above_average,
        equal_average,
        std_dev,
        time_period,
    })
}

/// Parse a single cfRule element
fn parse_cf_rule<R: BufRead>(
    start_element: &BytesStart,
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<CFRule> {
    let mut rule_type_str = String::new();
    let mut priority: u32 = 0;
    let mut operator: Option<String> = None;
    let mut dxf_id: Option<u32> = None;

    // Top10 rule attributes
    let mut rank: Option<u32> = None;
    let mut percent: Option<bool> = None;
    let mut bottom: Option<bool> = None;

    // AboveAverage rule attributes
    let mut above_average: Option<bool> = None;
    let mut equal_average: Option<bool> = None;
    let mut std_dev: Option<u32> = None;

    // TimePeriod rule attributes
    let mut time_period: Option<String> = None;

    // Parse attributes from cfRule
    for attr in start_element.attributes().flatten() {
        match attr.key.as_ref() {
            b"type" => {
                rule_type_str = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"priority" => {
                priority = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            b"operator" => {
                operator = std::str::from_utf8(&attr.value)
                    .ok()
                    .map(ToString::to_string);
            }
            b"dxfId" => {
                dxf_id = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            // Top10 attributes
            b"rank" => {
                rank = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"percent" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                percent = Some(val == "1" || val == "true");
            }
            b"bottom" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                bottom = Some(val == "1" || val == "true");
            }
            // AboveAverage attributes
            b"aboveAverage" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("1");
                // Note: aboveAverage defaults to true when not present, but when present "0" means below average
                above_average = Some(val != "0" && val != "false");
            }
            b"equalAverage" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                equal_average = Some(val == "1" || val == "true");
            }
            b"stdDev" => {
                std_dev = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            // TimePeriod attributes
            b"timePeriod" => {
                time_period = std::str::from_utf8(&attr.value)
                    .ok()
                    .map(ToString::to_string);
            }
            _ => {}
        }
    }

    let mut color_scale: Option<ColorScale> = None;
    let mut data_bar: Option<DataBar> = None;
    let mut icon_set: Option<IconSet> = None;
    let mut formula: Option<String> = None;

    let mut buf = Vec::new();

    // Parse child elements
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "colorScale" => {
                        color_scale = parse_color_scale(xml, theme_colors, indexed_colors);
                    }
                    "dataBar" => {
                        data_bar = parse_data_bar(e, xml, theme_colors, indexed_colors);
                    }
                    "iconSet" => {
                        icon_set = parse_icon_set(e, xml);
                    }
                    "formula" => {
                        // Read the formula text content
                        let mut text_buf = Vec::new();
                        if let Ok(Event::Text(text)) = xml.read_event_into(&mut text_buf) {
                            formula = text.unescape().ok().map(|s| s.to_string());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"cfRule" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Some(CFRule {
        rule_type: CFRuleType::from_str_val(&rule_type_str),
        priority,
        color_scale,
        data_bar,
        icon_set,
        formula,
        operator,
        dxf_id,
        rank,
        percent,
        bottom,
        above_average,
        equal_average,
        std_dev,
        time_period,
    })
}

/// Parse a colorScale element
/// ```xml
/// <colorScale>
///   <cfvo type="min"/>
///   <cfvo type="percentile" val="50"/>
///   <cfvo type="max"/>
///   <color rgb="FFF8696B"/>
///   <color rgb="FFFCFCFF"/>
///   <color rgb="FF63BE7B"/>
/// </colorScale>
/// ```
fn parse_color_scale<R: BufRead>(
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<ColorScale> {
    let mut cfvo = Vec::new();
    let mut colors = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "cfvo" => {
                        if let Some(vo) = parse_cfvo(e) {
                            cfvo.push(vo);
                        }
                    }
                    "color" => {
                        if let Some(color) = parse_color_element(e, theme_colors, indexed_colors) {
                            colors.push(color);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"colorScale" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if cfvo.is_empty() || colors.is_empty() {
        return None;
    }

    Some(ColorScale { cfvo, colors })
}

/// Parse a dataBar element
/// ```xml
/// <dataBar minLength="10" maxLength="90" showValue="1">
///   <cfvo type="min"/>
///   <cfvo type="max"/>
///   <color rgb="FF638EC6"/>
/// </dataBar>
/// ```
fn parse_data_bar<R: BufRead>(
    start_element: &BytesStart,
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<DataBar> {
    let mut show_value: Option<bool> = None;
    let mut min_length: Option<u32> = None;
    let mut max_length: Option<u32> = None;

    // Parse attributes from dataBar element
    for attr in start_element.attributes().flatten() {
        match attr.key.as_ref() {
            b"showValue" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("1");
                show_value = Some(val != "0");
            }
            b"minLength" => {
                min_length = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"maxLength" => {
                max_length = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            _ => {}
        }
    }

    let mut cfvo = Vec::new();
    let mut color = String::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "cfvo" => {
                        if let Some(vo) = parse_cfvo(e) {
                            cfvo.push(vo);
                        }
                    }
                    "color" => {
                        if let Some(c) = parse_color_element(e, theme_colors, indexed_colors) {
                            color = c;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"dataBar" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if cfvo.is_empty() || color.is_empty() {
        return None;
    }

    Some(DataBar {
        cfvo,
        color,
        show_value,
        min_length,
        max_length,
    })
}

/// Parse an iconSet element
/// ```xml
/// <iconSet iconSet="3Arrows" showValue="0" reverse="0">
///   <cfvo type="percent" val="0"/>
///   <cfvo type="percent" val="33"/>
///   <cfvo type="percent" val="67"/>
/// </iconSet>
/// ```
fn parse_icon_set<R: BufRead>(start_element: &BytesStart, xml: &mut Reader<R>) -> Option<IconSet> {
    let mut icon_set_name = String::from("3TrafficLights1"); // Default icon set
    let mut show_value: Option<bool> = None;
    let mut reverse: Option<bool> = None;

    // Parse attributes from iconSet element
    for attr in start_element.attributes().flatten() {
        match attr.key.as_ref() {
            b"iconSet" => {
                icon_set_name = std::str::from_utf8(&attr.value)
                    .unwrap_or("3TrafficLights1")
                    .to_string();
            }
            b"showValue" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("1");
                show_value = Some(val != "0");
            }
            b"reverse" => {
                let val = std::str::from_utf8(&attr.value).unwrap_or("0");
                reverse = Some(val == "1");
            }
            _ => {}
        }
    }

    let mut cfvo = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "cfvo" {
                    if let Some(vo) = parse_cfvo(e) {
                        cfvo.push(vo);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"iconSet" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if cfvo.is_empty() {
        return None;
    }

    Some(IconSet {
        icon_set: icon_set_name,
        cfvo,
        show_value,
        reverse,
    })
}

/// Parse a cfvo (conditional formatting value object) element
/// ```xml
/// <cfvo type="min"/>
/// <cfvo type="num" val="100"/>
/// <cfvo type="percent" val="50"/>
/// <cfvo type="percentile" val="90"/>
/// <cfvo type="formula" val="$A$1"/>
/// ```
fn parse_cfvo(element: &BytesStart) -> Option<CFValueObject> {
    let mut cfvo_type = String::new();
    let mut val: Option<String> = None;

    for attr in element.attributes().flatten() {
        match attr.key.as_ref() {
            b"type" => {
                cfvo_type = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"val" => {
                val = std::str::from_utf8(&attr.value)
                    .ok()
                    .map(ToString::to_string);
            }
            _ => {}
        }
    }

    if cfvo_type.is_empty() {
        return None;
    }

    Some(CFValueObject { cfvo_type, val })
}

/// Parse a color element and resolve it to an #RRGGBB string
/// ```xml
/// <color rgb="FFF8696B"/>
/// <color theme="4" tint="0.59999389629810485"/>
/// <color indexed="5"/>
/// ```
fn parse_color_element(
    element: &BytesStart,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<String> {
    let color_spec = parse_color_attrs(element);
    resolve_color(&color_spec, theme_colors, indexed_colors)
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

    fn default_theme_colors() -> Vec<String> {
        vec![
            "#000000".to_string(),
            "#FFFFFF".to_string(),
            "#44546A".to_string(),
            "#E7E6E6".to_string(),
            "#4472C4".to_string(),
            "#ED7D31".to_string(),
            "#A5A5A5".to_string(),
            "#FFC000".to_string(),
            "#5B9BD5".to_string(),
            "#70AD47".to_string(),
            "#0563C1".to_string(),
            "#954F72".to_string(),
        ]
    }

    #[test]
    fn test_parse_color_scale() {
        let xml_content = r#"<conditionalFormatting sqref="A1:A10">
            <cfRule type="colorScale" priority="1">
                <colorScale>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFF8696B"/>
                    <color rgb="FF63BE7B"/>
                </colorScale>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        // Read until we find the conditionalFormatting start tag
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();
                        assert_eq!(cf.sqref, "A1:A10");
                        assert_eq!(cf.rules.len(), 1);

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "colorScale");
                        assert_eq!(rule.priority, 1);
                        assert!(rule.color_scale.is_some());

                        let cs = rule.color_scale.as_ref().unwrap();
                        assert_eq!(cs.cfvo.len(), 2);
                        assert_eq!(cs.cfvo[0].cfvo_type, "min");
                        assert_eq!(cs.cfvo[1].cfvo_type, "max");
                        assert_eq!(cs.colors.len(), 2);
                        assert_eq!(cs.colors[0], "#F8696B");
                        assert_eq!(cs.colors[1], "#63BE7B");
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_data_bar() {
        let xml_content = r#"<conditionalFormatting sqref="B1:B20">
            <cfRule type="dataBar" priority="2">
                <dataBar minLength="10" maxLength="90">
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF638EC6"/>
                </dataBar>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();
                        assert_eq!(cf.sqref, "B1:B20");

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "dataBar");
                        assert!(rule.data_bar.is_some());

                        let db = rule.data_bar.as_ref().unwrap();
                        assert_eq!(db.cfvo.len(), 2);
                        assert_eq!(db.color, "#638EC6");
                        assert_eq!(db.min_length, Some(10));
                        assert_eq!(db.max_length, Some(90));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_icon_set() {
        let xml_content = r#"<conditionalFormatting sqref="C1:C10">
            <cfRule type="iconSet" priority="3">
                <iconSet iconSet="3Arrows" reverse="1">
                    <cfvo type="percent" val="0"/>
                    <cfvo type="percent" val="33"/>
                    <cfvo type="percent" val="67"/>
                </iconSet>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "iconSet");
                        assert!(rule.icon_set.is_some());

                        let is = rule.icon_set.as_ref().unwrap();
                        assert_eq!(is.icon_set, "3Arrows");
                        assert_eq!(is.reverse, Some(true));
                        assert_eq!(is.cfvo.len(), 3);
                        assert_eq!(is.cfvo[0].cfvo_type, "percent");
                        assert_eq!(is.cfvo[0].val, Some("0".to_string()));
                        assert_eq!(is.cfvo[1].val, Some("33".to_string()));
                        assert_eq!(is.cfvo[2].val, Some("67".to_string()));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_cell_is_rule() {
        let xml_content = r#"<conditionalFormatting sqref="D1:D10">
            <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
                <formula>100</formula>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "cellIs");
                        assert_eq!(rule.operator, Some("greaterThan".to_string()));
                        assert_eq!(rule.dxf_id, Some(0));
                        assert_eq!(rule.formula, Some("100".to_string()));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_multiple_rules() {
        let xml_content = r#"<conditionalFormatting sqref="E1:E10">
            <cfRule type="colorScale" priority="1">
                <colorScale>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF0000"/>
                    <color rgb="FF00FF00"/>
                </colorScale>
            </cfRule>
            <cfRule type="dataBar" priority="2">
                <dataBar>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FF0000FF"/>
                </dataBar>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();
                        assert_eq!(cf.rules.len(), 2);
                        assert_eq!(cf.rules[0].rule_type, "colorScale");
                        assert_eq!(cf.rules[1].rule_type, "dataBar");
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_three_color_scale() {
        let xml_content = r#"<conditionalFormatting sqref="F1:F10">
            <cfRule type="colorScale" priority="1">
                <colorScale>
                    <cfvo type="min"/>
                    <cfvo type="percentile" val="50"/>
                    <cfvo type="max"/>
                    <color rgb="FFF8696B"/>
                    <color rgb="FFFCFCFF"/>
                    <color rgb="FF63BE7B"/>
                </colorScale>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        let cs = rule.color_scale.as_ref().unwrap();
                        assert_eq!(cs.cfvo.len(), 3);
                        assert_eq!(cs.colors.len(), 3);
                        assert_eq!(cs.cfvo[1].cfvo_type, "percentile");
                        assert_eq!(cs.cfvo[1].val, Some("50".to_string()));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_theme_color() {
        let xml_content = r#"<conditionalFormatting sqref="G1:G10">
            <cfRule type="colorScale" priority="1">
                <colorScale>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color theme="4"/>
                    <color theme="9"/>
                </colorScale>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        let cs = rule.color_scale.as_ref().unwrap();
                        // theme 4 = accent1 = #4472C4
                        // theme 9 = accent6 = #70AD47
                        assert_eq!(cs.colors[0], "#4472C4");
                        assert_eq!(cs.colors[1], "#70AD47");
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_empty_sqref_returns_none() {
        let xml_content = r#"<conditionalFormatting sqref="">
            <cfRule type="colorScale" priority="1">
                <colorScale>
                    <cfvo type="min"/>
                    <cfvo type="max"/>
                    <color rgb="FFFF0000"/>
                    <color rgb="FF00FF00"/>
                </colorScale>
            </cfRule>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_none());
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_top10_rule() {
        let xml_content = r#"<conditionalFormatting sqref="A1:A100">
            <cfRule type="top10" dxfId="0" priority="1" rank="10" percent="0" bottom="0"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "top10");
                        assert_eq!(rule.rank, Some(10));
                        assert_eq!(rule.percent, Some(false));
                        assert_eq!(rule.bottom, Some(false));
                        assert_eq!(rule.dxf_id, Some(0));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_top10_bottom_percent_rule() {
        let xml_content = r#"<conditionalFormatting sqref="B1:B100">
            <cfRule type="top10" dxfId="1" priority="2" rank="25" percent="1" bottom="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "top10");
                        assert_eq!(rule.rank, Some(25));
                        assert_eq!(rule.percent, Some(true));
                        assert_eq!(rule.bottom, Some(true));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_above_average_rule() {
        let xml_content = r#"<conditionalFormatting sqref="C1:C100">
            <cfRule type="aboveAverage" dxfId="2" priority="1" aboveAverage="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "aboveAverage");
                        assert_eq!(rule.above_average, Some(true));
                        assert_eq!(rule.dxf_id, Some(2));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_below_average_rule() {
        let xml_content = r#"<conditionalFormatting sqref="D1:D100">
            <cfRule type="aboveAverage" dxfId="3" priority="1" aboveAverage="0" equalAverage="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "aboveAverage");
                        assert_eq!(rule.above_average, Some(false));
                        assert_eq!(rule.equal_average, Some(true));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_std_dev_rule() {
        let xml_content = r#"<conditionalFormatting sqref="E1:E100">
            <cfRule type="aboveAverage" dxfId="4" priority="1" aboveAverage="1" stdDev="2"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "aboveAverage");
                        assert_eq!(rule.above_average, Some(true));
                        assert_eq!(rule.std_dev, Some(2));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_time_period_rule() {
        let xml_content = r#"<conditionalFormatting sqref="F1:F100">
            <cfRule type="timePeriod" dxfId="5" priority="1" timePeriod="today"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "timePeriod");
                        assert_eq!(rule.time_period, Some("today".to_string()));
                        assert_eq!(rule.dxf_id, Some(5));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_time_period_last_week() {
        let xml_content = r#"<conditionalFormatting sqref="G1:G100">
            <cfRule type="timePeriod" dxfId="6" priority="1" timePeriod="lastWeek"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "timePeriod");
                        assert_eq!(rule.time_period, Some("lastWeek".to_string()));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_duplicate_values_rule() {
        let xml_content = r#"<conditionalFormatting sqref="H1:H100">
            <cfRule type="duplicateValues" dxfId="7" priority="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "duplicateValues");
                        assert_eq!(rule.dxf_id, Some(7));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_unique_values_rule() {
        let xml_content = r#"<conditionalFormatting sqref="I1:I100">
            <cfRule type="uniqueValues" dxfId="8" priority="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "uniqueValues");
                        assert_eq!(rule.dxf_id, Some(8));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_contains_blanks_rule() {
        let xml_content = r#"<conditionalFormatting sqref="J1:J100">
            <cfRule type="containsBlanks" dxfId="9" priority="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "containsBlanks");
                        assert_eq!(rule.dxf_id, Some(9));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }

    #[test]
    fn test_parse_not_contains_blanks_rule() {
        let xml_content = r#"<conditionalFormatting sqref="K1:K100">
            <cfRule type="notContainsBlanks" dxfId="10" priority="1"/>
        </conditionalFormatting>"#;

        let cursor = Cursor::new(xml_content);
        let mut xml = Reader::from_reader(cursor);
        xml.trim_text(true);

        let mut buf = Vec::new();
        let theme_colors = default_theme_colors();

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"conditionalFormatting" {
                        let result = parse_conditional_formatting(e, &mut xml, &theme_colors, None);
                        assert!(result.is_some());
                        let cf = result.unwrap();

                        let rule = &cf.rules[0];
                        assert_eq!(rule.rule_type, "notContainsBlanks");
                        assert_eq!(rule.dxf_id, Some(10));
                        return;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        panic!("Did not find conditionalFormatting element");
    }
}

//! Rich text parsing module
//! This module handles parsing of rich text runs with per-character formatting.

use crate::color::resolve_color;
use crate::types::{RichTextRun, RunStyle, SharedString, VerticalAlign};
use crate::xml_helpers::parse_color_attrs;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::BufRead;

/// Parse a single rich text run (`<r>` element)
///
/// Structure:
/// ```xml
/// <r>
///   <rPr>
///     <b/><i/><u/><strike/>
///     <sz val="12"/>
///     <color rgb="FFFF0000"/>
///     <rFont val="Arial"/>
///     <vertAlign val="superscript"/>
///   </rPr>
///   <t>Text content</t>
/// </r>
/// ```
pub fn parse_rich_text_run<R: BufRead>(
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<RichTextRun> {
    let mut buf = Vec::new();
    let mut text = String::new();
    let mut style: Option<RunStyle> = None;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"rPr" => {
                        style = Some(parse_run_properties(xml, theme_colors, indexed_colors));
                    }
                    b"t" => {
                        // Text element - read the content
                        text = read_text_content(xml);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                if local_name.as_ref() == b"t" {
                    // Empty <t/> element - empty text
                    text = String::new();
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Some(RichTextRun { text, style })
}

/// Parse run properties (`<rPr>` element)
///
/// Parses all formatting: bold, italic, underline, strikethrough,
/// font_family, font_size, font_color, vert_align
pub fn parse_run_properties<R: BufRead>(
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> RunStyle {
    let mut buf = Vec::new();
    let mut style = RunStyle::default();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"color" => {
                        // Parse color with potential nested elements
                        let color_spec = parse_color_attrs(e);
                        if let Some(resolved) =
                            resolve_color(&color_spec, theme_colors, indexed_colors)
                        {
                            style.font_color = Some(resolved);
                        }
                    }
                    b"sz" => {
                        // Font size
                        if let Some(val) = get_attribute(e, b"val") {
                            if let Ok(size) = val.parse::<f64>() {
                                style.font_size = Some(size);
                            }
                        }
                    }
                    b"rFont" => {
                        // Font family
                        if let Some(val) = get_attribute(e, b"val") {
                            style.font_family = Some(val);
                        }
                    }
                    b"vertAlign" => {
                        // Vertical alignment (superscript/subscript)
                        if let Some(val) = get_attribute(e, b"val") {
                            style.vert_align = match val.as_str() {
                                "superscript" => Some(VerticalAlign::Superscript),
                                "subscript" => Some(VerticalAlign::Subscript),
                                "baseline" => Some(VerticalAlign::Baseline),
                                _ => None,
                            };
                        }
                    }
                    b"b" => {
                        // Bold - presence means true, val="0" means false
                        style.bold = Some(check_boolean_element(e));
                    }
                    b"i" => {
                        // Italic
                        style.italic = Some(check_boolean_element(e));
                    }
                    b"u" => {
                        // Underline
                        style.underline = Some(check_boolean_element(e));
                    }
                    b"strike" => {
                        // Strikethrough
                        style.strikethrough = Some(check_boolean_element(e));
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"color" => {
                        let color_spec = parse_color_attrs(e);
                        if let Some(resolved) =
                            resolve_color(&color_spec, theme_colors, indexed_colors)
                        {
                            style.font_color = Some(resolved);
                        }
                    }
                    b"sz" => {
                        if let Some(val) = get_attribute(e, b"val") {
                            if let Ok(size) = val.parse::<f64>() {
                                style.font_size = Some(size);
                            }
                        }
                    }
                    b"rFont" => {
                        if let Some(val) = get_attribute(e, b"val") {
                            style.font_family = Some(val);
                        }
                    }
                    b"vertAlign" => {
                        if let Some(val) = get_attribute(e, b"val") {
                            style.vert_align = match val.as_str() {
                                "superscript" => Some(VerticalAlign::Superscript),
                                "subscript" => Some(VerticalAlign::Subscript),
                                "baseline" => Some(VerticalAlign::Baseline),
                                _ => None,
                            };
                        }
                    }
                    b"b" => {
                        style.bold = Some(check_boolean_element(e));
                    }
                    b"i" => {
                        style.italic = Some(check_boolean_element(e));
                    }
                    b"u" => {
                        style.underline = Some(check_boolean_element(e));
                    }
                    b"strike" => {
                        style.strikethrough = Some(check_boolean_element(e));
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"rPr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    style
}

/// Parse shared string item - can be plain text or rich text
///
/// Structure:
/// ```xml
/// <si>
///   <t>Plain text</t>  -- OR --
///   <r><rPr>...</rPr><t>Styled</t></r>
///   <r><t>Normal</t></r>
/// </si>
/// ```
pub fn parse_shared_string_item<R: BufRead>(
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> SharedString {
    let mut buf = Vec::new();
    let mut plain_text: Option<String> = None;
    let mut runs: Vec<RichTextRun> = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"t" => {
                        // Plain text element (not inside <r>)
                        plain_text = Some(read_text_content(xml));
                    }
                    b"r" => {
                        // Rich text run
                        if let Some(run) = parse_rich_text_run(xml, theme_colors, indexed_colors) {
                            runs.push(run);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                if local_name.as_ref() == b"t" {
                    // Empty <t/> element
                    plain_text = Some(String::new());
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"si" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // If we have rich text runs, return those; otherwise return plain text
    if !runs.is_empty() {
        SharedString::Rich(runs)
    } else {
        SharedString::Plain(plain_text.unwrap_or_default())
    }
}

/// Extract plain text from rich text runs
pub fn rich_text_to_plain(runs: &[RichTextRun]) -> String {
    runs.iter().map(|r| r.text.as_str()).collect()
}

// =============================================================================
// Helper functions
// =============================================================================

/// Read text content from inside a <t> element until </t>
fn read_text_content<R: BufRead>(xml: &mut Reader<R>) -> String {
    let mut buf = Vec::new();
    let mut text = String::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Text(ref e)) => {
                if let Ok(t) = e.unescape() {
                    text.push_str(&t);
                }
            }
            Ok(Event::CData(ref e)) => {
                if let Ok(t) = std::str::from_utf8(e.as_ref()) {
                    text.push_str(t);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"t" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    text
}

/// Get an attribute value from an element
fn get_attribute(element: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in element.attributes().flatten() {
        if attr.key.local_name().as_ref() == name {
            return attr.unescape_value().ok().map(|s| s.into_owned());
        }
    }
    None
}

/// Check if a boolean element is true or false
/// In Excel XML, presence of the element means true, unless val="0" or val="false"
fn check_boolean_element(element: &quick_xml::events::BytesStart<'_>) -> bool {
    if let Some(val) = get_attribute(element, b"val") {
        !matches!(val.as_str(), "0" | "false")
    } else {
        // No val attribute means the element is present = true
        true
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
    fn test_parse_plain_shared_string() {
        let xml_str = r#"<si><t>Hello World</t></si>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        // Skip to the <si> element
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let result = parse_shared_string_item(&mut reader, &default_theme_colors(), None);

        match result {
            SharedString::Plain(text) => assert_eq!(text, "Hello World"),
            SharedString::Rich(_) => panic!("Expected plain text"),
        }
    }

    #[test]
    fn test_parse_rich_text_shared_string() {
        let xml_str =
            r#"<si><r><rPr><b/><sz val="14"/></rPr><t>Bold</t></r><r><t> Normal</t></r></si>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        // Don't use trim_text(true) here - whitespace in <t> elements is significant in Excel
        reader.trim_text(false);

        // Skip to the <si> element
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let result = parse_shared_string_item(&mut reader, &default_theme_colors(), None);

        match result {
            SharedString::Rich(runs) => {
                assert_eq!(runs.len(), 2);
                assert_eq!(runs[0].text, "Bold");
                assert!(runs[0].style.is_some());
                let style = runs[0].style.as_ref().unwrap();
                assert_eq!(style.bold, Some(true));
                assert_eq!(style.font_size, Some(14.0));

                assert_eq!(runs[1].text, " Normal");
            }
            SharedString::Plain(_) => panic!("Expected rich text"),
        }
    }

    #[test]
    fn test_rich_text_to_plain() {
        let runs = vec![
            RichTextRun {
                text: "Hello ".to_string(),
                style: None,
            },
            RichTextRun {
                text: "World".to_string(),
                style: Some(RunStyle {
                    bold: Some(true),
                    ..Default::default()
                }),
            },
        ];

        let plain = rich_text_to_plain(&runs);
        assert_eq!(plain, "Hello World");
    }

    #[test]
    fn test_parse_run_properties_with_color() {
        let xml_str =
            r#"<rPr><b/><i/><color rgb="FFFF0000"/><sz val="12"/><rFont val="Arial"/></rPr>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        // Skip to the <rPr> element
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let style = parse_run_properties(&mut reader, &default_theme_colors(), None);

        assert_eq!(style.bold, Some(true));
        assert_eq!(style.italic, Some(true));
        assert_eq!(style.font_color, Some("#FF0000".to_string()));
        assert_eq!(style.font_size, Some(12.0));
        assert_eq!(style.font_family, Some("Arial".to_string()));
    }

    #[test]
    fn test_parse_run_properties_with_theme_color() {
        let xml_str = r#"<rPr><color theme="4"/></rPr>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let style = parse_run_properties(&mut reader, &default_theme_colors(), None);

        // Theme 4 is accent1 = #4472C4
        assert_eq!(style.font_color, Some("#4472C4".to_string()));
    }

    #[test]
    fn test_parse_run_properties_with_vert_align() {
        let xml_str = r#"<rPr><vertAlign val="superscript"/></rPr>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let style = parse_run_properties(&mut reader, &default_theme_colors(), None);

        assert!(matches!(style.vert_align, Some(VerticalAlign::Superscript)));
    }

    #[test]
    fn test_parse_boolean_with_val_false() {
        let xml_str = r#"<rPr><b val="0"/></rPr>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let style = parse_run_properties(&mut reader, &default_theme_colors(), None);

        assert_eq!(style.bold, Some(false));
    }

    #[test]
    fn test_parse_empty_shared_string() {
        let xml_str = r#"<si><t/></si>"#;
        let cursor = Cursor::new(xml_str);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => break,
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"si" => {
                    // Handle empty <si/> case
                    break;
                }
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e:?}"),
                _ => {}
            }
            buf.clear();
        }

        let result = parse_shared_string_item(&mut reader, &default_theme_colors(), None);

        match result {
            SharedString::Plain(text) => assert_eq!(text, ""),
            SharedString::Rich(_) => panic!("Expected plain text"),
        }
    }
}

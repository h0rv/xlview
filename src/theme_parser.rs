//! Theme parsing module
//! This module handles parsing of theme colors and fonts from theme1.xml.

use crate::color::DEFAULT_THEME_COLORS;
use crate::types::Theme;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{BufRead, BufReader, Read, Seek};
use zip::ZipArchive;

/// Parse theme from theme1.xml
pub fn parse_theme<R: Read + Seek>(archive: &mut ZipArchive<R>, path: Option<&str>) -> Theme {
    let theme_path = path.unwrap_or("xl/theme/theme1.xml");

    let Ok(file) = archive.by_name(theme_path) else {
        // No theme file, return defaults
        return Theme {
            colors: DEFAULT_THEME_COLORS
                .iter()
                .map(|s| (*s).to_string())
                .collect(),
            major_font: None,
            minor_font: None,
        };
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut colors = Vec::new();
    let mut major_font = None;
    let mut minor_font = None;
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"clrScheme" => {
                        colors = parse_color_scheme(&mut xml);
                    }
                    b"fontScheme" => {
                        let (major, minor) = parse_font_scheme(&mut xml);
                        major_font = major;
                        minor_font = minor;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Fall back to defaults if no colors were parsed
    if colors.is_empty() {
        colors = DEFAULT_THEME_COLORS
            .iter()
            .map(|s| (*s).to_string())
            .collect();
    }

    Theme {
        colors,
        major_font,
        minor_font,
    }
}

/// Parse theme colors from clrScheme
/// <a:clrScheme name="Office">
///   <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
///   <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
///   <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
///   ...
/// </a:clrScheme>
fn parse_color_scheme<R: BufRead>(xml: &mut Reader<R>) -> Vec<String> {
    // Excel theme color indices (per ECMA-376):
    // 0: lt1 (Background 1 / light1) - typically white
    // 1: dk1 (Text 1 / dark1) - typically black
    // 2: lt2 (Background 2 / light2)
    // 3: dk2 (Text 2 / dark2)
    // 4-9: accent1-accent6
    // 10: hlink (hyperlink)
    // 11: folHlink (followed hyperlink)
    let color_order = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    let mut color_map: std::collections::HashMap<String, String> = std::collections::HashMap::new();
    let mut buf = Vec::new();
    let mut current_color_name: Option<String> = None;
    let mut depth = 1; // We've already entered clrScheme

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                let local_name = e.local_name();
                let name_str = String::from_utf8_lossy(local_name.as_ref()).to_string();

                // Check if this is a color element (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
                if color_order.contains(&name_str.as_str()) {
                    current_color_name = Some(name_str);
                } else if current_color_name.is_some() {
                    // Look for sysClr or srgbClr
                    match local_name.as_ref() {
                        b"sysClr" => {
                            // System color - use lastClr attribute for the actual color
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"lastClr" {
                                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                                    if let Some(ref name) = current_color_name {
                                        color_map.insert(name.clone(), format!("#{val}"));
                                    }
                                }
                            }
                        }
                        b"srgbClr" => {
                            // sRGB color - use val attribute
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                                    if let Some(ref name) = current_color_name {
                                        color_map.insert(name.clone(), format!("#{val}"));
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name_str = String::from_utf8_lossy(local_name.as_ref()).to_string();

                // Handle empty color elements (dk1, lt1, etc. that only contain sysClr/srgbClr)
                if color_order.contains(&name_str.as_str()) {
                    // This shouldn't happen - color elements always have children
                } else if current_color_name.is_some() {
                    // Look for sysClr or srgbClr
                    match local_name.as_ref() {
                        b"sysClr" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"lastClr" {
                                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                                    if let Some(ref name) = current_color_name {
                                        color_map.insert(name.clone(), format!("#{val}"));
                                    }
                                }
                            }
                        }
                        b"srgbClr" => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = String::from_utf8_lossy(&attr.value).to_uppercase();
                                    if let Some(ref name) = current_color_name {
                                        color_map.insert(name.clone(), format!("#{val}"));
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth -= 1;
                let local_name = e.local_name();
                let name_str = String::from_utf8_lossy(local_name.as_ref()).to_string();

                // Clear current color name when exiting a color element
                if color_order.contains(&name_str.as_str()) {
                    current_color_name = None;
                }

                // Exit when we close clrScheme
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Build the color vector in the correct order
    color_order
        .iter()
        .map(|name| {
            color_map.get(*name).cloned().unwrap_or_else(|| {
                // Fall back to default theme colors if missing
                let idx = color_order.iter().position(|n| n == name).unwrap_or(0);
                DEFAULT_THEME_COLORS
                    .get(idx)
                    .map(|s| (*s).to_string())
                    .unwrap_or_else(|| "#000000".to_string())
            })
        })
        .collect()
}

/// Parse theme fonts from fontScheme
/// Returns (major_font, minor_font)
/// <a:fontScheme name="Office">
///   <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
///   <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
/// </a:fontScheme>
fn parse_font_scheme<R: BufRead>(xml: &mut Reader<R>) -> (Option<String>, Option<String>) {
    let mut major_font = None;
    let mut minor_font = None;
    let mut buf = Vec::new();
    let mut in_major_font = false;
    let mut in_minor_font = false;
    let mut depth = 1; // We've already entered fontScheme

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                depth += 1;
                let local_name = e.local_name();

                match local_name.as_ref() {
                    b"majorFont" => {
                        in_major_font = true;
                        in_minor_font = false;
                    }
                    b"minorFont" => {
                        in_minor_font = true;
                        in_major_font = false;
                    }
                    b"latin" => {
                        // Get the typeface attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"typeface" {
                                let typeface = String::from_utf8_lossy(&attr.value).to_string();
                                if in_major_font {
                                    major_font = Some(typeface);
                                } else if in_minor_font {
                                    minor_font = Some(typeface);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();

                if local_name.as_ref() == b"latin" {
                    // Get the typeface attribute
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"typeface" {
                            let typeface = String::from_utf8_lossy(&attr.value).to_string();
                            if in_major_font {
                                major_font = Some(typeface);
                            } else if in_minor_font {
                                minor_font = Some(typeface);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                depth -= 1;
                let local_name = e.local_name();

                match local_name.as_ref() {
                    b"majorFont" => {
                        in_major_font = false;
                    }
                    b"minorFont" => {
                        in_minor_font = false;
                    }
                    _ => {}
                }

                // Exit when we close fontScheme
                if depth == 0 {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    (major_font, minor_font)
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
    fn test_parse_color_scheme() {
        let xml_content = r#"
        <a:clrScheme name="Office">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
            <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
            <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
            <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
            <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
            <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
            <a:accent6><a:srgbClr val="F79646"/></a:accent6>
            <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
            <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        // Skip to clrScheme start
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let colors = parse_color_scheme(&mut reader);

        assert_eq!(colors.len(), 12);
        assert_eq!(colors[0], "#FFFFFF"); // lt1 (Background 1)
        assert_eq!(colors[1], "#000000"); // dk1 (Text 1)
        assert_eq!(colors[2], "#EEECE1"); // lt2 (Background 2)
        assert_eq!(colors[3], "#1F497D"); // dk2 (Text 2)
        assert_eq!(colors[4], "#4F81BD"); // accent1
        assert_eq!(colors[5], "#C0504D"); // accent2
        assert_eq!(colors[6], "#9BBB59"); // accent3
        assert_eq!(colors[7], "#8064A2"); // accent4
        assert_eq!(colors[8], "#4BACC6"); // accent5
        assert_eq!(colors[9], "#F79646"); // accent6
        assert_eq!(colors[10], "#0000FF"); // hlink
        assert_eq!(colors[11], "#800080"); // folHlink
    }

    #[test]
    fn test_parse_font_scheme() {
        let xml_content = r#"
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Cambria"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        // Skip to fontScheme start
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let (major, minor) = parse_font_scheme(&mut reader);

        assert_eq!(major, Some("Cambria".to_string()));
        assert_eq!(minor, Some("Calibri".to_string()));
    }

    #[test]
    fn test_parse_font_scheme_empty_elements() {
        // Test with self-closing latin elements
        let xml_content = r#"
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Arial"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Times New Roman"/>
            </a:minorFont>
        </a:fontScheme>
        "#;

        let cursor = Cursor::new(xml_content);
        let mut reader = Reader::from_reader(cursor);
        reader.trim_text(true);

        // Skip to fontScheme start
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontScheme" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                _ => {}
            }
            buf.clear();
        }

        let (major, minor) = parse_font_scheme(&mut reader);

        assert_eq!(major, Some("Arial".to_string()));
        assert_eq!(minor, Some("Times New Roman".to_string()));
    }
}

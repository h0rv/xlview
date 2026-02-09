//! Named styles (cellStyles) parsing module
//! This module handles parsing of named/built-in styles from XLSX styles.xml.

use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::BufRead;

/// Named style definition
#[derive(Debug, Clone)]
pub struct NamedStyle {
    pub name: String,
    pub xf_id: u32,
    pub builtin_id: Option<u32>,
}

/// Cell style XF - base styles that cellXfs inherit from
#[derive(Debug, Clone, Default)]
pub struct CellStyleXf {
    pub num_fmt_id: Option<u32>,
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
}

/// Parse cellStyles from styles.xml
///
/// Expected XML format:
/// ```xml
/// <cellStyles count="2">
///   <cellStyle name="Normal" xfId="0" builtinId="0"/>
///   <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
///   <cellStyle name="Custom" xfId="2"/>
/// </cellStyles>
/// ```
pub fn parse_cell_styles<R: BufRead>(xml: &mut Reader<R>) -> Vec<NamedStyle> {
    let mut styles = Vec::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"cellStyle" => {
                let mut name = String::new();
                let mut xf_id = 0u32;
                let mut builtin_id = None;

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            name = String::from_utf8_lossy(&attr.value).to_string();
                        }
                        b"xfId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                xf_id = s.parse().unwrap_or(0);
                            }
                        }
                        b"builtinId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                builtin_id = s.parse().ok();
                            }
                        }
                        _ => {}
                    }
                }

                styles.push(NamedStyle {
                    name,
                    xf_id,
                    builtin_id,
                });
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"cellStyles" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    styles
}

/// Parse cellStyleXfs (base styles that cellXfs inherit from)
///
/// Expected XML format:
/// ```xml
/// <cellStyleXfs count="2">
///   <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
///   <xf numFmtId="0" fontId="1" fillId="2" borderId="0"/>
/// </cellStyleXfs>
/// ```
pub fn parse_cell_style_xfs<R: BufRead>(xml: &mut Reader<R>) -> Vec<CellStyleXf> {
    let mut xfs = Vec::new();
    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"xf" => {
                let mut cell_style_xf = CellStyleXf::default();

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"numFmtId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                cell_style_xf.num_fmt_id = s.parse().ok();
                            }
                        }
                        b"fontId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                cell_style_xf.font_id = s.parse().ok();
                            }
                        }
                        b"fillId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                cell_style_xf.fill_id = s.parse().ok();
                            }
                        }
                        b"borderId" => {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                cell_style_xf.border_id = s.parse().ok();
                            }
                        }
                        _ => {}
                    }
                }

                xfs.push(cell_style_xf);

                // If it was a Start event, we need to skip to the end of xf element
                if matches!(xml.read_event_into(&mut buf), Ok(Event::Start(_))) {
                    // Skip nested content until we hit End for xf
                    let mut depth = 1;
                    while depth > 0 {
                        buf.clear();
                        match xml.read_event_into(&mut buf) {
                            Ok(Event::Start(_)) => depth += 1,
                            Ok(Event::End(_)) => depth -= 1,
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"cellStyleXfs" => {
                break;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    xfs
}

/// Get builtin style name by ID
///
/// See ECMA-376 Part 1 Section 18.8.7 for the full list of built-in style IDs.
pub fn get_builtin_style_name(builtin_id: u32) -> Option<&'static str> {
    match builtin_id {
        0 => Some("Normal"),
        1 => Some("RowLevel_1"),
        2 => Some("RowLevel_2"),
        3 => Some("RowLevel_3"),
        4 => Some("RowLevel_4"),
        5 => Some("RowLevel_5"),
        6 => Some("RowLevel_6"),
        7 => Some("RowLevel_7"),
        8 => Some("ColLevel_1"),
        9 => Some("ColLevel_2"),
        10 => Some("ColLevel_3"),
        11 => Some("ColLevel_4"),
        12 => Some("ColLevel_5"),
        13 => Some("ColLevel_6"),
        14 => Some("ColLevel_7"),
        15 => Some("Comma"),
        16 => Some("Heading 1"),
        17 => Some("Heading 2"),
        18 => Some("Heading 3"),
        19 => Some("Heading 4"),
        20 => Some("Currency"),
        21 => Some("Comma [0]"),
        22 => Some("Currency [0]"),
        23 => Some("Hyperlink"),
        24 => Some("Followed Hyperlink"),
        25 => Some("Note"),
        26 => Some("Warning Text"),
        27 => Some("Title"),
        28 => Some("Explanatory Text"),
        29 => Some("Input"),
        30 => Some("Output"),
        31 => Some("Calculation"),
        32 => Some("Check Cell"),
        33 => Some("Linked Cell"),
        34 => Some("Total"),
        35 => Some("Good"),
        36 => Some("Bad"),
        37 => Some("Neutral"),
        38 => Some("Accent1"),
        39 => Some("20% - Accent1"),
        40 => Some("40% - Accent1"),
        41 => Some("60% - Accent1"),
        42 => Some("Accent2"),
        43 => Some("20% - Accent2"),
        44 => Some("40% - Accent2"),
        45 => Some("60% - Accent2"),
        46 => Some("Accent3"),
        47 => Some("20% - Accent3"),
        48 => Some("40% - Accent3"),
        49 => Some("60% - Accent3"),
        50 => Some("Accent4"),
        51 => Some("20% - Accent4"),
        52 => Some("40% - Accent4"),
        53 => Some("60% - Accent4"),
        54 => Some("Accent5"),
        55 => Some("20% - Accent5"),
        56 => Some("40% - Accent5"),
        57 => Some("60% - Accent5"),
        58 => Some("Accent6"),
        59 => Some("20% - Accent6"),
        60 => Some("40% - Accent6"),
        61 => Some("60% - Accent6"),
        62 => Some("Percent"),
        _ => None,
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
    fn test_parse_cell_styles() {
        let xml_data = r#"<cellStyles count="3">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
            <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
            <cellStyle name="Custom" xfId="2"/>
        </cellStyles>"#;

        let mut reader = Reader::from_str(xml_data);
        // Skip the opening cellStyles tag
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"cellStyles" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {:?}", e),
                _ => {}
            }
            buf.clear();
        }

        let styles = parse_cell_styles(&mut reader);

        assert_eq!(styles.len(), 3);

        assert_eq!(styles[0].name, "Normal");
        assert_eq!(styles[0].xf_id, 0);
        assert_eq!(styles[0].builtin_id, Some(0));

        assert_eq!(styles[1].name, "Heading 1");
        assert_eq!(styles[1].xf_id, 1);
        assert_eq!(styles[1].builtin_id, Some(16));

        assert_eq!(styles[2].name, "Custom");
        assert_eq!(styles[2].xf_id, 2);
        assert_eq!(styles[2].builtin_id, None);
    }

    #[test]
    fn test_parse_cell_style_xfs() {
        let xml_data = r#"<cellStyleXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            <xf numFmtId="164" fontId="1" fillId="2" borderId="1"/>
        </cellStyleXfs>"#;

        let mut reader = Reader::from_str(xml_data);
        // Skip the opening cellStyleXfs tag
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"cellStyleXfs" => break,
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {:?}", e),
                _ => {}
            }
            buf.clear();
        }

        let xfs = parse_cell_style_xfs(&mut reader);

        assert_eq!(xfs.len(), 2);

        assert_eq!(xfs[0].num_fmt_id, Some(0));
        assert_eq!(xfs[0].font_id, Some(0));
        assert_eq!(xfs[0].fill_id, Some(0));
        assert_eq!(xfs[0].border_id, Some(0));

        assert_eq!(xfs[1].num_fmt_id, Some(164));
        assert_eq!(xfs[1].font_id, Some(1));
        assert_eq!(xfs[1].fill_id, Some(2));
        assert_eq!(xfs[1].border_id, Some(1));
    }

    #[test]
    fn test_get_builtin_style_name() {
        assert_eq!(get_builtin_style_name(0), Some("Normal"));
        assert_eq!(get_builtin_style_name(16), Some("Heading 1"));
        assert_eq!(get_builtin_style_name(17), Some("Heading 2"));
        assert_eq!(get_builtin_style_name(27), Some("Title"));
        assert_eq!(get_builtin_style_name(35), Some("Good"));
        assert_eq!(get_builtin_style_name(36), Some("Bad"));
        assert_eq!(get_builtin_style_name(62), Some("Percent"));
        assert_eq!(get_builtin_style_name(999), None);
    }
}

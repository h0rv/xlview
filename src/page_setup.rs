//! Page setup and print settings parsing module
//! This module handles parsing of page margins, orientation, headers/footers from XLSX files.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

/// Page margins in inches
#[derive(Debug, Clone, Default)]
pub struct PageMargins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

/// Page orientation
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum Orientation {
    #[default]
    Portrait,
    Landscape,
}

/// Page setup configuration
#[derive(Debug, Clone, Default)]
pub struct PageSetup {
    /// Paper size code (e.g., 1=Letter, 9=A4)
    pub paper_size: Option<u32>,
    /// Page orientation
    pub orientation: Orientation,
    /// Print scale percentage (10-400)
    pub scale: Option<u32>,
    /// Number of pages wide to fit to (0 = don't fit)
    pub fit_to_width: Option<u32>,
    /// Number of pages tall to fit to (0 = don't fit)
    pub fit_to_height: Option<u32>,
}

/// Header and footer text for printing
#[derive(Debug, Clone, Default)]
pub struct HeaderFooter {
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

/// Parse pageMargins element
///
/// Example XML:
/// ```xml
/// <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
/// ```
pub fn parse_page_margins(e: &BytesStart) -> PageMargins {
    let mut margins = PageMargins::default();

    for attr in e.attributes().flatten() {
        let value: f64 = std::str::from_utf8(&attr.value)
            .ok()
            .and_then(|s| s.parse().ok())
            .unwrap_or(0.0);

        match attr.key.as_ref() {
            b"left" => margins.left = value,
            b"right" => margins.right = value,
            b"top" => margins.top = value,
            b"bottom" => margins.bottom = value,
            b"header" => margins.header = value,
            b"footer" => margins.footer = value,
            _ => {}
        }
    }

    margins
}

/// Parse pageSetup element
///
/// Example XML:
/// ```xml
/// <pageSetup paperSize="9" orientation="landscape" scale="100" fitToWidth="1" fitToHeight="0"/>
/// ```
pub fn parse_page_setup(e: &BytesStart) -> PageSetup {
    let mut setup = PageSetup::default();

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"paperSize" => {
                setup.paper_size = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"orientation" => {
                let orientation_str = std::str::from_utf8(&attr.value).unwrap_or("");
                setup.orientation = match orientation_str {
                    "landscape" => Orientation::Landscape,
                    _ => Orientation::Portrait,
                };
            }
            b"scale" => {
                setup.scale = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"fitToWidth" => {
                setup.fit_to_width = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"fitToHeight" => {
                setup.fit_to_height = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            _ => {}
        }
    }

    setup
}

/// Parse headerFooter element
///
/// Example XML:
/// ```xml
/// <headerFooter>
///   <oddHeader>&amp;C&amp;"Arial,Bold"&amp;12Header</oddHeader>
///   <oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter>
///   <evenHeader>Even Header Text</evenHeader>
///   <evenFooter>Even Footer Text</evenFooter>
///   <firstHeader>First Page Header</firstHeader>
///   <firstFooter>First Page Footer</firstFooter>
/// </headerFooter>
/// ```
pub fn parse_header_footer<R: BufRead>(xml: &mut Reader<R>) -> HeaderFooter {
    let mut header_footer = HeaderFooter::default();
    let mut buf = Vec::new();
    let mut current_element: Option<String> = None;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                current_element = Some(name.to_string());
            }
            Ok(Event::Text(ref e)) => {
                if let Some(ref element_name) = current_element {
                    if let Ok(text) = e.unescape() {
                        let text_str = text.to_string();
                        match element_name.as_str() {
                            "oddHeader" => header_footer.odd_header = Some(text_str),
                            "oddFooter" => header_footer.odd_footer = Some(text_str),
                            "evenHeader" => header_footer.even_header = Some(text_str),
                            "evenFooter" => header_footer.even_footer = Some(text_str),
                            "firstHeader" => header_footer.first_header = Some(text_str),
                            "firstFooter" => header_footer.first_footer = Some(text_str),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                // Clear current element when we exit it
                if current_element.as_deref() == Some(name) {
                    current_element = None;
                }

                // Exit when we reach the end of headerFooter
                if name == "headerFooter" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    header_footer
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
    use quick_xml::Reader;

    #[test]
    fn test_parse_page_margins() {
        let xml = r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let margins = parse_page_margins(e);
            assert!((margins.left - 0.7).abs() < 0.001);
            assert!((margins.right - 0.7).abs() < 0.001);
            assert!((margins.top - 0.75).abs() < 0.001);
            assert!((margins.bottom - 0.75).abs() < 0.001);
            assert!((margins.header - 0.3).abs() < 0.001);
            assert!((margins.footer - 0.3).abs() < 0.001);
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_page_setup_landscape() {
        let xml = r#"<pageSetup paperSize="9" orientation="landscape" scale="100" fitToWidth="1" fitToHeight="0"/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let setup = parse_page_setup(e);
            assert_eq!(setup.paper_size, Some(9));
            assert_eq!(setup.orientation, Orientation::Landscape);
            assert_eq!(setup.scale, Some(100));
            assert_eq!(setup.fit_to_width, Some(1));
            assert_eq!(setup.fit_to_height, Some(0));
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_page_setup_portrait() {
        let xml = r#"<pageSetup paperSize="1" orientation="portrait"/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let setup = parse_page_setup(e);
            assert_eq!(setup.paper_size, Some(1));
            assert_eq!(setup.orientation, Orientation::Portrait);
            assert_eq!(setup.scale, None);
            assert_eq!(setup.fit_to_width, None);
            assert_eq!(setup.fit_to_height, None);
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_header_footer() {
        let xml = r#"<headerFooter>
            <oddHeader>&amp;C&amp;"Arial,Bold"&amp;12Header</oddHeader>
            <oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter>
        </headerFooter>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);
        let mut buf = Vec::new();

        // Skip the start element
        if let Ok(Event::Start(_)) = reader.read_event_into(&mut buf) {
            let hf = parse_header_footer(&mut reader);
            assert!(hf.odd_header.is_some());
            assert!(hf.odd_footer.is_some());
            assert!(hf.even_header.is_none());
            assert!(hf.even_footer.is_none());
            assert!(hf.first_header.is_none());
            assert!(hf.first_footer.is_none());
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_header_footer_all_fields() {
        let xml = r#"<headerFooter>
            <oddHeader>Odd Header</oddHeader>
            <oddFooter>Odd Footer</oddFooter>
            <evenHeader>Even Header</evenHeader>
            <evenFooter>Even Footer</evenFooter>
            <firstHeader>First Header</firstHeader>
            <firstFooter>First Footer</firstFooter>
        </headerFooter>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);
        let mut buf = Vec::new();

        // Skip the start element
        if let Ok(Event::Start(_)) = reader.read_event_into(&mut buf) {
            let hf = parse_header_footer(&mut reader);
            assert_eq!(hf.odd_header, Some("Odd Header".to_string()));
            assert_eq!(hf.odd_footer, Some("Odd Footer".to_string()));
            assert_eq!(hf.even_header, Some("Even Header".to_string()));
            assert_eq!(hf.even_footer, Some("Even Footer".to_string()));
            assert_eq!(hf.first_header, Some("First Header".to_string()));
            assert_eq!(hf.first_footer, Some("First Footer".to_string()));
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_page_margins_defaults() {
        let xml = r#"<pageMargins/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let margins = parse_page_margins(e);
            assert!((margins.left - 0.0).abs() < 0.001);
            assert!((margins.right - 0.0).abs() < 0.001);
            assert!((margins.top - 0.0).abs() < 0.001);
            assert!((margins.bottom - 0.0).abs() < 0.001);
            assert!((margins.header - 0.0).abs() < 0.001);
            assert!((margins.footer - 0.0).abs() < 0.001);
        } else {
            panic!("Failed to parse XML");
        }
    }

    #[test]
    fn test_parse_page_setup_defaults() {
        let xml = r#"<pageSetup/>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let setup = parse_page_setup(e);
            assert_eq!(setup.paper_size, None);
            assert_eq!(setup.orientation, Orientation::Portrait);
            assert_eq!(setup.scale, None);
            assert_eq!(setup.fit_to_width, None);
            assert_eq!(setup.fit_to_height, None);
        } else {
            panic!("Failed to parse XML");
        }
    }
}

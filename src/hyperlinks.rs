//! Hyperlink parsing module
//! This module handles parsing of hyperlinks from XLSX files.

use crate::types::Hyperlink;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{BufRead, BufReader, Read, Seek};
use zip::ZipArchive;

/// Intermediate hyperlink data parsed from sheet XML
/// Contains the r:id reference that needs to be resolved via relationships
#[derive(Debug, Clone)]
pub struct RawHyperlink {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Relationship ID for external links (e.g., "rId1")
    pub r_id: Option<String>,
    /// Internal location (e.g., "Sheet2!A1" for internal links)
    pub location: Option<String>,
    /// Display text
    pub display: Option<String>,
    /// Tooltip text
    pub tooltip: Option<String>,
}

/// Parse hyperlinks from sheet XML
/// Returns a vector of RawHyperlink structs that need relationship resolution
pub fn parse_hyperlinks<R: BufRead>(xml: &mut Reader<R>) -> Vec<RawHyperlink> {
    let mut hyperlinks = Vec::new();
    let mut buf = Vec::new();
    let mut in_hyperlinks = false;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "hyperlinks" {
                    in_hyperlinks = true;
                } else if name == "hyperlink" && in_hyperlinks {
                    if let Some(link) = parse_hyperlink_element(e) {
                        hyperlinks.push(link);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "hyperlinks" {
                    // Empty hyperlinks element, nothing to do
                    in_hyperlinks = false;
                } else if name == "hyperlink" && in_hyperlinks {
                    if let Some(link) = parse_hyperlink_element(e) {
                        hyperlinks.push(link);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "hyperlinks" {
                    // We've finished parsing all hyperlinks
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    hyperlinks
}

/// Parse a single hyperlink element
fn parse_hyperlink_element(e: &quick_xml::events::BytesStart<'_>) -> Option<RawHyperlink> {
    let mut cell_ref = String::new();
    let mut r_id: Option<String> = None;
    let mut location: Option<String> = None;
    let mut display: Option<String> = None;
    let mut tooltip: Option<String> = None;

    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let value = std::str::from_utf8(&attr.value).unwrap_or("");

        match key {
            b"ref" => {
                cell_ref = value.to_string();
            }
            // r:id attribute (namespace prefixed) - for external hyperlinks
            key if key.ends_with(b":id") || key == b"id" => {
                if !value.is_empty() {
                    r_id = Some(value.to_string());
                }
            }
            b"location" => {
                if !value.is_empty() {
                    location = Some(value.to_string());
                }
            }
            b"display" => {
                if !value.is_empty() {
                    display = Some(value.to_string());
                }
            }
            b"tooltip" => {
                if !value.is_empty() {
                    tooltip = Some(value.to_string());
                }
            }
            _ => {}
        }
    }

    if cell_ref.is_empty() {
        return None;
    }

    Some(RawHyperlink {
        cell_ref,
        r_id,
        location,
        display,
        tooltip,
    })
}

/// Parse sheet relationships to get hyperlink targets
/// Returns HashMap of rId -> target URL
pub fn parse_hyperlink_rels<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_path: &str,
) -> HashMap<String, String> {
    let mut rels = HashMap::new();

    // Construct the relationship file path
    // For "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
    let rels_path = construct_rels_path(sheet_path);

    let Ok(file) = archive.by_name(&rels_path) else {
        return rels; // Relationships file is optional
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = String::new();
                    let mut target = String::new();
                    let mut rel_type = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => {
                                id = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Target" => {
                                target = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Type" => {
                                rel_type =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            _ => {}
                        }
                    }

                    // Only include hyperlink relationships
                    if rel_type.contains("hyperlink") && !id.is_empty() && !target.is_empty() {
                        rels.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    rels
}

/// Construct the relationships file path from a sheet path
/// e.g., "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
fn construct_rels_path(sheet_path: &str) -> String {
    if let Some(pos) = sheet_path.rfind('/') {
        let dir = &sheet_path[..pos];
        let filename = &sheet_path[pos + 1..];
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{sheet_path}.rels")
    }
}

/// Resolve raw hyperlinks with relationship targets to create final Hyperlink structs
pub fn resolve_hyperlinks(
    raw_hyperlinks: &[RawHyperlink],
    rels: &HashMap<String, String>,
) -> Vec<(String, Hyperlink)> {
    raw_hyperlinks
        .iter()
        .filter_map(|raw| {
            let (target, is_external) = if let Some(ref r_id) = raw.r_id {
                // External link - resolve via relationship
                if let Some(url) = rels.get(r_id) {
                    (url.clone(), true)
                } else {
                    // r_id not found in relationships, skip this hyperlink
                    return None;
                }
            } else if let Some(ref loc) = raw.location {
                // Internal link (no r:id, only location)
                (loc.clone(), false)
            } else {
                // No target information, skip
                return None;
            };

            let hyperlink = Hyperlink {
                target,
                location: raw.location.clone(),
                tooltip: raw.tooltip.clone(),
                is_external,
            };

            Some((raw.cell_ref.clone(), hyperlink))
        })
        .collect()
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
    fn test_construct_rels_path() {
        assert_eq!(
            construct_rels_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(
            construct_rels_path("xl/worksheets/sheet2.xml"),
            "xl/worksheets/_rels/sheet2.xml.rels"
        );
        assert_eq!(construct_rels_path("sheet.xml"), "_rels/sheet.xml.rels");
    }

    #[test]
    fn test_parse_hyperlinks_xml() {
        let xml_data = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
    <hyperlinks>
        <hyperlink ref="A1" r:id="rId1" display="Click here" tooltip="Visit site"/>
        <hyperlink ref="B2" location="Sheet2!A1" display="Go to Sheet2"/>
        <hyperlink ref="C3" r:id="rId2"/>
    </hyperlinks>
</worksheet>"#;

        let mut reader = Reader::from_str(xml_data);
        reader.trim_text(true);

        // Skip to hyperlinks element
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hyperlinks" => {
                    break;
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }

        // Now parse hyperlinks - we need to create a new reader positioned before hyperlinks
        let mut reader2 = Reader::from_str(xml_data);
        reader2.trim_text(true);

        let hyperlinks = parse_hyperlinks(&mut reader2);

        assert_eq!(hyperlinks.len(), 3);

        // First hyperlink - external with display and tooltip
        assert_eq!(hyperlinks[0].cell_ref, "A1");
        assert!(hyperlinks[0].r_id.is_some());
        assert_eq!(hyperlinks[0].display.as_deref(), Some("Click here"));
        assert_eq!(hyperlinks[0].tooltip.as_deref(), Some("Visit site"));

        // Second hyperlink - internal with location
        assert_eq!(hyperlinks[1].cell_ref, "B2");
        assert!(hyperlinks[1].r_id.is_none());
        assert_eq!(hyperlinks[1].location.as_deref(), Some("Sheet2!A1"));
        assert_eq!(hyperlinks[1].display.as_deref(), Some("Go to Sheet2"));

        // Third hyperlink - external without display
        assert_eq!(hyperlinks[2].cell_ref, "C3");
        assert!(hyperlinks[2].r_id.is_some());
    }

    #[test]
    fn test_resolve_hyperlinks() {
        let raw = vec![
            RawHyperlink {
                cell_ref: "A1".to_string(),
                r_id: Some("rId1".to_string()),
                location: None,
                display: Some("Google".to_string()),
                tooltip: Some("Visit Google".to_string()),
            },
            RawHyperlink {
                cell_ref: "B2".to_string(),
                r_id: None,
                location: Some("Sheet2!A1".to_string()),
                display: Some("Go to Sheet2".to_string()),
                tooltip: None,
            },
            RawHyperlink {
                cell_ref: "C3".to_string(),
                r_id: Some("rId999".to_string()), // Non-existent
                location: None,
                display: None,
                tooltip: None,
            },
        ];

        let mut rels = HashMap::new();
        rels.insert("rId1".to_string(), "https://www.google.com".to_string());

        let resolved = resolve_hyperlinks(&raw, &rels);

        // Should have 2 resolved hyperlinks (rId999 not found)
        assert_eq!(resolved.len(), 2);

        // First - external
        assert_eq!(resolved[0].0, "A1");
        assert_eq!(resolved[0].1.target, "https://www.google.com");
        assert!(resolved[0].1.is_external);
        assert_eq!(resolved[0].1.tooltip.as_deref(), Some("Visit Google"));

        // Second - internal
        assert_eq!(resolved[1].0, "B2");
        assert_eq!(resolved[1].1.target, "Sheet2!A1");
        assert!(!resolved[1].1.is_external);
        assert_eq!(resolved[1].1.location.as_deref(), Some("Sheet2!A1"));
    }
}

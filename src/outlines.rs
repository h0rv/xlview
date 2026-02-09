//! Row/column outline (grouping) parsing module
//! This module handles parsing of outline levels and grouping from XLSX files.
//!
//! Excel outlines allow users to group rows or columns into hierarchical levels
//! that can be collapsed and expanded. This is commonly used for:
//! - Subtotals and summary rows
//! - Hierarchical data like organizational charts
//! - Collapsible sections in large spreadsheets
//!
//! Key XLSX elements parsed:
//! - `<row outlineLevel="N" collapsed="1">` - Row outline levels
//! - `<col outlineLevel="N" collapsed="1">` - Column outline levels
//! - `<sheetFormatPr>` - Default outline level settings
//! - `<outlinePr summaryBelow="0" summaryRight="0">` - Summary position settings

use quick_xml::events::BytesStart;

use crate::types::OutlineLevel;

/// Parse outline level from a row element.
///
/// The row element in XLSX can contain outline information:
/// ```xml
/// <row r="5" outlineLevel="1" collapsed="1" hidden="1">
/// ```
///
/// # Arguments
/// * `e` - The BytesStart event for the row element
///
/// # Returns
/// * `Some(OutlineLevel)` if the row has an outline level > 0
/// * `None` if the row has no outline or outline level is 0
pub fn parse_row_outline(e: &BytesStart) -> Option<OutlineLevel> {
    let mut row_index: Option<u32> = None;
    let mut outline_level: u8 = 0;
    let mut collapsed = false;
    let mut hidden = false;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"r" => {
                // Row number is 1-based in XLSX, convert to 0-based
                row_index = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse::<u32>().ok())
                    .map(|r| r.saturating_sub(1));
            }
            b"outlineLevel" => {
                outline_level = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            b"collapsed" => {
                collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"hidden" => {
                hidden = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            _ => {}
        }
    }

    // Only return an OutlineLevel if there's actually an outline level set
    if outline_level > 0 {
        Some(OutlineLevel {
            index: row_index.unwrap_or(0),
            level: outline_level,
            collapsed,
            hidden,
        })
    } else {
        None
    }
}

/// Parse outline levels from a col element.
///
/// Column elements in XLSX can span multiple columns and contain outline information:
/// ```xml
/// <col min="2" max="4" outlineLevel="2" collapsed="0" hidden="0"/>
/// ```
///
/// # Arguments
/// * `e` - The BytesStart event for the col element
/// * `min` - The minimum column index (1-based from the XML)
/// * `max` - The maximum column index (1-based from the XML)
///
/// # Returns
/// A vector of OutlineLevel, one for each column in the range that has an outline level > 0.
/// Returns empty vector if no outline level is set.
pub fn parse_col_outline(e: &BytesStart, min: u32, max: u32) -> Vec<OutlineLevel> {
    let mut outline_level: u8 = 0;
    let mut collapsed = false;
    let mut hidden = false;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"outlineLevel" => {
                outline_level = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
            }
            b"collapsed" => {
                collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"hidden" => {
                hidden = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            _ => {}
        }
    }

    // Only create OutlineLevels if there's actually an outline level set
    if outline_level == 0 {
        return Vec::new();
    }

    // Create an OutlineLevel for each column in the range
    // min and max are 1-based in XLSX, convert to 0-based
    let start_col = min.saturating_sub(1);
    let end_col = max.saturating_sub(1);

    (start_col..=end_col)
        .map(|col| OutlineLevel {
            index: col,
            level: outline_level,
            collapsed,
            hidden,
        })
        .collect()
}

/// Parse outline/summary position properties from sheetFormatPr or outlinePr elements.
///
/// These elements control where summary rows/columns are positioned relative to detail:
///
/// ```xml
/// <!-- From sheetPr/outlinePr -->
/// <outlinePr summaryBelow="0" summaryRight="0"/>
///
/// <!-- From sheetFormatPr (less common for these settings) -->
/// <sheetFormatPr outlineLevelRow="2" outlineLevelCol="1"/>
/// ```
///
/// # Arguments
/// * `e` - The BytesStart event for either sheetFormatPr or outlinePr element
///
/// # Returns
/// A tuple of (summary_below, summary_right):
/// - `summary_below`: true if summary rows appear below detail rows (default: true)
/// - `summary_right`: true if summary columns appear to the right of detail columns (default: true)
pub fn parse_outline_properties(e: &BytesStart) -> (bool, bool) {
    let mut summary_below = true; // Excel default
    let mut summary_right = true; // Excel default

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"summaryBelow" => {
                // "0" = summary above, "1" or absent = summary below
                summary_below = std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
            }
            b"summaryRight" => {
                // "0" = summary left, "1" or absent = summary right
                summary_right = std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
            }
            _ => {}
        }
    }

    (summary_below, summary_right)
}

/// Represents outline property settings parsed from outlinePr
#[derive(Debug, Clone)]
pub struct OutlineProperties {
    /// Summary rows are below detail rows (default: true)
    pub summary_below: bool,
    /// Summary columns are right of detail columns (default: true)
    pub summary_right: bool,
}

impl Default for OutlineProperties {
    fn default() -> Self {
        Self {
            summary_below: true,
            summary_right: true,
        }
    }
}

/// Parse outlinePr element for outline summary settings.
///
/// This is a more structured version that returns an OutlineProperties struct.
///
/// ```xml
/// <outlinePr summaryBelow="0" summaryRight="0"/>
/// ```
///
/// # Arguments
/// * `e` - The BytesStart event for the outlinePr element
///
/// # Returns
/// An OutlineProperties struct with the parsed settings
pub fn parse_outline_pr(e: &BytesStart) -> OutlineProperties {
    let (summary_below, summary_right) = parse_outline_properties(e);
    OutlineProperties {
        summary_below,
        summary_right,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::cast_possible_truncation,
    clippy::panic
)]
mod tests {
    use super::*;
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn parse_element(xml: &str) -> BytesStart<'static> {
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    return e.into_owned();
                }
                Ok(Event::Eof) => panic!("Unexpected EOF"),
                Err(e) => panic!("Error: {e}"),
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_row_outline_with_level() {
        let xml = r#"<row r="5" outlineLevel="2" collapsed="1" hidden="1"/>"#;
        let e = parse_element(xml);
        let result = parse_row_outline(&e).unwrap();

        assert_eq!(result.index, 4); // 0-based
        assert_eq!(result.level, 2);
        assert!(result.collapsed);
        assert!(result.hidden);
    }

    #[test]
    fn test_parse_row_outline_no_level() {
        let xml = r#"<row r="5"/>"#;
        let e = parse_element(xml);
        let result = parse_row_outline(&e);

        assert!(result.is_none());
    }

    #[test]
    fn test_parse_row_outline_level_zero() {
        let xml = r#"<row r="5" outlineLevel="0"/>"#;
        let e = parse_element(xml);
        let result = parse_row_outline(&e);

        assert!(result.is_none());
    }

    #[test]
    fn test_parse_col_outline_single_column() {
        let xml = r#"<col min="3" max="3" outlineLevel="1" collapsed="0"/>"#;
        let e = parse_element(xml);
        let result = parse_col_outline(&e, 3, 3);

        assert_eq!(result.len(), 1);
        assert_eq!(result[0].index, 2); // 0-based
        assert_eq!(result[0].level, 1);
        assert!(!result[0].collapsed);
    }

    #[test]
    fn test_parse_col_outline_multiple_columns() {
        let xml = r#"<col min="2" max="4" outlineLevel="2" collapsed="1" hidden="1"/>"#;
        let e = parse_element(xml);
        let result = parse_col_outline(&e, 2, 4);

        assert_eq!(result.len(), 3);
        for (i, outline) in result.iter().enumerate() {
            assert_eq!(outline.index, (i + 1) as u32); // 0-based: 1, 2, 3
            assert_eq!(outline.level, 2);
            assert!(outline.collapsed);
            assert!(outline.hidden);
        }
    }

    #[test]
    fn test_parse_col_outline_no_level() {
        let xml = r#"<col min="2" max="4" width="10"/>"#;
        let e = parse_element(xml);
        let result = parse_col_outline(&e, 2, 4);

        assert!(result.is_empty());
    }

    #[test]
    fn test_parse_outline_properties_defaults() {
        let xml = r#"<outlinePr/>"#;
        let e = parse_element(xml);
        let (summary_below, summary_right) = parse_outline_properties(&e);

        assert!(summary_below);
        assert!(summary_right);
    }

    #[test]
    fn test_parse_outline_properties_summary_above() {
        let xml = r#"<outlinePr summaryBelow="0"/>"#;
        let e = parse_element(xml);
        let (summary_below, summary_right) = parse_outline_properties(&e);

        assert!(!summary_below);
        assert!(summary_right);
    }

    #[test]
    fn test_parse_outline_properties_summary_left() {
        let xml = r#"<outlinePr summaryRight="0"/>"#;
        let e = parse_element(xml);
        let (summary_below, summary_right) = parse_outline_properties(&e);

        assert!(summary_below);
        assert!(!summary_right);
    }

    #[test]
    fn test_parse_outline_properties_both_reversed() {
        let xml = r#"<outlinePr summaryBelow="0" summaryRight="0"/>"#;
        let e = parse_element(xml);
        let (summary_below, summary_right) = parse_outline_properties(&e);

        assert!(!summary_below);
        assert!(!summary_right);
    }

    #[test]
    fn test_parse_outline_pr_struct() {
        let xml = r#"<outlinePr summaryBelow="0" summaryRight="1"/>"#;
        let e = parse_element(xml);
        let props = parse_outline_pr(&e);

        assert!(!props.summary_below);
        assert!(props.summary_right);
    }
}

//! Workbook metadata parsing module
//! This module handles parsing workbook.xml for sheets, date system, defined names.

use crate::error::Result;
use crate::types::{DefinedName, Sheet, SheetState};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use zip::ZipArchive;

/// Excel date system - determines how serial dates are interpreted
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    /// Windows 1900 date system (default) - serial date 1 = January 1, 1900
    /// Note: Excel incorrectly treats 1900 as a leap year for compatibility
    #[default]
    Date1900,
    /// Mac 1904 date system - serial date 0 = January 1, 1904
    /// Used historically by Mac Excel, still supported for compatibility
    Date1904,
}

/// Workbook relationships parsed from xl/_rels/workbook.xml.rels
///
/// Contains paths to all related files in the workbook package.
/// Paths are resolved relative to the xl/ directory and stored as full paths.
#[derive(Default, Debug)]
pub struct WorkbookRelationships {
    /// Map of rId -> full path for worksheet relationships
    /// e.g., "rId1" -> "xl/worksheets/sheet1.xml"
    pub worksheets: HashMap<String, String>,
    /// Path to shared strings file (e.g., "xl/sharedStrings.xml")
    pub shared_strings: Option<String>,
    /// Path to styles file (e.g., "xl/styles.xml")
    pub styles: Option<String>,
    /// Path to theme file (e.g., "xl/theme/theme1.xml")
    pub theme: Option<String>,
}

/// Sheet metadata from workbook.xml
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub path: String,
    pub state: SheetState,
}

/// Workbook metadata parsed from workbook.xml
#[derive(Debug)]
pub struct WorkbookMeta {
    pub sheets: Vec<SheetInfo>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
}

/// Parse workbook relationships (xl/_rels/workbook.xml.rels)
///
/// Parses the relationships file to get paths for:
/// - Worksheets (mapped by rId)
/// - Shared strings
/// - Styles
/// - Theme
pub fn parse_workbook_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> WorkbookRelationships {
    let mut rels = WorkbookRelationships::default();

    let Ok(file) = archive.by_name("xl/_rels/workbook.xml.rels") else {
        return rels; // Relationships file is optional
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
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

                    // Resolve target path relative to xl/
                    let full_path = resolve_relationship_path(&target);

                    // Categorize by relationship type
                    if rel_type.contains("worksheet") && !id.is_empty() && !target.is_empty() {
                        rels.worksheets.insert(id, full_path);
                    } else if rel_type.contains("sharedStrings") {
                        rels.shared_strings = Some(full_path);
                    } else if rel_type.contains("/styles") {
                        rels.styles = Some(full_path);
                    } else if rel_type.contains("/theme") {
                        rels.theme = Some(full_path);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    rels
}

/// Resolve a relationship target path to a full path within the archive
fn resolve_relationship_path(target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        // Absolute path from archive root
        stripped.to_string()
    } else if target.starts_with("../") {
        // Relative path going up from xl/ directory
        // e.g., "../customXml/item1.xml" -> "customXml/item1.xml"
        let mut path = target;
        while let Some(stripped) = path.strip_prefix("../") {
            path = stripped;
        }
        path.to_string()
    } else {
        // Relative path within xl/ directory
        format!("xl/{target}")
    }
}

/// Parse workbook.xml for sheet info, date system, and defined names
///
/// Parses:
/// - `<sheets><sheet name="..." r:id="rId1" state="hidden"/></sheets>`
/// - `<workbookPr date1904="1"/>`
/// - `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$10</definedName></definedNames>`
pub fn parse_workbook_xml<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    relationships: &HashMap<String, String>,
) -> Result<WorkbookMeta> {
    let file = archive.by_name("xl/workbook.xml")?;

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut sheets = Vec::new();
    let mut date_system = DateSystem::Date1900;
    let mut defined_names = Vec::new();
    let mut buf = Vec::new();

    // Track parsing state
    let mut in_defined_names = false;
    let mut current_defined_name: Option<DefinedNameBuilder> = None;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name_str = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name_str {
                    "definedNames" => {
                        in_defined_names = true;
                    }
                    "definedName" if in_defined_names => {
                        current_defined_name = Some(parse_defined_name_attributes(e));
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name_str = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name_str {
                    "sheet" => {
                        if let Some(info) = parse_sheet_element(e, relationships, sheets.len()) {
                            sheets.push(info);
                        }
                    }
                    "workbookPr" => {
                        date_system = parse_workbook_pr(e);
                    }
                    "definedName" if in_defined_names => {
                        // Empty definedName element (unlikely but handle it)
                        let builder = parse_defined_name_attributes(e);
                        if !builder.name.is_empty() {
                            defined_names.push(builder.build(String::new()));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Some(ref mut builder) = current_defined_name {
                    if let Ok(text) = e.unescape() {
                        builder.value.push_str(&text);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name_str = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name_str {
                    "definedNames" => {
                        in_defined_names = false;
                    }
                    "definedName" => {
                        if let Some(builder) = current_defined_name.take() {
                            if !builder.name.is_empty() {
                                let value = builder.value.clone();
                                defined_names.push(builder.build(value));
                            }
                        }
                    }
                    "workbookPr" => {
                        // workbookPr can also be a start/end element pair
                        // date system already parsed from attributes in Start event
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(WorkbookMeta {
        sheets,
        date_system,
        defined_names,
    })
}

/// Helper struct for building DefinedName
struct DefinedNameBuilder {
    name: String,
    local_sheet_id: Option<u32>,
    hidden: bool,
    comment: Option<String>,
    value: String,
}

impl DefinedNameBuilder {
    fn build(self, value: String) -> DefinedName {
        DefinedName {
            name: self.name,
            value: if self.value.is_empty() {
                value
            } else {
                self.value
            },
            local_sheet_id: self.local_sheet_id,
            hidden: self.hidden,
            comment: self.comment,
        }
    }
}

/// Parse attributes from a definedName element
fn parse_defined_name_attributes(e: &quick_xml::events::BytesStart<'_>) -> DefinedNameBuilder {
    let mut name = String::new();
    let mut local_sheet_id = None;
    let mut hidden = false;
    let mut comment = None;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => {
                name = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"localSheetId" => {
                local_sheet_id = std::str::from_utf8(&attr.value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"hidden" => {
                hidden = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            b"comment" => {
                let c = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                if !c.is_empty() {
                    comment = Some(c);
                }
            }
            _ => {}
        }
    }

    DefinedNameBuilder {
        name,
        local_sheet_id,
        hidden,
        comment,
        value: String::new(),
    }
}

/// Parse a sheet element and return SheetInfo
fn parse_sheet_element(
    e: &quick_xml::events::BytesStart<'_>,
    relationships: &HashMap<String, String>,
    sheet_index: usize,
) -> Option<SheetInfo> {
    let mut name = String::new();
    let mut r_id = String::new();
    let mut state = SheetState::Visible;

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"name" => {
                name = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            b"state" => {
                let state_str = std::str::from_utf8(&attr.value).unwrap_or("");
                state = match state_str {
                    "hidden" => SheetState::Hidden,
                    "veryHidden" => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
            }
            // r:id attribute (namespace prefixed)
            key if key.ends_with(b":id") || key == b"id" => {
                r_id = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
            }
            _ => {}
        }
    }

    if name.is_empty() {
        return None;
    }

    // Try to get path from relationships, fallback to default
    let path = relationships.get(&r_id).cloned().unwrap_or_else(|| {
        let idx = sheet_index + 1;
        format!("xl/worksheets/sheet{idx}.xml")
    });

    Some(SheetInfo { name, path, state })
}

/// Parse workbookPr element for date system
fn parse_workbook_pr(e: &quick_xml::events::BytesStart<'_>) -> DateSystem {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"date1904" {
            let val = std::str::from_utf8(&attr.value).unwrap_or("0");
            if val == "1" || val == "true" {
                return DateSystem::Date1904;
            }
        }
    }
    DateSystem::Date1900
}

/// Apply defined names to sheets (print area, print titles, etc.)
///
/// Looks for special built-in names:
/// - `_xlnm.Print_Area` - Print area for a sheet
/// - `_xlnm.Print_Titles` - Print titles (rows/columns to repeat on each page)
///
/// Parses the formula values and applies them to the appropriate sheet.
pub fn apply_defined_names_to_sheets(sheets: &mut [Sheet], defined_names: &[DefinedName]) {
    for dn in defined_names {
        // Skip names that aren't scoped to a specific sheet
        let Some(local_sheet_id) = dn.local_sheet_id else {
            continue;
        };

        let sheet_idx = local_sheet_id as usize;
        let Some(sheet) = sheets.get_mut(sheet_idx) else {
            continue;
        };

        match dn.name.as_str() {
            "_xlnm.Print_Area" => {
                // Print area format: "Sheet1!$A$1:$D$10" or "'Sheet Name'!$A$1:$D$10"
                // We need to extract just the range part
                if let Some(range) = extract_range_from_formula(&dn.value) {
                    sheet.print_area = Some(range);
                }
            }
            "_xlnm.Print_Titles" => {
                // Print titles can contain row ranges, column ranges, or both
                // Format: "Sheet1!$1:$3" (rows) or "Sheet1!$A:$B" (columns)
                // Combined: "Sheet1!$A:$B,Sheet1!$1:$3"
                parse_print_titles(&dn.value, sheet);
            }
            _ => {}
        }
    }
}

/// Extract the range part from a formula like "Sheet1!$A$1:$D$10"
fn extract_range_from_formula(formula: &str) -> Option<String> {
    // Handle formulas with sheet reference
    if let Some(pos) = formula.rfind('!') {
        let range_part = &formula[pos + 1..];
        // Remove $ signs and return the clean range
        Some(clean_range(range_part))
    } else {
        // No sheet reference, just a range
        Some(clean_range(formula))
    }
}

/// Remove $ signs from a range reference
fn clean_range(range: &str) -> String {
    range.replace('$', "")
}

/// Parse print titles formula and apply to sheet
fn parse_print_titles(formula: &str, sheet: &mut Sheet) {
    // Split by comma for combined row/column titles
    for part in formula.split(',') {
        let part = part.trim();

        // Extract the range part after the sheet reference
        let range_part = if let Some(pos) = part.rfind('!') {
            &part[pos + 1..]
        } else {
            part
        };

        // Check if it's a row range (e.g., $1:$3) or column range (e.g., $A:$B)
        if let Some((start, end)) = parse_range_bounds(range_part) {
            if is_row_reference(&start) && is_row_reference(&end) {
                // Row titles
                if let (Some(start_row), Some(end_row)) = (parse_row(&start), parse_row(&end)) {
                    sheet.print_titles_rows = Some((start_row, end_row));
                }
            } else if is_column_reference(&start) && is_column_reference(&end) {
                // Column titles
                if let (Some(start_col), Some(end_col)) = (parse_column(&start), parse_column(&end))
                {
                    sheet.print_titles_cols = Some((start_col, end_col));
                }
            }
        }
    }
}

/// Parse range bounds from a range like "$1:$3" or "$A:$B"
fn parse_range_bounds(range: &str) -> Option<(String, String)> {
    let clean = range.replace('$', "");
    let parts: Vec<&str> = clean.split(':').collect();
    let first = parts.first()?;
    let second = parts.get(1)?;
    Some(((*first).to_string(), (*second).to_string()))
}

/// Check if a reference is a row reference (all digits)
fn is_row_reference(s: &str) -> bool {
    !s.is_empty() && s.chars().all(|c| c.is_ascii_digit())
}

/// Check if a reference is a column reference (all letters)
fn is_column_reference(s: &str) -> bool {
    !s.is_empty() && s.chars().all(|c| c.is_ascii_alphabetic())
}

/// Parse a row reference like "1" or "3" to 0-indexed row number
fn parse_row(s: &str) -> Option<u32> {
    s.parse::<u32>().ok().map(|r| r.saturating_sub(1))
}

/// Parse a column reference like "A" or "AB" to 0-indexed column number
fn parse_column(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }

    let mut col: u32 = 0;
    for c in s.chars() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            return None;
        }
    }
    Some(col.saturating_sub(1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extract_range_from_formula() {
        assert_eq!(
            extract_range_from_formula("Sheet1!$A$1:$D$10"),
            Some("A1:D10".to_string())
        );
        assert_eq!(
            extract_range_from_formula("'Sheet Name'!$A$1:$Z$100"),
            Some("A1:Z100".to_string())
        );
        assert_eq!(
            extract_range_from_formula("$A$1:$D$10"),
            Some("A1:D10".to_string())
        );
    }

    #[test]
    fn test_parse_range_bounds() {
        assert_eq!(
            parse_range_bounds("$1:$3"),
            Some(("1".to_string(), "3".to_string()))
        );
        assert_eq!(
            parse_range_bounds("$A:$B"),
            Some(("A".to_string(), "B".to_string()))
        );
        assert_eq!(
            parse_range_bounds("A1:D10"),
            Some(("A1".to_string(), "D10".to_string()))
        );
    }

    #[test]
    fn test_is_row_reference() {
        assert!(is_row_reference("1"));
        assert!(is_row_reference("123"));
        assert!(!is_row_reference("A"));
        assert!(!is_row_reference("A1"));
        assert!(!is_row_reference(""));
    }

    #[test]
    fn test_is_column_reference() {
        assert!(is_column_reference("A"));
        assert!(is_column_reference("AB"));
        assert!(!is_column_reference("1"));
        assert!(!is_column_reference("A1"));
        assert!(!is_column_reference(""));
    }

    #[test]
    fn test_parse_row() {
        assert_eq!(parse_row("1"), Some(0));
        assert_eq!(parse_row("3"), Some(2));
        assert_eq!(parse_row("100"), Some(99));
    }

    #[test]
    fn test_parse_column() {
        assert_eq!(parse_column("A"), Some(0));
        assert_eq!(parse_column("B"), Some(1));
        assert_eq!(parse_column("Z"), Some(25));
        assert_eq!(parse_column("AA"), Some(26));
        assert_eq!(parse_column("AB"), Some(27));
    }

    #[test]
    fn test_resolve_relationship_path() {
        assert_eq!(
            resolve_relationship_path("worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_relationship_path("/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_relationship_path("../customXml/item1.xml"),
            "customXml/item1.xml"
        );
    }

    #[test]
    fn test_date_system_default() {
        assert_eq!(DateSystem::default(), DateSystem::Date1900);
    }
}

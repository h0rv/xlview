//! Cell and sheet protection parsing module
//! This module handles parsing of protection settings from XLSX files.

use quick_xml::events::BytesStart;
use std::io::BufRead;

/// Protection settings from xf element
///
/// Represents the cell-level protection attributes that determine
/// whether a cell is locked and/or has its formula hidden.
#[derive(Debug, Clone, Default)]
pub struct CellProtection {
    /// Whether the cell is locked (default true in Excel)
    /// When the sheet is protected, locked cells cannot be edited
    pub locked: bool,
    /// Whether the formula is hidden (default false)
    /// When the sheet is protected, formulas in hidden cells are not displayed
    pub hidden: bool,
}

impl CellProtection {
    /// Create a new CellProtection with Excel default values
    /// (locked=true, hidden=false)
    pub fn new() -> Self {
        Self {
            locked: true,
            hidden: false,
        }
    }
}

/// Sheet protection settings
///
/// Represents the sheetProtection element that defines what operations
/// are restricted when the sheet is protected.
#[derive(Debug, Clone, Default)]
pub struct SheetProtection {
    /// Whether the sheet is protected
    pub sheet: bool,
    /// Whether objects (shapes, charts, etc.) are protected
    pub objects: bool,
    /// Whether scenarios are protected
    pub scenarios: bool,
    /// Whether formatting cells is allowed
    pub format_cells: bool,
    /// Whether formatting columns is allowed
    pub format_columns: bool,
    /// Whether formatting rows is allowed
    pub format_rows: bool,
    /// Whether inserting columns is allowed
    pub insert_columns: bool,
    /// Whether inserting rows is allowed
    pub insert_rows: bool,
    /// Whether inserting hyperlinks is allowed
    pub insert_hyperlinks: bool,
    /// Whether deleting columns is allowed
    pub delete_columns: bool,
    /// Whether deleting rows is allowed
    pub delete_rows: bool,
    /// Whether selecting locked cells is allowed
    pub select_locked_cells: bool,
    /// Whether sorting is allowed
    pub sort: bool,
    /// Whether auto-filter is allowed
    pub auto_filter: bool,
    /// Whether pivot tables are allowed
    pub pivot_tables: bool,
    /// Whether selecting unlocked cells is allowed
    pub select_unlocked_cells: bool,
    /// Password hash (if password protected)
    pub password_hash: Option<String>,
}

/// Parse protection element from xf
///
/// The protection element inside an xf (cell format) element defines
/// cell-level protection settings.
///
/// # Arguments
/// * `e` - The BytesStart event for the protection element
///
/// # Returns
/// A CellProtection struct with parsed settings
///
/// # Example XML
/// ```xml
/// <protection locked="0" hidden="1"/>
/// ```
pub fn parse_cell_protection(e: &BytesStart) -> CellProtection {
    // Excel defaults: locked=true, hidden=false
    let mut protection = CellProtection {
        locked: true,
        hidden: false,
    };

    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"locked" => {
                // locked="0" means unlocked, locked="1" or absent means locked
                protection.locked = std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
            }
            b"hidden" => {
                // hidden="1" means formula is hidden
                protection.hidden = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
            }
            _ => {}
        }
    }

    protection
}

/// Parse sheetProtection element
///
/// The sheetProtection element defines sheet-level protection settings.
/// Most attributes use "0" for not protected/allowed and "1" for protected/restricted.
/// Note: The attribute semantics are inverted for some fields - presence of "1"
/// means the action is PREVENTED, not allowed.
///
/// # Arguments
/// * `e` - The BytesStart event for the sheetProtection element
///
/// # Returns
/// A SheetProtection struct with parsed settings
///
/// # Example XML
/// ```xml
/// <sheetProtection sheet="1" objects="1" scenarios="1"
///     formatCells="0" formatColumns="0" formatRows="0"
///     insertColumns="0" insertRows="0" insertHyperlinks="0"
///     deleteColumns="0" deleteRows="0" selectLockedCells="0"
///     sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="0"
///     password="CC1A"/>
/// ```
pub fn parse_sheet_protection<R: BufRead>(e: &BytesStart) -> SheetProtection {
    let mut protection = SheetProtection::default();

    for attr in e.attributes().flatten() {
        let value_str = std::str::from_utf8(&attr.value).unwrap_or("0");
        let is_true = value_str == "1" || value_str == "true";

        match attr.key.as_ref() {
            b"sheet" => {
                protection.sheet = is_true;
            }
            b"objects" => {
                protection.objects = is_true;
            }
            b"scenarios" => {
                protection.scenarios = is_true;
            }
            b"formatCells" => {
                // When formatCells="0", formatting cells is NOT restricted
                // When formatCells="1", formatting cells IS restricted
                protection.format_cells = is_true;
            }
            b"formatColumns" => {
                protection.format_columns = is_true;
            }
            b"formatRows" => {
                protection.format_rows = is_true;
            }
            b"insertColumns" => {
                protection.insert_columns = is_true;
            }
            b"insertRows" => {
                protection.insert_rows = is_true;
            }
            b"insertHyperlinks" => {
                protection.insert_hyperlinks = is_true;
            }
            b"deleteColumns" => {
                protection.delete_columns = is_true;
            }
            b"deleteRows" => {
                protection.delete_rows = is_true;
            }
            b"selectLockedCells" => {
                protection.select_locked_cells = is_true;
            }
            b"sort" => {
                protection.sort = is_true;
            }
            b"autoFilter" => {
                protection.auto_filter = is_true;
            }
            b"pivotTables" => {
                protection.pivot_tables = is_true;
            }
            b"selectUnlockedCells" => {
                protection.select_unlocked_cells = is_true;
            }
            b"password" => {
                // Password hash (legacy Excel format, 2-byte hash)
                protection.password_hash = Some(value_str.to_string());
            }
            b"hashValue" => {
                // SHA-512 hash value (modern Excel format)
                // Prefer this over legacy password attribute
                protection.password_hash = Some(value_str.to_string());
            }
            b"algorithmName" | b"saltValue" | b"spinCount" => {
                // These are related to the hash algorithm but we just store the hash
                // algorithmName: e.g., "SHA-512"
                // saltValue: base64-encoded salt
                // spinCount: number of iterations
            }
            _ => {}
        }
    }

    protection
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
    use quick_xml::events::Event;
    use quick_xml::Reader;

    #[test]
    fn test_parse_cell_protection_defaults() {
        let xml = r#"<protection/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_cell_protection(&e);
            assert!(protection.locked); // Default true
            assert!(!protection.hidden); // Default false
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_cell_protection_unlocked() {
        let xml = r#"<protection locked="0"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_cell_protection(&e);
            assert!(!protection.locked);
            assert!(!protection.hidden);
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_cell_protection_hidden() {
        let xml = r#"<protection hidden="1"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_cell_protection(&e);
            assert!(protection.locked);
            assert!(protection.hidden);
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_cell_protection_both() {
        let xml = r#"<protection locked="0" hidden="1"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_cell_protection(&e);
            assert!(!protection.locked);
            assert!(protection.hidden);
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_sheet_protection_basic() {
        let xml = r#"<sheetProtection sheet="1" objects="1" scenarios="1"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_sheet_protection::<&[u8]>(&e);
            assert!(protection.sheet);
            assert!(protection.objects);
            assert!(protection.scenarios);
            assert!(!protection.format_cells);
            assert!(protection.password_hash.is_none());
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_sheet_protection_with_password() {
        let xml = r#"<sheetProtection sheet="1" password="CC1A"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_sheet_protection::<&[u8]>(&e);
            assert!(protection.sheet);
            assert_eq!(protection.password_hash, Some("CC1A".to_string()));
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_sheet_protection_all_attrs() {
        let xml = r#"<sheetProtection sheet="1" objects="1" scenarios="1"
            formatCells="1" formatColumns="1" formatRows="1"
            insertColumns="1" insertRows="1" insertHyperlinks="1"
            deleteColumns="1" deleteRows="1" selectLockedCells="1"
            sort="1" autoFilter="1" pivotTables="1" selectUnlockedCells="1"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_sheet_protection::<&[u8]>(&e);
            assert!(protection.sheet);
            assert!(protection.objects);
            assert!(protection.scenarios);
            assert!(protection.format_cells);
            assert!(protection.format_columns);
            assert!(protection.format_rows);
            assert!(protection.insert_columns);
            assert!(protection.insert_rows);
            assert!(protection.insert_hyperlinks);
            assert!(protection.delete_columns);
            assert!(protection.delete_rows);
            assert!(protection.select_locked_cells);
            assert!(protection.sort);
            assert!(protection.auto_filter);
            assert!(protection.pivot_tables);
            assert!(protection.select_unlocked_cells);
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_parse_sheet_protection_modern_hash() {
        let xml = r#"<sheetProtection sheet="1" hashValue="abc123" algorithmName="SHA-512"/>"#;
        let mut reader = Reader::from_str(xml);
        reader.trim_text(true);

        if let Ok(Event::Empty(e)) = reader.read_event() {
            let protection = parse_sheet_protection::<&[u8]>(&e);
            assert!(protection.sheet);
            assert_eq!(protection.password_hash, Some("abc123".to_string()));
        } else {
            panic!("Expected Empty event");
        }
    }

    #[test]
    fn test_cell_protection_new() {
        let protection = CellProtection::new();
        assert!(protection.locked);
        assert!(!protection.hidden);
    }
}

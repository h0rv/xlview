//! XML namespace constants for XLSX parsing
//!
//! XLSX files use XML namespaces extensively. Different Excel versions and tools
//! may use different namespace prefixes or default namespaces. This module provides
//! constants and helper functions for namespace-aware parsing.

use quick_xml::events::BytesStart;

// =============================================================================
// Spreadsheet namespaces
// =============================================================================

/// Main spreadsheet namespace (Transitional conformance)
pub const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Strict OOXML spreadsheet namespace (Office 2013+ Strict conformance)
pub const NS_SPREADSHEET_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

// =============================================================================
// Package namespaces
// =============================================================================

/// Relationships namespace
pub const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Content types namespace
pub const NS_CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

// =============================================================================
// Drawing namespaces
// =============================================================================

/// DrawingML main namespace (for themes, charts, etc.)
pub const NS_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// DrawingML spreadsheet drawing namespace
pub const NS_DRAWING_SPREADSHEET: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

// =============================================================================
// Office document relationship types
// =============================================================================

/// Relationship type for worksheets
pub const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

/// Relationship type for styles
pub const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

/// Relationship type for shared strings
pub const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// Relationship type for theme
pub const REL_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/// Relationship type for workbook (from root .rels)
pub const REL_WORKBOOK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/// Relationship type for comments
pub const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

// =============================================================================
// Strict OOXML relationship types (Office 2013+)
// =============================================================================

/// Strict relationship type for worksheets
pub const REL_WORKSHEET_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";

/// Strict relationship type for styles
pub const REL_STYLES_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";

/// Strict relationship type for shared strings
pub const REL_SHARED_STRINGS_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings";

/// Strict relationship type for theme
pub const REL_THEME_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

/// Strict relationship type for comments
pub const REL_COMMENTS_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";

// =============================================================================
// Helper functions for namespace-aware parsing
// =============================================================================

/// Check if an element matches a local name, ignoring namespace prefix.
///
/// This uses `local_name()` which already strips the namespace prefix,
/// so "a:clrScheme" and "clrScheme" both match when checking for "clrScheme".
///
/// # Example
/// ```ignore
/// if element_matches(&e, b"sheet") {
///     // Matches <sheet>, <x:sheet>, etc.
/// }
/// ```
#[inline]
pub fn element_matches(e: &BytesStart, local_name: &[u8]) -> bool {
    e.local_name().as_ref() == local_name
}

/// Check if an element matches any of the given local names.
///
/// # Example
/// ```ignore
/// if element_matches_any(&e, &[b"dk1", b"lt1", b"dk2", b"lt2"]) {
///     // Matches any theme color element
/// }
/// ```
#[inline]
#[allow(clippy::manual_contains)]
pub fn element_matches_any(e: &BytesStart, names: &[&[u8]]) -> bool {
    let local = e.local_name();
    let local_bytes = local.as_ref();
    names.iter().any(|&name| local_bytes == name)
}

/// Get an attribute value by name, handling both prefixed and unprefixed forms.
///
/// For simple attributes like `name="Sheet1"`, just pass the attribute name.
/// For namespaced attributes like `r:id="rId1"`, this will check both forms.
///
/// # Example
/// ```ignore
/// // Gets "name" attribute
/// let name = get_attribute(&e, b"name");
///
/// // Gets "r:id" attribute (checks both "r:id" and "id" with relationship namespace)
/// let rel_id = get_rel_attribute(&e, b"id");
/// ```
pub fn get_attribute(e: &BytesStart, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == name {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Get an attribute value, accepting multiple possible names.
///
/// Useful for attributes that might appear with different prefixes.
///
/// # Example
/// ```ignore
/// // Gets "id" whether it appears as "id", "r:id", etc.
/// let id = get_attribute_any(&e, &[b"id", b"r:id"]);
/// ```
pub fn get_attribute_any(e: &BytesStart, names: &[&[u8]]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        for &name in names {
            if key == name {
                return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
            }
        }
    }
    None
}

/// Get a relationship ID attribute (commonly `r:id`).
///
/// This handles the common case where relationship IDs can appear as:
/// - `r:id` (most common, with relationship namespace prefix)
/// - `id` (unprefixed)
/// - Other prefixes like `rel:id`
///
/// The function checks for any attribute ending with `:id` or exactly `id`.
pub fn get_rel_id(e: &BytesStart) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        // Check for r:id, rel:id, or any other prefixed form, or plain "id"
        if key == b"id" || key == b"r:id" || (key.len() > 3 && key.ends_with(b":id")) {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Check if a relationship type matches a known type, handling both
/// transitional and strict OOXML variants.
///
/// # Example
/// ```ignore
/// if is_worksheet_relationship(&rel_type) {
///     // Handle worksheet relationship
/// }
/// ```
pub fn is_worksheet_relationship(rel_type: &str) -> bool {
    rel_type == REL_WORKSHEET || rel_type == REL_WORKSHEET_STRICT || rel_type.contains("worksheet")
}

/// Check if a relationship type is for styles.
pub fn is_styles_relationship(rel_type: &str) -> bool {
    rel_type == REL_STYLES || rel_type == REL_STYLES_STRICT || rel_type.contains("/styles")
}

/// Check if a relationship type is for shared strings.
pub fn is_shared_strings_relationship(rel_type: &str) -> bool {
    rel_type == REL_SHARED_STRINGS
        || rel_type == REL_SHARED_STRINGS_STRICT
        || rel_type.contains("sharedStrings")
}

/// Check if a relationship type is for theme.
pub fn is_theme_relationship(rel_type: &str) -> bool {
    rel_type == REL_THEME || rel_type == REL_THEME_STRICT || rel_type.contains("/theme")
}

/// Check if a relationship type is for comments.
pub fn is_comments_relationship(rel_type: &str) -> bool {
    rel_type == REL_COMMENTS || rel_type == REL_COMMENTS_STRICT || rel_type.contains("/comments")
}

/// Parse a local element name from a BytesStart, returning it as an owned String.
///
/// This is a convenience wrapper that handles the conversion from bytes to string.
#[inline]
pub fn local_name_str(e: &BytesStart) -> String {
    let local = e.local_name();
    std::str::from_utf8(local.as_ref())
        .unwrap_or("")
        .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_relationship_type_matching() {
        // Transitional conformance
        assert!(is_worksheet_relationship(REL_WORKSHEET));
        assert!(is_styles_relationship(REL_STYLES));
        assert!(is_shared_strings_relationship(REL_SHARED_STRINGS));
        assert!(is_theme_relationship(REL_THEME));

        // Strict conformance
        assert!(is_worksheet_relationship(REL_WORKSHEET_STRICT));
        assert!(is_styles_relationship(REL_STYLES_STRICT));
        assert!(is_shared_strings_relationship(REL_SHARED_STRINGS_STRICT));
        assert!(is_theme_relationship(REL_THEME_STRICT));

        // Substring matching for edge cases
        assert!(is_worksheet_relationship(
            "http://example.com/relationships/worksheet"
        ));
    }
}

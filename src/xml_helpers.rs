//! Shared XML attribute parsing utilities for XLSX parsing.
//!
//! These helpers eliminate duplicated attribute extraction boilerplate
//! across the 15+ parser modules. All functions handle namespace-prefixed
//! attributes and UTF-8 conversion safely.

use quick_xml::events::BytesStart;

use crate::types::ColorSpec;

/// Extract a string attribute value by key.
///
/// Returns `None` if the attribute is missing or not valid UTF-8.
pub fn attr_string(e: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Extract a string attribute by local name (ignoring namespace prefix).
pub fn attr_string_local(e: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() == key {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Extract a `u32` attribute value by key.
pub fn attr_u32(e: &BytesStart, key: &[u8]) -> Option<u32> {
    attr_string(e, key).and_then(|s| s.parse().ok())
}

/// Extract an `i32` attribute value by key.
pub fn attr_i32(e: &BytesStart, key: &[u8]) -> Option<i32> {
    attr_string(e, key).and_then(|s| s.parse().ok())
}

/// Extract an `i64` attribute value by key.
pub fn attr_i64(e: &BytesStart, key: &[u8]) -> Option<i64> {
    attr_string(e, key).and_then(|s| s.parse().ok())
}

/// Extract an `f64` attribute value by key.
pub fn attr_f64(e: &BytesStart, key: &[u8]) -> Option<f64> {
    attr_string(e, key).and_then(|s| s.parse().ok())
}

/// Extract a boolean attribute value by key.
///
/// Returns `None` if missing. Recognizes `"1"`, `"true"` as true; `"0"`, `"false"` as false.
pub fn attr_bool(e: &BytesStart, key: &[u8]) -> Option<bool> {
    attr_string(e, key).map(|s| matches!(s.as_str(), "1" | "true"))
}

/// Extract a boolean attribute with a default value.
///
/// Returns the default if the attribute is missing.
pub fn attr_bool_default(e: &BytesStart, key: &[u8], default: bool) -> bool {
    attr_bool(e, key).unwrap_or(default)
}

/// Extract the `val` attribute as a string. Very common in XLSX XML.
pub fn attr_val(e: &BytesStart) -> Option<String> {
    attr_string(e, b"val")
}

/// Extract the `val` attribute as `u32`.
pub fn attr_val_u32(e: &BytesStart) -> Option<u32> {
    attr_u32(e, b"val")
}

/// Extract the `val` attribute as `f64`.
pub fn attr_val_f64(e: &BytesStart) -> Option<f64> {
    attr_f64(e, b"val")
}

/// Parse color attributes from an XML element into a `ColorSpec`.
///
/// Handles `rgb`, `theme`, `tint`, `indexed`, and `auto` attributes.
/// This replaces 5+ duplicate `parse_color_attrs` implementations across parser modules.
pub fn parse_color_attrs(e: &BytesStart) -> ColorSpec {
    ColorSpec {
        rgb: attr_string(e, b"rgb"),
        theme: attr_u32(e, b"theme"),
        tint: attr_f64(e, b"tint"),
        indexed: attr_u32(e, b"indexed"),
        auto: attr_bool_default(e, b"auto", false),
    }
}

/// Get the local element name as an owned string.
///
/// Returns empty string if not valid UTF-8.
#[inline]
pub fn local_name_string(e: &BytesStart) -> String {
    let bytes = e.local_name();
    std::str::from_utf8(bytes.as_ref())
        .unwrap_or("")
        .to_string()
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::approx_constant,
    clippy::panic
)]
mod tests {
    use super::*;

    fn make_start(xml: &str) -> BytesStart<'_> {
        // Strip < and > / /> to get just the tag content
        let content = xml
            .trim_start_matches('<')
            .trim_end_matches('/')
            .trim_end_matches('>');
        BytesStart::from_content(content, content.find(' ').unwrap_or(content.len()))
    }

    #[test]
    fn test_attr_string() {
        let e = make_start(r#"<foo name="hello" />"#);
        assert_eq!(attr_string(&e, b"name"), Some("hello".to_string()));
        assert_eq!(attr_string(&e, b"missing"), None);
    }

    #[test]
    fn test_attr_u32() {
        let e = make_start(r#"<foo count="42" />"#);
        assert_eq!(attr_u32(&e, b"count"), Some(42));
        assert_eq!(attr_u32(&e, b"missing"), None);
    }

    #[test]
    fn test_attr_f64() {
        let e = make_start(r#"<foo val="3.14" />"#);
        let v = attr_f64(&e, b"val");
        assert!(v.is_some());
        let diff = v.unwrap_or(0.0) - 3.14;
        assert!(diff.abs() < f64::EPSILON);
    }

    #[test]
    fn test_attr_bool() {
        let e = make_start(r#"<foo a="1" b="0" c="true" d="false" />"#);
        assert_eq!(attr_bool(&e, b"a"), Some(true));
        assert_eq!(attr_bool(&e, b"b"), Some(false));
        assert_eq!(attr_bool(&e, b"c"), Some(true));
        assert_eq!(attr_bool(&e, b"d"), Some(false));
        assert_eq!(attr_bool(&e, b"missing"), None);
    }

    #[test]
    fn test_attr_bool_default() {
        let e = make_start(r#"<foo a="1" />"#);
        assert!(attr_bool_default(&e, b"a", false));
        assert!(!attr_bool_default(&e, b"missing", false));
        assert!(attr_bool_default(&e, b"missing", true));
    }

    #[test]
    fn test_parse_color_attrs() {
        let e = make_start(r#"<color rgb="FF0000" theme="1" tint="0.5" />"#);
        let color = parse_color_attrs(&e);
        assert_eq!(color.rgb, Some("FF0000".to_string()));
        assert_eq!(color.theme, Some(1));
        let tint_diff = color.tint.unwrap_or(0.0) - 0.5;
        assert!(tint_diff.abs() < f64::EPSILON);
        assert_eq!(color.indexed, None);
        assert!(!color.auto);
    }

    #[test]
    fn test_attr_val() {
        let e = make_start(r#"<sz val="12" />"#);
        assert_eq!(attr_val(&e), Some("12".to_string()));
        assert_eq!(attr_val_u32(&e), Some(12));
    }
}

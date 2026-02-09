//! Font styling tests for xlview
//!
//! Tests parsing of font properties from styles.xml
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

use std::io::Cursor;

// Import from the library
use xlview::color::{resolve_color, INDEXED_COLORS};
use xlview::styles::parse_styles;
use xlview::types::ColorSpec;

// ============================================================================
// Font Family Tests
// ============================================================================

#[test]
fn test_font_family_arial() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts.len(), 1);
    assert_eq!(stylesheet.fonts[0].name, Some("Arial".to_string()));
}

#[test]
fn test_font_family_calibri() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].name, Some("Calibri".to_string()));
}

#[test]
fn test_font_family_times_new_roman() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Times New Roman"/>
      <sz val="12"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(
        stylesheet.fonts[0].name,
        Some("Times New Roman".to_string())
    );
}

#[test]
fn test_font_family_courier_new() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Courier New"/>
      <sz val="10"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].name, Some("Courier New".to_string()));
}

#[test]
fn test_multiple_font_families() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
    </font>
    <font>
      <name val="Verdana"/>
      <sz val="10"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts.len(), 3);
    assert_eq!(stylesheet.fonts[0].name, Some("Arial".to_string()));
    assert_eq!(stylesheet.fonts[1].name, Some("Calibri".to_string()));
    assert_eq!(stylesheet.fonts[2].name, Some("Verdana".to_string()));
}

// ============================================================================
// Font Size Tests
// ============================================================================

#[test]
fn test_font_size_8() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="8"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(8.0));
}

#[test]
fn test_font_size_10() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="10"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(10.0));
}

#[test]
fn test_font_size_11() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(11.0));
}

#[test]
fn test_font_size_12() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="12"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(12.0));
}

#[test]
fn test_font_size_14() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="14"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(14.0));
}

#[test]
fn test_font_size_18() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="18"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(18.0));
}

#[test]
fn test_font_size_24() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="24"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(24.0));
}

#[test]
fn test_font_size_36() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="36"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(36.0));
}

#[test]
fn test_font_size_72() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="72"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(72.0));
}

#[test]
fn test_font_size_decimal() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="10.5"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].size, Some(10.5));
}

// ============================================================================
// Font Color RGB Tests
// ============================================================================

#[test]
fn test_font_color_rgb_red() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FFFF0000"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].color.is_some());
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.rgb, Some("FFFF0000".to_string()));

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#FF0000".to_string()));
}

#[test]
fn test_font_color_rgb_green() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FF00FF00"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.rgb, Some("FF00FF00".to_string()));

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#00FF00".to_string()));
}

#[test]
fn test_font_color_rgb_blue() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FF0000FF"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.rgb, Some("FF0000FF".to_string()));

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#0000FF".to_string()));
}

#[test]
fn test_font_color_rgb_black() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FF000000"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#000000".to_string()));
}

#[test]
fn test_font_color_rgb_white() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#FFFFFF".to_string()));
}

#[test]
fn test_font_color_rgb_custom() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color rgb="FF4472C4"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#4472C4".to_string()));
}

// ============================================================================
// Font Color Theme Tests
// ============================================================================

fn get_default_theme_colors() -> Vec<String> {
    vec![
        "#000000".to_string(), // 0: dk1
        "#FFFFFF".to_string(), // 1: lt1
        "#44546A".to_string(), // 2: dk2
        "#E7E6E6".to_string(), // 3: lt2
        "#4472C4".to_string(), // 4: accent1
        "#ED7D31".to_string(), // 5: accent2
        "#A5A5A5".to_string(), // 6: accent3
        "#FFC000".to_string(), // 7: accent4
        "#5B9BD5".to_string(), // 8: accent5
        "#70AD47".to_string(), // 9: accent6
        "#0563C1".to_string(), // 10: hlink
        "#954F72".to_string(), // 11: folHlink
    ]
}

#[test]
fn test_font_color_theme_0_dark1() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="0"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(0));

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#000000".to_string()));
}

#[test]
fn test_font_color_theme_1_light1() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="1"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(1));

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#FFFFFF".to_string()));
}

#[test]
fn test_font_color_theme_4_accent1() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="4"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(4));

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#4472C4".to_string()));
}

#[test]
fn test_font_color_theme_5_accent2() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="5"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#ED7D31".to_string()));
}

#[test]
fn test_font_color_theme_10_hyperlink() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="10"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#0563C1".to_string()));
}

#[test]
fn test_font_color_theme_11_followed_hyperlink() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#954F72".to_string()));
}

// ============================================================================
// Font Color Indexed Tests
// ============================================================================

#[test]
fn test_font_color_indexed_8_black() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="8"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.indexed, Some(8));

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some(INDEXED_COLORS[8].to_string()));
}

#[test]
fn test_font_color_indexed_9_white() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="9"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.indexed, Some(9));

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some(INDEXED_COLORS[9].to_string()));
}

#[test]
fn test_font_color_indexed_10_red() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="10"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some(INDEXED_COLORS[10].to_string()));
}

#[test]
fn test_font_color_indexed_30() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="30"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some(INDEXED_COLORS[30].to_string()));
}

#[test]
fn test_font_color_indexed_63() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="63"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some(INDEXED_COLORS[63].to_string()));
}

#[test]
fn test_font_color_indexed_64_system_foreground() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color indexed="64"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();

    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#000000".to_string()));
}

// ============================================================================
// Font Color with Tint Tests
// ============================================================================

#[test]
fn test_font_color_theme_with_positive_tint() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="0" tint="0.5"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(0));
    assert_eq!(color.tint, Some(0.5));

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    // Black (#000000) with 0.5 tint should lighten to gray
    assert!(resolved.is_some());
    assert_eq!(resolved, Some("#808080".to_string()));
}

#[test]
fn test_font_color_theme_with_negative_tint() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="1" tint="-0.5"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(1));
    assert_eq!(color.tint, Some(-0.5));

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    // White (#FFFFFF) with -0.5 tint should darken to gray
    assert!(resolved.is_some());
    assert_eq!(resolved, Some("#808080".to_string()));
}

#[test]
fn test_font_color_theme_with_small_positive_tint() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="4" tint="0.39997558519241921"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(4));
    assert!(color.tint.is_some());

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    // Should be a lighter version of accent1
    assert!(resolved.is_some());
}

#[test]
fn test_font_color_theme_with_small_negative_tint() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <color theme="4" tint="-0.249977111117893"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    assert_eq!(color.theme, Some(4));
    assert!(color.tint.is_some());

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    // Should be a darker version of accent1
    assert!(resolved.is_some());
}

// ============================================================================
// Bold Tests
// ============================================================================

#[test]
fn test_font_bold() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <b/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].bold);
}

#[test]
fn test_font_not_bold() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(!stylesheet.fonts[0].bold);
}

#[test]
fn test_font_bold_with_val_true() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <b val="true"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].bold);
}

// ============================================================================
// Italic Tests
// ============================================================================

#[test]
fn test_font_italic() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <i/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].italic);
}

#[test]
fn test_font_not_italic() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(!stylesheet.fonts[0].italic);
}

#[test]
fn test_font_italic_with_val() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <i val="1"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].italic);
}

// ============================================================================
// Underline Single Tests
// ============================================================================

#[test]
fn test_font_underline_single_empty_tag() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <u/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].underline.is_some());
}

#[test]
fn test_font_underline_single_explicit() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <u val="single"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].underline.is_some());
}

#[test]
fn test_font_no_underline() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].underline.is_none());
}

// ============================================================================
// Underline Double Tests
// ============================================================================

#[test]
fn test_font_underline_double() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <u val="double"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    // The current implementation treats any <u> tag as underline=true
    assert!(stylesheet.fonts[0].underline.is_some());
}

// ============================================================================
// Strikethrough Tests
// ============================================================================

#[test]
fn test_font_strikethrough() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <strike/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].strikethrough);
}

#[test]
fn test_font_no_strikethrough() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(!stylesheet.fonts[0].strikethrough);
}

#[test]
fn test_font_strikethrough_with_val() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <strike val="true"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].strikethrough);
}

// ============================================================================
// Combination Tests
// ============================================================================

#[test]
fn test_font_bold_and_italic() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="11"/>
      <b/>
      <i/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].bold);
    assert!(stylesheet.fonts[0].italic);
    assert!(stylesheet.fonts[0].underline.is_none());
    assert!(!stylesheet.fonts[0].strikethrough);
}

#[test]
fn test_font_bold_underline_color() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
      <sz val="12"/>
      <b/>
      <u/>
      <color rgb="FFFF0000"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert!(stylesheet.fonts[0].bold);
    assert!(stylesheet.fonts[0].underline.is_some());
    assert!(!stylesheet.fonts[0].italic);

    let color = stylesheet.fonts[0].color.as_ref().unwrap();
    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#FF0000".to_string()));
}

#[test]
fn test_font_all_styles_combined() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Times New Roman"/>
      <sz val="14"/>
      <b/>
      <i/>
      <u/>
      <strike/>
      <color rgb="FF0000FF"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let font = &stylesheet.fonts[0];

    assert_eq!(font.name, Some("Times New Roman".to_string()));
    assert_eq!(font.size, Some(14.0));
    assert!(font.bold);
    assert!(font.italic);
    assert!(font.underline.is_some());
    assert!(font.strikethrough);

    let color = font.color.as_ref().unwrap();
    let resolved = resolve_color(color, &[], None);
    assert_eq!(resolved, Some("#0000FF".to_string()));
}

#[test]
fn test_font_italic_strikethrough_theme_color() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
      <i/>
      <strike/>
      <color theme="5"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    let font = &stylesheet.fonts[0];

    assert!(!font.bold);
    assert!(font.italic);
    assert!(font.underline.is_none());
    assert!(font.strikethrough);

    let color = font.color.as_ref().unwrap();
    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(color, &theme_colors, None);
    assert_eq!(resolved, Some("#ED7D31".to_string()));
}

#[test]
fn test_multiple_fonts_with_different_styles() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
    </font>
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
      <b/>
    </font>
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
      <i/>
      <color rgb="FFFF0000"/>
    </font>
    <font>
      <name val="Arial"/>
      <sz val="14"/>
      <b/>
      <u/>
      <color theme="4"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts.len(), 4);

    // Font 0: plain
    assert!(!stylesheet.fonts[0].bold);
    assert!(!stylesheet.fonts[0].italic);
    assert!(stylesheet.fonts[0].color.is_none());

    // Font 1: bold only
    assert!(stylesheet.fonts[1].bold);
    assert!(!stylesheet.fonts[1].italic);

    // Font 2: italic with red color
    assert!(!stylesheet.fonts[2].bold);
    assert!(stylesheet.fonts[2].italic);
    let color2 = stylesheet.fonts[2].color.as_ref().unwrap();
    assert_eq!(color2.rgb, Some("FFFF0000".to_string()));

    // Font 3: bold, underline, theme color
    assert!(stylesheet.fonts[3].bold);
    assert!(stylesheet.fonts[3].underline.is_some());
    let color3 = stylesheet.fonts[3].color.as_ref().unwrap();
    assert_eq!(color3.theme, Some(4));
}

// ============================================================================
// CellXf with Font Reference Tests
// ============================================================================

#[test]
fn test_cellxf_references_font() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <name val="Calibri"/>
      <sz val="11"/>
    </font>
    <font>
      <name val="Arial"/>
      <sz val="14"/>
      <b/>
      <color rgb="FFFF0000"/>
    </font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellXfs count="2">
    <xf fontId="0" fillId="0" borderId="0"/>
    <xf fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.cell_xfs.len(), 2);
    assert_eq!(stylesheet.cell_xfs[0].font_id, Some(0));
    assert_eq!(stylesheet.cell_xfs[1].font_id, Some(1));
    assert!(stylesheet.cell_xfs[1].apply_font);

    // Verify the referenced font
    let font = &stylesheet.fonts[1];
    assert_eq!(font.name, Some("Arial".to_string()));
    assert_eq!(font.size, Some(14.0));
    assert!(font.bold);
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_font_with_empty_name() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val=""/>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].name, Some("".to_string()));
}

#[test]
fn test_font_missing_size() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <name val="Arial"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].name, Some("Arial".to_string()));
    assert_eq!(stylesheet.fonts[0].size, None);
}

#[test]
fn test_font_missing_name() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
    </font>
  </fonts>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();
    assert_eq!(stylesheet.fonts[0].name, None);
    assert_eq!(stylesheet.fonts[0].size, Some(11.0));
}

#[test]
fn test_color_auto() {
    let color = ColorSpec {
        rgb: None,
        theme: None,
        tint: None,
        indexed: None,
        auto: true,
    };

    let resolved = resolve_color(&color, &[], None);
    assert_eq!(resolved, Some("#000000".to_string()));
}

#[test]
fn test_color_priority_rgb_over_theme() {
    let color = ColorSpec {
        rgb: Some("FF123456".to_string()),
        theme: Some(4),
        tint: None,
        indexed: None,
        auto: false,
    };

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(&color, &theme_colors, None);
    // RGB should take priority
    assert_eq!(resolved, Some("#123456".to_string()));
}

#[test]
fn test_color_priority_theme_over_indexed() {
    let color = ColorSpec {
        rgb: None,
        theme: Some(4),
        tint: None,
        indexed: Some(10),
        auto: false,
    };

    let theme_colors = get_default_theme_colors();
    let resolved = resolve_color(&color, &theme_colors, None);
    // Theme should take priority over indexed
    assert_eq!(resolved, Some("#4472C4".to_string()));
}

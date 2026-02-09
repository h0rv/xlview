//! Comprehensive tests for text alignment styling in xlview
//!
//! These tests verify that alignment properties from xl/styles.xml are correctly
//! parsed into the `RawAlignment` and resolved into the `Style` struct.
//!
//! XLSX alignment element format:
//! ```xml
//! <xf ...>
//!   <alignment horizontal="center" vertical="center" wrapText="1" textRotation="45" indent="1"/>
//! </xf>
//! ```
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
use xlview::styles::parse_styles;
use xlview::types::{RawAlignment, StyleSheet};

/// Helper function to create a minimal styles.xml with the given alignment attributes
fn create_styles_xml(alignment_attrs: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/>
    </border>
  </borders>
  <cellXfs count="1">
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment {alignment_attrs}/>
    </xf>
  </cellXfs>
</styleSheet>"#
    )
}

/// Helper function to create styles.xml with an xf that has no alignment element
fn create_styles_xml_no_alignment() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/>
    </border>
  </borders>
  <cellXfs count="1">
    <xf fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#
        .to_string()
}

/// Helper to parse styles XML and get the first xf's alignment
fn parse_alignment(xml: &str) -> Option<RawAlignment> {
    let cursor = Cursor::new(xml.as_bytes());
    let stylesheet = parse_styles(cursor).expect("Failed to parse styles XML");
    stylesheet
        .cell_xfs
        .first()
        .and_then(|xf| xf.alignment.clone())
}

/// Helper to parse and return the full stylesheet
fn parse_stylesheet(xml: &str) -> StyleSheet {
    let cursor = Cursor::new(xml.as_bytes());
    parse_styles(cursor).expect("Failed to parse styles XML")
}

// ============================================================================
// HORIZONTAL ALIGNMENT TESTS
// ============================================================================

/// Test 1: horizontal="general" - Default, context-dependent alignment
#[test]
fn test_horizontal_alignment_general() {
    let xml = create_styles_xml(r#"horizontal="general""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("general".to_string()),
        "horizontal='general' should be parsed as 'general'"
    );
}

/// Test 2: horizontal="left" - Left-aligned text
#[test]
fn test_horizontal_alignment_left() {
    let xml = create_styles_xml(r#"horizontal="left""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("left".to_string()),
        "horizontal='left' should be parsed as 'left'"
    );
}

/// Test 3: horizontal="center" - Center-aligned text
#[test]
fn test_horizontal_alignment_center() {
    let xml = create_styles_xml(r#"horizontal="center""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("center".to_string()),
        "horizontal='center' should be parsed as 'center'"
    );
}

/// Test 4: horizontal="right" - Right-aligned text
#[test]
fn test_horizontal_alignment_right() {
    let xml = create_styles_xml(r#"horizontal="right""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("right".to_string()),
        "horizontal='right' should be parsed as 'right'"
    );
}

/// Test 5: horizontal="fill" - Text repeats to fill cell width
#[test]
fn test_horizontal_alignment_fill() {
    let xml = create_styles_xml(r#"horizontal="fill""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("fill".to_string()),
        "horizontal='fill' should be parsed as 'fill'"
    );
}

/// Test 6: horizontal="justify" - Text justified across cell width
#[test]
fn test_horizontal_alignment_justify() {
    let xml = create_styles_xml(r#"horizontal="justify""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("justify".to_string()),
        "horizontal='justify' should be parsed as 'justify'"
    );
}

/// Test 7: horizontal="centerContinuous" - Center across selection without merge
#[test]
fn test_horizontal_alignment_center_continuous() {
    let xml = create_styles_xml(r#"horizontal="centerContinuous""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("centerContinuous".to_string()),
        "horizontal='centerContinuous' should be parsed as 'centerContinuous'"
    );
}

/// Test 8: horizontal="distributed" - Text distributed evenly across cell
#[test]
fn test_horizontal_alignment_distributed() {
    let xml = create_styles_xml(r#"horizontal="distributed""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("distributed".to_string()),
        "horizontal='distributed' should be parsed as 'distributed'"
    );
}

// ============================================================================
// VERTICAL ALIGNMENT TESTS
// ============================================================================

/// Test 9: vertical="top" - Top-aligned text
#[test]
fn test_vertical_alignment_top() {
    let xml = create_styles_xml(r#"vertical="top""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.vertical,
        Some("top".to_string()),
        "vertical='top' should be parsed as 'top'"
    );
}

/// Test 10: vertical="center" - Vertically centered text
#[test]
fn test_vertical_alignment_center() {
    let xml = create_styles_xml(r#"vertical="center""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.vertical,
        Some("center".to_string()),
        "vertical='center' should be parsed as 'center'"
    );
}

/// Test 11: vertical="bottom" - Bottom-aligned text (often default)
#[test]
fn test_vertical_alignment_bottom() {
    let xml = create_styles_xml(r#"vertical="bottom""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.vertical,
        Some("bottom".to_string()),
        "vertical='bottom' should be parsed as 'bottom'"
    );
}

/// Test 12: vertical="justify" - Text justified vertically
#[test]
fn test_vertical_alignment_justify() {
    let xml = create_styles_xml(r#"vertical="justify""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.vertical,
        Some("justify".to_string()),
        "vertical='justify' should be parsed as 'justify'"
    );
}

/// Test 13: vertical="distributed" - Text distributed vertically
#[test]
fn test_vertical_alignment_distributed() {
    let xml = create_styles_xml(r#"vertical="distributed""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.vertical,
        Some("distributed".to_string()),
        "vertical='distributed' should be parsed as 'distributed'"
    );
}

// ============================================================================
// TEXT CONTROL TESTS
// ============================================================================

/// Test 14: wrapText="1" - Text wraps within cell
#[test]
fn test_wrap_text_enabled() {
    let xml = create_styles_xml(r#"wrapText="1""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert!(
        alignment.wrap_text,
        "wrapText='1' should set wrap_text to true"
    );
}

/// Test 14b: wrapText="0" - Text does not wrap
#[test]
fn test_wrap_text_disabled() {
    let xml = create_styles_xml(r#"wrapText="0""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert!(
        !alignment.wrap_text,
        "wrapText='0' should set wrap_text to false"
    );
}

/// Test 14c: wrapText absent - Default is no wrap
#[test]
fn test_wrap_text_absent() {
    let xml = create_styles_xml(r#"horizontal="left""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert!(
        !alignment.wrap_text,
        "Absent wrapText should default to false"
    );
}

/// Test 15: shrinkToFit="1" - Font shrinks to fit cell width
/// Note: This test documents expected behavior - shrinkToFit parsing may need implementation
#[test]
fn test_shrink_to_fit_enabled() {
    let xml = create_styles_xml(r#"shrinkToFit="1""#);
    let alignment = parse_alignment(&xml);

    // shrinkToFit should be parsed if the field exists in RawAlignment
    // This test documents the expected attribute parsing
    assert!(
        alignment.is_some(),
        "Alignment element with shrinkToFit should be parsed"
    );

    // TODO: When shrinkToFit is added to RawAlignment, verify:
    // assert!(alignment.unwrap().shrink_to_fit, "shrinkToFit='1' should set shrink_to_fit to true");
}

/// Test 16: wrapText + shrinkToFit together - wrapText takes precedence
/// When both are set, wrapText wins according to Excel/OOXML spec
#[test]
fn test_wrap_text_and_shrink_to_fit_combination() {
    let xml = create_styles_xml(r#"wrapText="1" shrinkToFit="1""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    // wrapText should be honored
    assert!(
        alignment.wrap_text,
        "wrapText='1' should be true even with shrinkToFit"
    );

    // According to OOXML spec, when both are true, wrapText takes precedence
    // shrinkToFit is effectively ignored when wrapText is enabled
}

// ============================================================================
// TEXT ROTATION TESTS
// ============================================================================

/// Test 17: textRotation="45" - 45 degrees counterclockwise
#[test]
fn test_text_rotation_45_degrees() {
    let xml = create_styles_xml(r#"textRotation="45""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(45),
        "textRotation='45' should be parsed as 45"
    );
}

/// Test 18: textRotation="90" - Vertical text, bottom to top
#[test]
fn test_text_rotation_90_degrees() {
    let xml = create_styles_xml(r#"textRotation="90""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(90),
        "textRotation='90' should be parsed as 90"
    );
}

/// Test 19: textRotation="135" - 45 degrees clockwise (stored as 90+45)
/// In OOXML, clockwise rotation is stored as 90 + degrees
/// So 45 degrees clockwise = 135 in the XML
#[test]
fn test_text_rotation_negative_45_degrees() {
    let xml = create_styles_xml(r#"textRotation="135""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(135),
        "textRotation='135' (45 degrees clockwise) should be parsed as 135"
    );
}

/// Test 19b: textRotation="180" - 90 degrees clockwise
#[test]
fn test_text_rotation_negative_90_degrees() {
    let xml = create_styles_xml(r#"textRotation="180""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(180),
        "textRotation='180' (90 degrees clockwise) should be parsed as 180"
    );
}

/// Test 20: textRotation="255" - Special value for stacked vertical text
/// This is a special OOXML value meaning each character is stacked vertically
#[test]
fn test_text_rotation_vertical_stacked() {
    let xml = create_styles_xml(r#"textRotation="255""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(255),
        "textRotation='255' (vertical stacked) should be parsed as 255"
    );
}

/// Test 21: textRotation="0" - No rotation (horizontal)
#[test]
fn test_text_rotation_zero() {
    let xml = create_styles_xml(r#"textRotation="0""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(0),
        "textRotation='0' should be parsed as 0"
    );
}

/// Test 21b: textRotation absent - Should be None
#[test]
fn test_text_rotation_absent() {
    let xml = create_styles_xml(r#"horizontal="left""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation, None,
        "Absent textRotation should be None"
    );
}

/// Test boundary: textRotation at maximum counterclockwise (90)
#[test]
fn test_text_rotation_max_counterclockwise() {
    let xml = create_styles_xml(r#"textRotation="90""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.text_rotation,
        Some(90),
        "Maximum counterclockwise rotation should be 90"
    );
}

// ============================================================================
// INDENT TESTS
// ============================================================================

/// Test 22: indent="1" - Single level indent
#[test]
fn test_indent_level_1() {
    let xml = create_styles_xml(r#"indent="1""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.indent,
        Some(1),
        "indent='1' should be parsed as 1"
    );
}

/// Test 23: indent="2" - Two level indent
#[test]
fn test_indent_level_2() {
    let xml = create_styles_xml(r#"indent="2""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.indent,
        Some(2),
        "indent='2' should be parsed as 2"
    );
}

/// Test 24: indent="5" - Five level indent
#[test]
fn test_indent_level_5() {
    let xml = create_styles_xml(r#"indent="5""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.indent,
        Some(5),
        "indent='5' should be parsed as 5"
    );
}

/// Test 25: indent with horizontal alignment
#[test]
fn test_indent_with_left_alignment() {
    let xml = create_styles_xml(r#"horizontal="left" indent="3""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("left".to_string()),
        "Should have left alignment"
    );
    assert_eq!(alignment.indent, Some(3), "Should have indent of 3");
}

/// Test indent with right alignment (indents from right edge)
#[test]
fn test_indent_with_right_alignment() {
    let xml = create_styles_xml(r#"horizontal="right" indent="2""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("right".to_string()),
        "Should have right alignment"
    );
    assert_eq!(alignment.indent, Some(2), "Should have indent of 2");
}

/// Test indent="0" - No indent
#[test]
fn test_indent_zero() {
    let xml = create_styles_xml(r#"indent="0""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.indent,
        Some(0),
        "indent='0' should be parsed as 0"
    );
}

/// Test indent absent - Should be None
#[test]
fn test_indent_absent() {
    let xml = create_styles_xml(r#"horizontal="left""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(alignment.indent, None, "Absent indent should be None");
}

/// Test large indent value
#[test]
fn test_indent_large_value() {
    let xml = create_styles_xml(r#"indent="15""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.indent,
        Some(15),
        "indent='15' should be parsed as 15"
    );
}

// ============================================================================
// READING ORDER TESTS
// ============================================================================

/// Test 26: readingOrder="0" - Context-dependent reading order
/// Note: readingOrder parsing may need to be implemented
#[test]
fn test_reading_order_context() {
    let xml = create_styles_xml(r#"readingOrder="0""#);
    let alignment = parse_alignment(&xml);

    // readingOrder should be parsed if the field exists in RawAlignment
    assert!(
        alignment.is_some(),
        "Alignment element with readingOrder should be parsed"
    );

    // TODO: When readingOrder is added to RawAlignment, verify:
    // assert_eq!(alignment.unwrap().reading_order, Some(0));
}

/// Test 27: readingOrder="1" - Left to Right reading order
#[test]
fn test_reading_order_ltr() {
    let xml = create_styles_xml(r#"readingOrder="1""#);
    let alignment = parse_alignment(&xml);

    assert!(
        alignment.is_some(),
        "Alignment element with readingOrder='1' should be parsed"
    );

    // TODO: When readingOrder is added to RawAlignment, verify:
    // assert_eq!(alignment.unwrap().reading_order, Some(1));
}

/// Test 28: readingOrder="2" - Right to Left reading order
#[test]
fn test_reading_order_rtl() {
    let xml = create_styles_xml(r#"readingOrder="2""#);
    let alignment = parse_alignment(&xml);

    assert!(
        alignment.is_some(),
        "Alignment element with readingOrder='2' should be parsed"
    );

    // TODO: When readingOrder is added to RawAlignment, verify:
    // assert_eq!(alignment.unwrap().reading_order, Some(2));
}

// ============================================================================
// COMBINATION TESTS
// ============================================================================

/// Test 29: Center + middle + wrap - Common table header combination
#[test]
fn test_combination_center_middle_wrap() {
    let xml = create_styles_xml(r#"horizontal="center" vertical="center" wrapText="1""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("center".to_string()),
        "Should have center horizontal alignment"
    );
    assert_eq!(
        alignment.vertical,
        Some("center".to_string()),
        "Should have center vertical alignment"
    );
    assert!(alignment.wrap_text, "Should have wrap text enabled");
}

/// Test 30: Right + bottom + rotation - Complex combination
#[test]
fn test_combination_right_bottom_rotation() {
    let xml = create_styles_xml(r#"horizontal="right" vertical="bottom" textRotation="45""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("right".to_string()),
        "Should have right horizontal alignment"
    );
    assert_eq!(
        alignment.vertical,
        Some("bottom".to_string()),
        "Should have bottom vertical alignment"
    );
    assert_eq!(
        alignment.text_rotation,
        Some(45),
        "Should have 45 degree rotation"
    );
}

/// Test all alignment properties together
#[test]
fn test_combination_all_properties() {
    let xml = create_styles_xml(
        r#"horizontal="left" vertical="top" wrapText="1" textRotation="30" indent="2""#,
    );
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("left".to_string()),
        "Should have left horizontal alignment"
    );
    assert_eq!(
        alignment.vertical,
        Some("top".to_string()),
        "Should have top vertical alignment"
    );
    assert!(alignment.wrap_text, "Should have wrap text enabled");
    assert_eq!(
        alignment.text_rotation,
        Some(30),
        "Should have 30 degree rotation"
    );
    assert_eq!(alignment.indent, Some(2), "Should have indent of 2");
}

/// Test justify horizontal with distributed vertical
#[test]
fn test_combination_justify_distributed() {
    let xml = create_styles_xml(r#"horizontal="justify" vertical="distributed""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("justify".to_string()),
        "Should have justify horizontal alignment"
    );
    assert_eq!(
        alignment.vertical,
        Some("distributed".to_string()),
        "Should have distributed vertical alignment"
    );
}

/// Test center continuous with indent (used for grouped headers)
#[test]
fn test_combination_center_continuous_indent() {
    let xml = create_styles_xml(r#"horizontal="centerContinuous" indent="1""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    assert_eq!(
        alignment.horizontal,
        Some("centerContinuous".to_string()),
        "Should have centerContinuous alignment"
    );
    assert_eq!(alignment.indent, Some(1), "Should have indent of 1");
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

/// Test empty alignment element
#[test]
fn test_empty_alignment_element() {
    let xml = create_styles_xml("");
    let alignment = parse_alignment(&xml);

    // An empty alignment element should still be parsed, with defaults
    assert!(
        alignment.is_some(),
        "Empty alignment element should create RawAlignment"
    );

    let align = alignment.unwrap();
    assert_eq!(align.horizontal, None, "Empty should have no horizontal");
    assert_eq!(align.vertical, None, "Empty should have no vertical");
    assert!(!align.wrap_text, "Empty should have wrap_text false");
    assert_eq!(align.indent, None, "Empty should have no indent");
    assert_eq!(align.text_rotation, None, "Empty should have no rotation");
}

/// Test no alignment element at all
#[test]
fn test_no_alignment_element() {
    let xml = create_styles_xml_no_alignment();
    let stylesheet = parse_stylesheet(&xml);

    assert!(
        !stylesheet.cell_xfs.is_empty(),
        "Should have at least one xf"
    );

    let xf = stylesheet.cell_xfs.first().unwrap();
    assert!(
        xf.alignment.is_none(),
        "xf without alignment element should have None alignment"
    );
}

/// Test attributes in different order
#[test]
fn test_attribute_order_independence() {
    // Different attribute order should produce same result
    let xml1 = create_styles_xml(r#"vertical="top" horizontal="left" indent="1" wrapText="1""#);
    let xml2 = create_styles_xml(r#"horizontal="left" wrapText="1" vertical="top" indent="1""#);

    let align1 = parse_alignment(&xml1).expect("Should parse first XML");
    let align2 = parse_alignment(&xml2).expect("Should parse second XML");

    assert_eq!(
        align1.horizontal, align2.horizontal,
        "Horizontal should match regardless of order"
    );
    assert_eq!(
        align1.vertical, align2.vertical,
        "Vertical should match regardless of order"
    );
    assert_eq!(
        align1.wrap_text, align2.wrap_text,
        "wrap_text should match regardless of order"
    );
    assert_eq!(
        align1.indent, align2.indent,
        "indent should match regardless of order"
    );
}

/// Test self-closing alignment element
#[test]
fn test_self_closing_alignment() {
    // Self-closing <alignment .../> vs <alignment ...></alignment>
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="1">
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

    let alignment = parse_alignment(xml).expect("Should parse self-closing alignment");

    assert_eq!(
        alignment.horizontal,
        Some("center".to_string()),
        "Should parse horizontal from self-closing element"
    );
    assert_eq!(
        alignment.vertical,
        Some("center".to_string()),
        "Should parse vertical from self-closing element"
    );
}

// ============================================================================
// MULTIPLE XF TESTS
// ============================================================================

/// Test multiple xf elements with different alignments
#[test]
fn test_multiple_xf_different_alignments() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="3">
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="left"/>
    </xf>
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="right" wrapText="1" indent="2"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

    let stylesheet = parse_stylesheet(xml);

    assert_eq!(stylesheet.cell_xfs.len(), 3, "Should have 3 xf elements");

    // First xf: left aligned
    let xf0 = &stylesheet.cell_xfs[0];
    let align0 = xf0
        .alignment
        .as_ref()
        .expect("First xf should have alignment");
    assert_eq!(align0.horizontal, Some("left".to_string()));
    assert_eq!(align0.vertical, None);

    // Second xf: center/center
    let xf1 = &stylesheet.cell_xfs[1];
    let align1 = xf1
        .alignment
        .as_ref()
        .expect("Second xf should have alignment");
    assert_eq!(align1.horizontal, Some("center".to_string()));
    assert_eq!(align1.vertical, Some("center".to_string()));

    // Third xf: right with wrap and indent
    let xf2 = &stylesheet.cell_xfs[2];
    let align2 = xf2
        .alignment
        .as_ref()
        .expect("Third xf should have alignment");
    assert_eq!(align2.horizontal, Some("right".to_string()));
    assert!(align2.wrap_text);
    assert_eq!(align2.indent, Some(2));
}

/// Test xf with applyAlignment="0" (alignment should still be parsed)
#[test]
fn test_apply_alignment_false() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="1">
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="0">
      <alignment horizontal="center"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

    let stylesheet = parse_stylesheet(xml);
    let xf = stylesheet.cell_xfs.first().expect("Should have xf");

    // The alignment element should still be parsed
    // applyAlignment="0" means don't apply in rendering, but we still parse it
    assert!(
        xf.alignment.is_some(),
        "Alignment should be parsed even with applyAlignment='0'"
    );

    // Verify applyAlignment flag is correctly set
    assert!(
        !xf.apply_alignment,
        "apply_alignment should be false when applyAlignment='0'"
    );
}

// ============================================================================
// SPECIAL VALUE TESTS
// ============================================================================

/// Test textRotation with all valid rotation values (0-90 counterclockwise, 91-180 for clockwise)
#[test]
fn test_text_rotation_full_range() {
    // Counterclockwise: 1-90
    for rotation in [1, 15, 30, 45, 60, 75, 89, 90] {
        let xml = create_styles_xml(&format!(r#"textRotation="{rotation}""#));
        let alignment = parse_alignment(&xml).expect("Should have alignment");
        assert_eq!(
            alignment.text_rotation,
            Some(rotation),
            "Should parse rotation {rotation}"
        );
    }

    // Clockwise (stored as 90 + degrees): 91-180
    for rotation in [91, 105, 120, 135, 150, 165, 179, 180] {
        let xml = create_styles_xml(&format!(r#"textRotation="{rotation}""#));
        let alignment = parse_alignment(&xml).expect("Should have alignment");
        assert_eq!(
            alignment.text_rotation,
            Some(rotation),
            "Should parse rotation {rotation}"
        );
    }
}

/// Test that "true" and "false" are not valid for wrapText (only "1" and "0")
#[test]
fn test_wrap_text_invalid_values() {
    // "true" should not be treated as true
    let xml = create_styles_xml(r#"wrapText="true""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");
    // The parser checks for "1", so "true" results in false
    assert!(
        !alignment.wrap_text,
        "wrapText='true' should not be treated as true (OOXML uses '1')"
    );

    // "false" should result in false
    let xml = create_styles_xml(r#"wrapText="false""#);
    let alignment = parse_alignment(&xml).expect("Should have alignment");
    assert!(
        !alignment.wrap_text,
        "wrapText='false' should result in false"
    );
}

// ============================================================================
// WHITESPACE AND FORMATTING TESTS
// ============================================================================

/// Test that extra whitespace in XML doesn't affect parsing
#[test]
fn test_whitespace_handling() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
    </border>
  </borders>
  <cellXfs count="1">
    <xf fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment
          horizontal="center"
          vertical="center"
          wrapText="1"
          indent="2"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

    let alignment = parse_alignment(xml).expect("Should parse with whitespace");

    assert_eq!(alignment.horizontal, Some("center".to_string()));
    assert_eq!(alignment.vertical, Some("center".to_string()));
    assert!(alignment.wrap_text);
    assert_eq!(alignment.indent, Some(2));
}

// ============================================================================
// DEFAULT VALUE TESTS
// ============================================================================

/// Test default values when alignment is present but empty
#[test]
fn test_default_values() {
    let xml = create_styles_xml("");
    let alignment = parse_alignment(&xml).expect("Should have alignment");

    // Verify all defaults
    assert_eq!(
        alignment.horizontal, None,
        "Default horizontal should be None"
    );
    assert_eq!(alignment.vertical, None, "Default vertical should be None");
    assert!(!alignment.wrap_text, "Default wrap_text should be false");
    assert_eq!(alignment.indent, None, "Default indent should be None");
    assert_eq!(
        alignment.text_rotation, None,
        "Default text_rotation should be None"
    );
}

/// Test that RawAlignment::default() produces expected defaults
#[test]
fn test_raw_alignment_default() {
    let default = RawAlignment::default();

    assert_eq!(default.horizontal, None);
    assert_eq!(default.vertical, None);
    assert!(!default.wrap_text);
    assert_eq!(default.indent, None);
    assert_eq!(default.text_rotation, None);
}

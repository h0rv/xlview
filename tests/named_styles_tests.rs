//! Named styles (cellStyles) tests for xlview
//!
//! Tests parsing of named styles and cellStyleXfs inheritance from styles.xml
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

use xlview::styles::parse_styles;

// ============================================================================
// Basic Named Style Parsing Tests
// ============================================================================

#[test]
fn test_parse_single_named_style() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles.len(), 1);
    assert_eq!(stylesheet.named_styles[0].name, "Normal");
    assert_eq!(stylesheet.named_styles[0].xf_id, 0);
    assert_eq!(stylesheet.named_styles[0].builtin_id, Some(0));
}

#[test]
fn test_parse_multiple_named_styles() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles.len(), 2);

    assert_eq!(stylesheet.named_styles[0].name, "Normal");
    assert_eq!(stylesheet.named_styles[0].xf_id, 0);
    assert_eq!(stylesheet.named_styles[0].builtin_id, Some(0));

    assert_eq!(stylesheet.named_styles[1].name, "Heading 1");
    assert_eq!(stylesheet.named_styles[1].xf_id, 1);
    assert_eq!(stylesheet.named_styles[1].builtin_id, Some(16));
}

#[test]
fn test_parse_named_style_without_builtin_id() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Custom Style" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles.len(), 1);
    assert_eq!(stylesheet.named_styles[0].name, "Custom Style");
    assert_eq!(stylesheet.named_styles[0].xf_id, 0);
    assert_eq!(stylesheet.named_styles[0].builtin_id, None);
}

// ============================================================================
// cellStyleXfs Parsing Tests
// ============================================================================

#[test]
fn test_parse_cell_style_xfs() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.cell_style_xfs.len(), 2);

    // First cellStyleXf
    assert_eq!(stylesheet.cell_style_xfs[0].font_id, Some(0));
    assert_eq!(stylesheet.cell_style_xfs[0].fill_id, Some(0));
    assert_eq!(stylesheet.cell_style_xfs[0].border_id, Some(0));
    assert_eq!(stylesheet.cell_style_xfs[0].num_fmt_id, Some(0));

    // Second cellStyleXf (Heading 1)
    assert_eq!(stylesheet.cell_style_xfs[1].font_id, Some(1));
    assert_eq!(stylesheet.cell_style_xfs[1].fill_id, Some(1));
    assert_eq!(stylesheet.cell_style_xfs[1].border_id, Some(0));
}

#[test]
fn test_cell_xf_references_cell_style_xf() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.cell_xfs.len(), 2);

    // First cellXf references cellStyleXfs[0]
    assert_eq!(stylesheet.cell_xfs[0].xf_id, Some(0));

    // Second cellXf references cellStyleXfs[1] (Heading 1 style)
    assert_eq!(stylesheet.cell_xfs[1].xf_id, Some(1));
}

// ============================================================================
// Built-in Style ID Tests
// ============================================================================

#[test]
fn test_common_builtin_style_ids() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="6">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
    <font><name val="Calibri"/><sz val="13"/><b/></font>
    <font><name val="Calibri"/><sz val="11"/><b/></font>
    <font><name val="Calibri"/><sz val="11"/><b/><i/></font>
    <font><name val="Calibri"/><sz val="11"/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="6">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
    <cellStyle name="Heading 2" xfId="2" builtinId="17"/>
    <cellStyle name="Heading 3" xfId="3" builtinId="18"/>
    <cellStyle name="Title" xfId="4" builtinId="15"/>
    <cellStyle name="Bad" xfId="5" builtinId="27"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles.len(), 6);

    // Normal - builtinId 0
    assert_eq!(stylesheet.named_styles[0].name, "Normal");
    assert_eq!(stylesheet.named_styles[0].builtin_id, Some(0));

    // Heading 1 - builtinId 16
    assert_eq!(stylesheet.named_styles[1].name, "Heading 1");
    assert_eq!(stylesheet.named_styles[1].builtin_id, Some(16));

    // Heading 2 - builtinId 17
    assert_eq!(stylesheet.named_styles[2].name, "Heading 2");
    assert_eq!(stylesheet.named_styles[2].builtin_id, Some(17));

    // Heading 3 - builtinId 18
    assert_eq!(stylesheet.named_styles[3].name, "Heading 3");
    assert_eq!(stylesheet.named_styles[3].builtin_id, Some(18));

    // Title - builtinId 15
    assert_eq!(stylesheet.named_styles[4].name, "Title");
    assert_eq!(stylesheet.named_styles[4].builtin_id, Some(15));

    // Bad - builtinId 27
    assert_eq!(stylesheet.named_styles[5].name, "Bad");
    assert_eq!(stylesheet.named_styles[5].builtin_id, Some(27));
}

// ============================================================================
// Empty/Missing Sections Tests
// ============================================================================

#[test]
fn test_missing_cell_styles_section() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // Should have empty named_styles
    assert!(stylesheet.named_styles.is_empty());
    // But should still have cellStyleXfs
    assert_eq!(stylesheet.cell_style_xfs.len(), 1);
}

#[test]
fn test_missing_cell_style_xfs_section() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // Should have named_styles
    assert_eq!(stylesheet.named_styles.len(), 1);
    // But empty cellStyleXfs
    assert!(stylesheet.cell_style_xfs.is_empty());
}

// ============================================================================
// Style Inheritance Tests (cellXfs inherits from cellStyleXfs)
// ============================================================================

#[test]
fn test_cell_xf_with_apply_flags() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // Second cellXf should reference Heading 1 style (xfId=1)
    // but override fill (applyFill=1)
    assert_eq!(stylesheet.cell_xfs[1].xf_id, Some(1));
    assert_eq!(stylesheet.cell_xfs[1].fill_id, Some(1));
    assert!(stylesheet.cell_xfs[1].apply_fill);
}

#[test]
fn test_cell_style_xf_with_alignment() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Centered" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // cellStyleXf should have alignment
    assert!(stylesheet.cell_style_xfs[0].alignment.is_some());
    let align = stylesheet.cell_style_xfs[0].alignment.as_ref().unwrap();
    assert_eq!(align.horizontal, Some("center".to_string()));
    assert_eq!(align.vertical, Some("center".to_string()));
    assert!(align.wrap_text);
}

// ============================================================================
// Real-world Style Configuration Tests
// ============================================================================

#[test]
fn test_typical_excel_style_structure() {
    // Simulates a typical Excel file with Normal, Heading 1, and custom styles
    let styles_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="3">
    <font><name val="Calibri"/><sz val="11"/><color theme="1"/></font>
    <font><name val="Calibri"/><sz val="15"/><b/><color theme="4"/></font>
    <font><name val="Calibri"/><sz val="11"/><color rgb="FF9C0006"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border/>
    <border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom></border>
  </borders>
  <cellStyleXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" applyFont="1" applyFill="1"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="2" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="3">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
    <cellStyle name="Bad" xfId="2" builtinId="27"/>
  </cellStyles>
</styleSheet>"##;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // Verify num formats
    assert_eq!(stylesheet.num_fmts.len(), 1);
    assert_eq!(stylesheet.num_fmts[0], (164, r##"#,##0.00"##.to_string()));

    // Verify fonts
    assert_eq!(stylesheet.fonts.len(), 3);
    assert!(stylesheet.fonts[1].bold);

    // Verify fills
    assert_eq!(stylesheet.fills.len(), 3);
    assert_eq!(stylesheet.fills[2].pattern_type, Some("solid".to_string()));

    // Verify borders
    assert_eq!(stylesheet.borders.len(), 2);
    assert!(stylesheet.borders[1].left.is_some());

    // Verify cellStyleXfs
    assert_eq!(stylesheet.cell_style_xfs.len(), 3);

    // Verify cellXfs
    assert_eq!(stylesheet.cell_xfs.len(), 4);
    assert_eq!(stylesheet.cell_xfs[3].num_fmt_id, Some(164));
    assert!(stylesheet.cell_xfs[3].apply_number_format);

    // Verify named styles
    assert_eq!(stylesheet.named_styles.len(), 3);
    assert_eq!(stylesheet.named_styles[0].name, "Normal");
    assert_eq!(stylesheet.named_styles[1].name, "Heading 1");
    assert_eq!(stylesheet.named_styles[2].name, "Bad");
}

#[test]
fn test_named_style_with_protection() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1">
      <protection locked="0" hidden="1"/>
    </xf>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Unlocked" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    // Verify protection on cellStyleXf
    assert!(stylesheet.cell_style_xfs[0].protection.is_some());
    let protection = stylesheet.cell_style_xfs[0].protection.as_ref().unwrap();
    assert!(!protection.locked); // unlocked
    assert!(protection.hidden); // formula hidden
}

// ============================================================================
// Edge Cases
// ============================================================================

#[test]
fn test_named_style_with_special_characters_in_name() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="My &quot;Custom&quot; Style &amp; More" xfId="0"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles.len(), 1);
    // Note: quick-xml does not automatically unescape XML entities in attribute values
    // by default. The raw attribute value is preserved.
    assert_eq!(
        stylesheet.named_styles[0].name,
        "My &quot;Custom&quot; Style &amp; More"
    );
}

#[test]
fn test_large_builtin_id() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><name val="Calibri"/><sz val="11"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Custom Style" xfId="0" builtinId="999"/>
  </cellStyles>
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert_eq!(stylesheet.named_styles[0].builtin_id, Some(999));
}

#[test]
fn test_empty_stylesheet() {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</styleSheet>"#;

    let stylesheet = parse_styles(Cursor::new(styles_xml)).unwrap();

    assert!(stylesheet.named_styles.is_empty());
    assert!(stylesheet.cell_style_xfs.is_empty());
    assert!(stylesheet.cell_xfs.is_empty());
    assert!(stylesheet.fonts.is_empty());
    assert!(stylesheet.fills.is_empty());
    assert!(stylesheet.borders.is_empty());
}

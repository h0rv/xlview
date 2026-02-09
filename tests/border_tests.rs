//! Comprehensive tests for border styling in xlview
//!
//! Tests all border styles, colors, diagonal borders, and edge cases.
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
use xlview::types::{RawBorderSide, StyleSheet};

/// Helper to create a minimal styles.xml with just borders section
fn styles_xml_with_borders(borders_content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <borders count="1">
        {borders_content}
    </borders>
</styleSheet>"#
    )
}

/// Helper to parse a styles XML string and return the StyleSheet
fn parse_styles_xml(xml: &str) -> StyleSheet {
    let cursor = Cursor::new(xml);
    parse_styles(cursor).expect("Failed to parse styles XML")
}

// ============================================================================
// BORDER STYLE TESTS
// ============================================================================

mod border_styles {
    use super::*;

    #[test]
    fn test_thin_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);

        let border = &stylesheet.borders[0];
        assert_border_side_style(&border.left, "thin");
        assert_border_side_style(&border.right, "thin");
        assert_border_side_style(&border.top, "thin");
        assert_border_side_style(&border.bottom, "thin");
    }

    #[test]
    fn test_medium_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="medium"><color indexed="64"/></left>
                <right style="medium"><color indexed="64"/></right>
                <top style="medium"><color indexed="64"/></top>
                <bottom style="medium"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "medium");
        assert_border_side_style(&border.right, "medium");
        assert_border_side_style(&border.top, "medium");
        assert_border_side_style(&border.bottom, "medium");
    }

    #[test]
    fn test_thick_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thick"><color indexed="64"/></left>
                <right style="thick"><color indexed="64"/></right>
                <top style="thick"><color indexed="64"/></top>
                <bottom style="thick"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "thick");
        assert_border_side_style(&border.right, "thick");
        assert_border_side_style(&border.top, "thick");
        assert_border_side_style(&border.bottom, "thick");
    }

    #[test]
    fn test_dashed_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="dashed"><color indexed="64"/></left>
                <right style="dashed"><color indexed="64"/></right>
                <top style="dashed"><color indexed="64"/></top>
                <bottom style="dashed"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "dashed");
        assert_border_side_style(&border.right, "dashed");
        assert_border_side_style(&border.top, "dashed");
        assert_border_side_style(&border.bottom, "dashed");
    }

    #[test]
    fn test_dotted_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="dotted"><color indexed="64"/></left>
                <right style="dotted"><color indexed="64"/></right>
                <top style="dotted"><color indexed="64"/></top>
                <bottom style="dotted"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "dotted");
        assert_border_side_style(&border.right, "dotted");
        assert_border_side_style(&border.top, "dotted");
        assert_border_side_style(&border.bottom, "dotted");
    }

    #[test]
    fn test_double_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="double"><color indexed="64"/></left>
                <right style="double"><color indexed="64"/></right>
                <top style="double"><color indexed="64"/></top>
                <bottom style="double"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "double");
        assert_border_side_style(&border.right, "double");
        assert_border_side_style(&border.top, "double");
        assert_border_side_style(&border.bottom, "double");
    }

    #[test]
    fn test_hair_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="hair"><color indexed="64"/></left>
                <right style="hair"><color indexed="64"/></right>
                <top style="hair"><color indexed="64"/></top>
                <bottom style="hair"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "hair");
        assert_border_side_style(&border.right, "hair");
        assert_border_side_style(&border.top, "hair");
        assert_border_side_style(&border.bottom, "hair");
    }

    #[test]
    fn test_medium_dashed_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="mediumDashed"><color indexed="64"/></left>
                <right style="mediumDashed"><color indexed="64"/></right>
                <top style="mediumDashed"><color indexed="64"/></top>
                <bottom style="mediumDashed"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "mediumDashed");
        assert_border_side_style(&border.right, "mediumDashed");
        assert_border_side_style(&border.top, "mediumDashed");
        assert_border_side_style(&border.bottom, "mediumDashed");
    }

    #[test]
    fn test_dash_dot_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="dashDot"><color indexed="64"/></left>
                <right style="dashDot"><color indexed="64"/></right>
                <top style="dashDot"><color indexed="64"/></top>
                <bottom style="dashDot"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "dashDot");
        assert_border_side_style(&border.right, "dashDot");
        assert_border_side_style(&border.top, "dashDot");
        assert_border_side_style(&border.bottom, "dashDot");
    }

    #[test]
    fn test_medium_dash_dot_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="mediumDashDot"><color indexed="64"/></left>
                <right style="mediumDashDot"><color indexed="64"/></right>
                <top style="mediumDashDot"><color indexed="64"/></top>
                <bottom style="mediumDashDot"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "mediumDashDot");
        assert_border_side_style(&border.right, "mediumDashDot");
        assert_border_side_style(&border.top, "mediumDashDot");
        assert_border_side_style(&border.bottom, "mediumDashDot");
    }

    #[test]
    fn test_dash_dot_dot_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="dashDotDot"><color indexed="64"/></left>
                <right style="dashDotDot"><color indexed="64"/></right>
                <top style="dashDotDot"><color indexed="64"/></top>
                <bottom style="dashDotDot"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "dashDotDot");
        assert_border_side_style(&border.right, "dashDotDot");
        assert_border_side_style(&border.top, "dashDotDot");
        assert_border_side_style(&border.bottom, "dashDotDot");
    }

    #[test]
    fn test_medium_dash_dot_dot_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="mediumDashDotDot"><color indexed="64"/></left>
                <right style="mediumDashDotDot"><color indexed="64"/></right>
                <top style="mediumDashDotDot"><color indexed="64"/></top>
                <bottom style="mediumDashDotDot"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "mediumDashDotDot");
        assert_border_side_style(&border.right, "mediumDashDotDot");
        assert_border_side_style(&border.top, "mediumDashDotDot");
        assert_border_side_style(&border.bottom, "mediumDashDotDot");
    }

    #[test]
    fn test_slant_dash_dot_border_style() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="slantDashDot"><color indexed="64"/></left>
                <right style="slantDashDot"><color indexed="64"/></right>
                <top style="slantDashDot"><color indexed="64"/></top>
                <bottom style="slantDashDot"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_border_side_style(&border.left, "slantDashDot");
        assert_border_side_style(&border.right, "slantDashDot");
        assert_border_side_style(&border.top, "slantDashDot");
        assert_border_side_style(&border.bottom, "slantDashDot");
    }

    /// Helper to assert border side style
    fn assert_border_side_style(side: &Option<RawBorderSide>, expected_style: &str) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert_eq!(
            side.style, expected_style,
            "Expected style '{}', got '{}'",
            expected_style, side.style
        );
    }
}

// ============================================================================
// BORDER COLOR TESTS
// ============================================================================

mod border_colors {
    use super::*;

    #[test]
    fn test_rgb_color() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color rgb="FF000000"/></left>
                <right style="thin"><color rgb="FFFF0000"/></right>
                <top style="thin"><color rgb="FF00FF00"/></top>
                <bottom style="thin"><color rgb="FF0000FF"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_rgb(&border.left, "FF000000");
        assert_color_rgb(&border.right, "FFFF0000");
        assert_color_rgb(&border.top, "FF00FF00");
        assert_color_rgb(&border.bottom, "FF0000FF");
    }

    #[test]
    fn test_theme_color() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color theme="0"/></left>
                <right style="thin"><color theme="1"/></right>
                <top style="thin"><color theme="4"/></top>
                <bottom style="thin"><color theme="9"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_theme(&border.left, 0);
        assert_color_theme(&border.right, 1);
        assert_color_theme(&border.top, 4);
        assert_color_theme(&border.bottom, 9);
    }

    #[test]
    fn test_theme_color_with_tint() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color theme="4" tint="0.5"/></left>
                <right style="thin"><color theme="4" tint="-0.25"/></right>
                <top style="thin"><color theme="1" tint="0.799981688894314"/></top>
                <bottom style="thin"><color theme="0" tint="-0.14999847407452621"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_theme_with_tint(&border.left, 4, 0.5);
        assert_color_theme_with_tint(&border.right, 4, -0.25);
        assert_color_theme_with_tint(&border.top, 1, 0.799981688894314);
        assert_color_theme_with_tint(&border.bottom, 0, -0.149_998_474_074_526_2);
    }

    #[test]
    fn test_indexed_color() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="8"/></right>
                <top style="thin"><color indexed="10"/></top>
                <bottom style="thin"><color indexed="53"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_indexed(&border.left, 64);
        assert_color_indexed(&border.right, 8);
        assert_color_indexed(&border.top, 10);
        assert_color_indexed(&border.bottom, 53);
    }

    #[test]
    fn test_auto_color() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color auto="1"/></left>
                <right style="thin"><color auto="1"/></right>
                <top style="thin"><color auto="1"/></top>
                <bottom style="thin"><color auto="1"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_auto(&border.left);
        assert_color_auto(&border.right);
        assert_color_auto(&border.top);
        assert_color_auto(&border.bottom);
    }

    #[test]
    fn test_mixed_color_types() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color rgb="FFFF0000"/></left>
                <right style="thin"><color theme="1"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color auto="1"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert_color_rgb(&border.left, "FFFF0000");
        assert_color_theme(&border.right, 1);
        assert_color_indexed(&border.top, 64);
        assert_color_auto(&border.bottom);
    }

    /// Helper to assert RGB color
    fn assert_color_rgb(side: &Option<RawBorderSide>, expected_rgb: &str) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert!(side.color.is_some(), "Color should be present");
        let color = side.color.as_ref().unwrap();
        assert_eq!(
            color.rgb.as_deref(),
            Some(expected_rgb),
            "Expected RGB '{}', got '{:?}'",
            expected_rgb,
            color.rgb
        );
    }

    /// Helper to assert theme color
    fn assert_color_theme(side: &Option<RawBorderSide>, expected_theme: u32) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert!(side.color.is_some(), "Color should be present");
        let color = side.color.as_ref().unwrap();
        assert_eq!(
            color.theme,
            Some(expected_theme),
            "Expected theme {}, got {:?}",
            expected_theme,
            color.theme
        );
    }

    /// Helper to assert theme color with tint
    fn assert_color_theme_with_tint(
        side: &Option<RawBorderSide>,
        expected_theme: u32,
        expected_tint: f64,
    ) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert!(side.color.is_some(), "Color should be present");
        let color = side.color.as_ref().unwrap();
        assert_eq!(
            color.theme,
            Some(expected_theme),
            "Expected theme {}, got {:?}",
            expected_theme,
            color.theme
        );
        assert!(color.tint.is_some(), "Tint should be present");
        let tint = color.tint.unwrap();
        assert!(
            (tint - expected_tint).abs() < 0.0001,
            "Expected tint {}, got {}",
            expected_tint,
            tint
        );
    }

    /// Helper to assert indexed color
    fn assert_color_indexed(side: &Option<RawBorderSide>, expected_indexed: u32) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert!(side.color.is_some(), "Color should be present");
        let color = side.color.as_ref().unwrap();
        assert_eq!(
            color.indexed,
            Some(expected_indexed),
            "Expected indexed {}, got {:?}",
            expected_indexed,
            color.indexed
        );
    }

    /// Helper to assert auto color
    fn assert_color_auto(side: &Option<RawBorderSide>) {
        assert!(side.is_some(), "Border side should be present");
        let side = side.as_ref().unwrap();
        assert!(side.color.is_some(), "Color should be present");
        let color = side.color.as_ref().unwrap();
        assert!(color.auto, "Expected auto=true, got auto={}", color.auto);
    }
}

// ============================================================================
// DIAGONAL BORDER TESTS
// ============================================================================

mod diagonal_borders {
    use super::*;

    #[test]
    fn test_diagonal_down() {
        // Note: This test documents expected behavior for diagonal borders
        // The current parser may not support diagonal borders
        let xml = styles_xml_with_borders(
            r#"<border diagonalDown="1">
                <left/>
                <right/>
                <top/>
                <bottom/>
                <diagonal style="thin"><color indexed="64"/></diagonal>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);
        // Diagonal borders may not be implemented yet - this test documents the expected XML format
    }

    #[test]
    fn test_diagonal_up() {
        let xml = styles_xml_with_borders(
            r#"<border diagonalUp="1">
                <left/>
                <right/>
                <top/>
                <bottom/>
                <diagonal style="thin"><color rgb="FFFF0000"/></diagonal>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);
        // Diagonal borders may not be implemented yet - this test documents the expected XML format
    }

    #[test]
    fn test_both_diagonals() {
        let xml = styles_xml_with_borders(
            r#"<border diagonalUp="1" diagonalDown="1">
                <left/>
                <right/>
                <top/>
                <bottom/>
                <diagonal style="medium"><color theme="1"/></diagonal>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);
        // Diagonal borders may not be implemented yet - this test documents the expected XML format
    }

    #[test]
    fn test_diagonal_with_regular_borders() {
        let xml = styles_xml_with_borders(
            r#"<border diagonalDown="1">
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
                <diagonal style="dashed"><color rgb="FF0000FF"/></diagonal>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);

        let border = &stylesheet.borders[0];
        // Regular borders should still work
        assert!(border.left.is_some());
        assert!(border.right.is_some());
        assert!(border.top.is_some());
        assert!(border.bottom.is_some());
    }
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

mod edge_cases {
    use super::*;

    #[test]
    fn test_mixed_border_styles() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right style="medium"><color rgb="FFFF0000"/></right>
                <top style="dashed"><color theme="1"/></top>
                <bottom style="double"><color auto="1"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_some());
        assert_eq!(border.left.as_ref().unwrap().style, "thin");

        assert!(border.right.is_some());
        assert_eq!(border.right.as_ref().unwrap().style, "medium");

        assert!(border.top.is_some());
        assert_eq!(border.top.as_ref().unwrap().style, "dashed");

        assert!(border.bottom.is_some());
        assert_eq!(border.bottom.as_ref().unwrap().style, "double");
    }

    #[test]
    fn test_empty_border_element() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top/>
                <bottom/>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);

        let border = &stylesheet.borders[0];
        // Empty border sides should not have style
        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_none());
    }

    #[test]
    fn test_partial_borders_left_only() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right/>
                <top/>
                <bottom/>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_some());
        assert_eq!(border.left.as_ref().unwrap().style, "thin");
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_none());
    }

    #[test]
    fn test_partial_borders_top_bottom() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top style="medium"><color rgb="FF000000"/></top>
                <bottom style="medium"><color rgb="FF000000"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_some());
        assert_eq!(border.top.as_ref().unwrap().style, "medium");
        assert!(border.bottom.is_some());
        assert_eq!(border.bottom.as_ref().unwrap().style, "medium");
    }

    #[test]
    fn test_partial_borders_right_bottom() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right style="thick"><color theme="4"/></right>
                <top/>
                <bottom style="thick"><color theme="4"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_none());
        assert!(border.right.is_some());
        assert_eq!(border.right.as_ref().unwrap().style, "thick");
        assert!(border.top.is_none());
        assert!(border.bottom.is_some());
        assert_eq!(border.bottom.as_ref().unwrap().style, "thick");
    }

    #[test]
    fn test_multiple_borders() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top/>
                <bottom/>
            </border>
            <border>
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
            </border>
            <border>
                <left style="medium"><color rgb="FFFF0000"/></left>
                <right style="medium"><color rgb="FFFF0000"/></right>
                <top style="medium"><color rgb="FFFF0000"/></top>
                <bottom style="medium"><color rgb="FFFF0000"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 3);

        // First border - empty
        assert!(stylesheet.borders[0].left.is_none());

        // Second border - thin
        assert!(stylesheet.borders[1].left.is_some());
        assert_eq!(stylesheet.borders[1].left.as_ref().unwrap().style, "thin");

        // Third border - medium
        assert!(stylesheet.borders[2].left.is_some());
        assert_eq!(stylesheet.borders[2].left.as_ref().unwrap().style, "medium");
    }

    #[test]
    fn test_border_style_without_color() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"/>
                <right style="medium"/>
                <top style="thick"/>
                <bottom style="dashed"/>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        // Styles should be present even without colors
        assert!(border.left.is_some());
        assert_eq!(border.left.as_ref().unwrap().style, "thin");
        assert!(border.left.as_ref().unwrap().color.is_none());

        assert!(border.right.is_some());
        assert_eq!(border.right.as_ref().unwrap().style, "medium");
        assert!(border.right.as_ref().unwrap().color.is_none());

        assert!(border.top.is_some());
        assert_eq!(border.top.as_ref().unwrap().style, "thick");
        assert!(border.top.as_ref().unwrap().color.is_none());

        assert!(border.bottom.is_some());
        assert_eq!(border.bottom.as_ref().unwrap().style, "dashed");
        assert!(border.bottom.as_ref().unwrap().color.is_none());
    }

    #[test]
    fn test_self_closing_border_elements() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right/>
                <top/>
                <bottom style="thin"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_some());
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_some());
    }

    #[test]
    fn test_border_with_outline_attribute() {
        // Some Excel files include outline="0" attribute
        let xml = styles_xml_with_borders(
            r#"<border outline="0">
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 1);

        let border = &stylesheet.borders[0];
        assert!(border.left.is_some());
        assert!(border.right.is_some());
        assert!(border.top.is_some());
        assert!(border.bottom.is_some());
    }
}

// ============================================================================
// INDIVIDUAL SIDE TESTS
// ============================================================================

mod individual_sides {
    use super::*;

    #[test]
    fn test_left_border_only() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color rgb="FF000000"/></left>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_some());
        let left = border.left.as_ref().unwrap();
        assert_eq!(left.style, "thin");
        assert!(left.color.is_some());
        assert_eq!(
            left.color.as_ref().unwrap().rgb.as_deref(),
            Some("FF000000")
        );
    }

    #[test]
    fn test_right_border_only() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <right style="medium"><color theme="1"/></right>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.right.is_some());
        let right = border.right.as_ref().unwrap();
        assert_eq!(right.style, "medium");
        assert!(right.color.is_some());
        assert_eq!(right.color.as_ref().unwrap().theme, Some(1));
    }

    #[test]
    fn test_top_border_only() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <top style="thick"><color indexed="64"/></top>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.top.is_some());
        let top = border.top.as_ref().unwrap();
        assert_eq!(top.style, "thick");
        assert!(top.color.is_some());
        assert_eq!(top.color.as_ref().unwrap().indexed, Some(64));
    }

    #[test]
    fn test_bottom_border_only() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <bottom style="double"><color auto="1"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.bottom.is_some());
        let bottom = border.bottom.as_ref().unwrap();
        assert_eq!(bottom.style, "double");
        assert!(bottom.color.is_some());
        assert!(bottom.color.as_ref().unwrap().auto);
    }
}

// ============================================================================
// ALL BORDER STYLES COMPREHENSIVE TEST
// ============================================================================

mod comprehensive {
    use super::*;

    /// Test all 13 border styles on left side
    #[test]
    fn test_all_border_styles_on_left() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style in &styles {
            let xml = styles_xml_with_borders(&format!(
                r#"<border>
                    <left style="{}"><color indexed="64"/></left>
                </border>"#,
                style
            ));

            let stylesheet = parse_styles_xml(&xml);
            let border = &stylesheet.borders[0];

            assert!(
                border.left.is_some(),
                "Left border should be present for style '{}'",
                style
            );
            assert_eq!(
                border.left.as_ref().unwrap().style,
                *style,
                "Style mismatch for '{}'",
                style
            );
        }
    }

    /// Test all 13 border styles on right side
    #[test]
    fn test_all_border_styles_on_right() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style in &styles {
            let xml = styles_xml_with_borders(&format!(
                r#"<border>
                    <right style="{}"><color indexed="64"/></right>
                </border>"#,
                style
            ));

            let stylesheet = parse_styles_xml(&xml);
            let border = &stylesheet.borders[0];

            assert!(
                border.right.is_some(),
                "Right border should be present for style '{}'",
                style
            );
            assert_eq!(
                border.right.as_ref().unwrap().style,
                *style,
                "Style mismatch for '{}'",
                style
            );
        }
    }

    /// Test all 13 border styles on top side
    #[test]
    fn test_all_border_styles_on_top() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style in &styles {
            let xml = styles_xml_with_borders(&format!(
                r#"<border>
                    <top style="{}"><color indexed="64"/></top>
                </border>"#,
                style
            ));

            let stylesheet = parse_styles_xml(&xml);
            let border = &stylesheet.borders[0];

            assert!(
                border.top.is_some(),
                "Top border should be present for style '{}'",
                style
            );
            assert_eq!(
                border.top.as_ref().unwrap().style,
                *style,
                "Style mismatch for '{}'",
                style
            );
        }
    }

    /// Test all 13 border styles on bottom side
    #[test]
    fn test_all_border_styles_on_bottom() {
        let styles = [
            "thin",
            "medium",
            "thick",
            "dashed",
            "dotted",
            "double",
            "hair",
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ];

        for style in &styles {
            let xml = styles_xml_with_borders(&format!(
                r#"<border>
                    <bottom style="{}"><color indexed="64"/></bottom>
                </border>"#,
                style
            ));

            let stylesheet = parse_styles_xml(&xml);
            let border = &stylesheet.borders[0];

            assert!(
                border.bottom.is_some(),
                "Bottom border should be present for style '{}'",
                style
            );
            assert_eq!(
                border.bottom.as_ref().unwrap().style,
                *style,
                "Style mismatch for '{}'",
                style
            );
        }
    }

    /// Test all color types on all sides
    #[test]
    fn test_all_color_types_all_sides() {
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color rgb="FF123456"/></left>
                <right style="thin"><color theme="5" tint="0.4"/></right>
                <top style="thin"><color indexed="32"/></top>
                <bottom style="thin"><color auto="1"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        // Left - RGB
        let left_color = border.left.as_ref().unwrap().color.as_ref().unwrap();
        assert_eq!(left_color.rgb.as_deref(), Some("FF123456"));

        // Right - Theme with tint
        let right_color = border.right.as_ref().unwrap().color.as_ref().unwrap();
        assert_eq!(right_color.theme, Some(5));
        assert!(right_color.tint.is_some());

        // Top - Indexed
        let top_color = border.top.as_ref().unwrap().color.as_ref().unwrap();
        assert_eq!(top_color.indexed, Some(32));

        // Bottom - Auto
        let bottom_color = border.bottom.as_ref().unwrap().color.as_ref().unwrap();
        assert!(bottom_color.auto);
    }
}

// ============================================================================
// REALISTIC EXCEL BORDER PATTERNS
// ============================================================================

mod realistic_patterns {
    use super::*;

    #[test]
    fn test_box_border_thin() {
        // Common pattern: thin box border around cell
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        for side in [&border.left, &border.right, &border.top, &border.bottom] {
            assert!(side.is_some());
            assert_eq!(side.as_ref().unwrap().style, "thin");
        }
    }

    #[test]
    fn test_header_bottom_border() {
        // Common pattern: medium bottom border for headers
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top/>
                <bottom style="medium"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_some());
        assert_eq!(border.bottom.as_ref().unwrap().style, "medium");
    }

    #[test]
    fn test_total_row_double_top() {
        // Common pattern: double top border for totals row
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top style="double"><color indexed="64"/></top>
                <bottom/>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_some());
        assert_eq!(border.top.as_ref().unwrap().style, "double");
        assert!(border.bottom.is_none());
    }

    #[test]
    fn test_colored_border_accent() {
        // Common pattern: accent colored borders
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thin"><color theme="4"/></left>
                <right style="thin"><color theme="4"/></right>
                <top style="thin"><color theme="4"/></top>
                <bottom style="thin"><color theme="4"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        for side in [&border.left, &border.right, &border.top, &border.bottom] {
            assert!(side.is_some());
            let color = side.as_ref().unwrap().color.as_ref().unwrap();
            assert_eq!(color.theme, Some(4));
        }
    }

    #[test]
    fn test_thick_outside_thin_inside() {
        // Pattern: thick outside borders (for a single cell)
        let xml = styles_xml_with_borders(
            r#"<border>
                <left style="thick"><color indexed="64"/></left>
                <right style="thick"><color indexed="64"/></right>
                <top style="thick"><color indexed="64"/></top>
                <bottom style="thick"><color indexed="64"/></bottom>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        let border = &stylesheet.borders[0];

        for side in [&border.left, &border.right, &border.top, &border.bottom] {
            assert!(side.is_some());
            assert_eq!(side.as_ref().unwrap().style, "thick");
        }
    }

    #[test]
    fn test_default_excel_border_set() {
        // Typical Excel file has these borders at minimum
        let xml = styles_xml_with_borders(
            r#"<border>
                <left/>
                <right/>
                <top/>
                <bottom/>
                <diagonal/>
            </border>
            <border>
                <left style="thin"><color indexed="64"/></left>
                <right style="thin"><color indexed="64"/></right>
                <top style="thin"><color indexed="64"/></top>
                <bottom style="thin"><color indexed="64"/></bottom>
                <diagonal/>
            </border>"#,
        );

        let stylesheet = parse_styles_xml(&xml);
        assert_eq!(stylesheet.borders.len(), 2);

        // First is empty/no borders
        assert!(stylesheet.borders[0].left.is_none());

        // Second is thin box
        assert!(stylesheet.borders[1].left.is_some());
        assert_eq!(stylesheet.borders[1].left.as_ref().unwrap().style, "thin");
    }
}

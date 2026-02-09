//! Color resolution utilities
//!
//! Handles theme colors, indexed colors, RGB, and tint/shade calculations.

use crate::types::ColorSpec;

/// Excel's 64 indexed colors (legacy palette)
pub const INDEXED_COLORS: [&str; 64] = [
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/// Default theme colors (Office theme) used when no theme is present
/// Excel theme color indices (per ECMA-376):
/// 0: lt1 (Background 1 / light1) - typically white
/// 1: dk1 (Text 1 / dark1) - typically black
/// 2: lt2 (Background 2 / light2)
/// 3: dk2 (Text 2 / dark2)
/// 4-9: accent1-accent6
/// 10: hlink (hyperlink)
/// 11: folHlink (followed hyperlink)
pub const DEFAULT_THEME_COLORS: [&str; 12] = [
    "#FFFFFF", // 0: lt1 (Background 1 - light)
    "#000000", // 1: dk1 (Text 1 - dark)
    "#E7E6E6", // 2: lt2 (Background 2 - light)
    "#44546A", // 3: dk2 (Text 2 - dark)
    "#4472C4", // 4: accent1
    "#ED7D31", // 5: accent2
    "#A5A5A5", // 6: accent3
    "#FFC000", // 7: accent4
    "#5B9BD5", // 8: accent5
    "#70AD47", // 9: accent6
    "#0563C1", // 10: hlink
    "#954F72", // 11: folHlink
];

/// Resolve a `ColorSpec` to an #RRGGBB string
pub fn resolve_color(
    color: &ColorSpec,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<String> {
    // Priority: rgb > theme > indexed > auto

    if let Some(rgb) = &color.rgb {
        // Excel sometimes uses ARGB (8 chars), we want RGB (6 chars)
        let rgb = rgb.trim_start_matches('#');
        if rgb.len() == 8 {
            return Some(format!("#{}", &rgb[2..]));
        } else if rgb.len() == 6 {
            return Some(format!("#{rgb}"));
        }
        return Some(format!("#{rgb}"));
    }

    if let Some(theme_idx) = color.theme {
        let idx = theme_idx as usize;
        let base_color = theme_colors
            .get(idx)
            .map(String::as_str)
            .or_else(|| DEFAULT_THEME_COLORS.get(idx).copied())?;

        // Apply tint if present
        if let Some(tint) = color.tint {
            return Some(apply_tint(base_color, tint));
        }
        return Some(base_color.to_string());
    }

    if let Some(indexed) = color.indexed {
        if indexed == 64 {
            // 64 is "system foreground" - usually black
            return Some("#000000".to_string());
        }

        let idx = indexed as usize;

        // Use custom palette if available, otherwise fall back to default
        if let Some(custom_palette) = indexed_colors {
            if let Some(color) = custom_palette.get(idx) {
                return Some(color.clone());
            }
        }

        // Fall back to default palette
        if let Some(color) = INDEXED_COLORS.get(idx) {
            return Some((*color).to_string());
        }
    }

    if color.auto {
        // Auto color - default to black for text, white for background
        return Some("#000000".to_string());
    }

    None
}

/// Apply a tint value to a color
/// tint < 0: shade (darken)
/// tint > 0: tint (lighten)
#[allow(clippy::many_single_char_names)]
pub fn apply_tint(hex_color: &str, tint: f64) -> String {
    let hex = hex_color.trim_start_matches('#');

    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    let (h, s, l) = rgb_to_hsl(r, g, b);

    let new_l = if tint < 0.0 {
        // Shade: darken
        l * (1.0 + tint)
    } else {
        // Tint: lighten
        (1.0 - l).mul_add(tint, l)
    };

    let (r, g, b) = hsl_to_rgb(h, s, new_l.clamp(0.0, 1.0));

    format!("#{r:02X}{g:02X}{b:02X}")
}

/// Convert RGB to HSL
#[allow(clippy::many_single_char_names)]
fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let r = f64::from(r) / 255.0;
    let g = f64::from(g) / 255.0;
    let b = f64::from(b) / 255.0;

    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = f64::midpoint(max, min);

    if (max - min).abs() < f64::EPSILON {
        return (0.0, 0.0, l);
    }

    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };

    let h = if (max - r).abs() < f64::EPSILON {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if (max - g).abs() < f64::EPSILON {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };

    (h / 6.0, s, l)
}

/// Convert HSL to RGB
#[allow(clippy::many_single_char_names)]
#[allow(clippy::cast_possible_truncation)]
#[allow(clippy::cast_sign_loss)]
fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    if s.abs() < f64::EPSILON {
        let v = (l * 255.0).round() as u8;
        return (v, v, v);
    }

    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l.mul_add(-s, l + s)
    };
    let p = 2.0f64.mul_add(l, -q);

    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);

    (
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8,
    )
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }

    if t < 1.0 / 6.0 {
        return ((q - p) * 6.0).mul_add(t, p);
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return ((q - p) * (2.0 / 3.0 - t)).mul_add(6.0, p);
    }
    p
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
    fn test_tint_lighten() {
        // 50% tint on black should give gray
        let result = apply_tint("#000000", 0.5);
        assert_eq!(result, "#808080");
    }

    #[test]
    fn test_tint_darken() {
        // 50% shade on white should give gray
        let result = apply_tint("#FFFFFF", -0.5);
        assert_eq!(result, "#808080");
    }
}

// =============================================================================
// Fill color resolution tests
// =============================================================================

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::excessive_precision,
    clippy::panic
)]
mod fill_color_tests {
    use super::*;

    fn default_theme_colors() -> Vec<String> {
        vec![
            "#FFFFFF".to_string(), // 0: lt1 (Background 1)
            "#000000".to_string(), // 1: dk1 (Text 1)
            "#E7E6E6".to_string(), // 2: lt2 (Background 2)
            "#44546A".to_string(), // 3: dk2 (Text 2)
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

    // =========================================================================
    // RGB color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_rgb_color_argb_format() {
        // ARGB format: first 2 chars are alpha (FF = opaque)
        let color = ColorSpec {
            rgb: Some("FFFFFF00".to_string()), // Yellow with alpha
            theme: None,
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFF00".to_string())); // Alpha stripped
    }

    #[test]
    fn test_resolve_rgb_color_rgb_format() {
        // Some files use 6-char RGB without alpha
        let color = ColorSpec {
            rgb: Some("FF0000".to_string()), // Red
            theme: None,
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string()));
    }

    #[test]
    fn test_resolve_rgb_color_with_hash() {
        // Handle color with leading hash
        let color = ColorSpec {
            rgb: Some("#00FF00".to_string()), // Green with hash
            theme: None,
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#00FF00".to_string()));
    }

    #[test]
    fn test_resolve_rgb_common_fill_colors() {
        // Test common fill colors
        let test_cases = [
            ("FFFF0000", "#FF0000"), // Red
            ("FF00FF00", "#00FF00"), // Green
            ("FF0000FF", "#0000FF"), // Blue
            ("FFFFFF00", "#FFFF00"), // Yellow
            ("FFFF00FF", "#FF00FF"), // Magenta
            ("FF00FFFF", "#00FFFF"), // Cyan
            ("FFFFFFFF", "#FFFFFF"), // White
            ("FF000000", "#000000"), // Black
            ("FFC0C0C0", "#C0C0C0"), // Silver
            ("FF808080", "#808080"), // Gray
        ];

        for (input, expected) in test_cases {
            let color = ColorSpec {
                rgb: Some(input.to_string()),
                theme: None,
                tint: None,
                indexed: None,
                auto: false,
            };

            let resolved = resolve_color(&color, &default_theme_colors(), None);
            assert_eq!(
                resolved,
                Some(expected.to_string()),
                "Failed for input: {input}"
            );
        }
    }

    // =========================================================================
    // Theme color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_theme_color_dark1() {
        // theme="1" is dk1 (Text 1 - usually black)
        let color = ColorSpec {
            rgb: None,
            theme: Some(1),
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_light1() {
        // theme="0" is lt1 (Background 1 - usually white)
        let color = ColorSpec {
            rgb: None,
            theme: Some(0),
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFFFF".to_string()));
    }

    #[test]
    fn test_resolve_theme_color_accent1() {
        // theme="4" is accent1 (blue in default Office theme)
        let color = ColorSpec {
            rgb: None,
            theme: Some(4),
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#4472C4".to_string()));
    }

    #[test]
    fn test_resolve_all_theme_colors() {
        let theme_colors = default_theme_colors();
        let expected = [
            (0, "#FFFFFF"),  // lt1 (Background 1)
            (1, "#000000"),  // dk1 (Text 1)
            (2, "#E7E6E6"),  // lt2 (Background 2)
            (3, "#44546A"),  // dk2 (Text 2)
            (4, "#4472C4"),  // accent1
            (5, "#ED7D31"),  // accent2
            (6, "#A5A5A5"),  // accent3
            (7, "#FFC000"),  // accent4
            (8, "#5B9BD5"),  // accent5
            (9, "#70AD47"),  // accent6
            (10, "#0563C1"), // hlink
            (11, "#954F72"), // folHlink
        ];

        for (theme_idx, expected_color) in expected {
            let color = ColorSpec {
                rgb: None,
                theme: Some(theme_idx),
                tint: None,
                indexed: None,
                auto: false,
            };

            let resolved = resolve_color(&color, &theme_colors, None);
            assert_eq!(
                resolved,
                Some(expected_color.to_string()),
                "Failed for theme index: {theme_idx}"
            );
        }
    }

    // =========================================================================
    // Theme color with tint tests
    // =========================================================================

    #[test]
    fn test_resolve_theme_color_with_positive_tint() {
        // Positive tint lightens the color
        let color = ColorSpec {
            rgb: None,
            theme: Some(4), // accent1 = #4472C4
            tint: Some(0.5),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
        let hex = resolved.expect("resolved color");

        // Lightened color should have higher RGB values
        // Original: #4472C4 (R=68, G=114, B=196)
        // After 50% tint should be lighter
        assert!(hex.starts_with('#'));
        assert_eq!(hex.len(), 7);
    }

    #[test]
    fn test_resolve_theme_color_with_negative_tint() {
        // Negative tint darkens the color (shade)
        let color = ColorSpec {
            rgb: None,
            theme: Some(4), // accent1 = #4472C4
            tint: Some(-0.5),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
        let hex = resolved.expect("resolved color");

        // Darkened color should have lower RGB values
        assert!(hex.starts_with('#'));
        assert_eq!(hex.len(), 7);
    }

    #[test]
    fn test_resolve_theme_color_with_small_tint() {
        // Small tint like Excel often uses
        let color = ColorSpec {
            rgb: None,
            theme: Some(4),
            tint: Some(0.39997558519241921),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
    }

    #[test]
    fn test_resolve_theme_color_with_small_shade() {
        // Small shade like Excel often uses
        let color = ColorSpec {
            rgb: None,
            theme: Some(4),
            tint: Some(-0.249977111117893),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_some());
    }

    #[test]
    fn test_tint_on_black_produces_gray() {
        // 50% tint on black should give 50% gray
        let color = ColorSpec {
            rgb: None,
            theme: Some(1), // dk1 = #000000 (Text 1)
            tint: Some(0.5),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#808080".to_string()));
    }

    #[test]
    fn test_shade_on_white_produces_gray() {
        // 50% shade on white should give 50% gray
        let color = ColorSpec {
            rgb: None,
            theme: Some(0), // lt1 = #FFFFFF (Background 1)
            tint: Some(-0.5),
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#808080".to_string()));
    }

    // =========================================================================
    // Indexed color resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_indexed_color_black() {
        // indexed="0" is black
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(0),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_white() {
        // indexed="1" is white
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(1),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFFFF".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_red() {
        // indexed="2" is red
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(2),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_yellow() {
        // indexed="5" is yellow
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(5),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FFFF00".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_8() {
        // indexed="8" is black (second occurrence in palette)
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(8),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_color_64_system_foreground() {
        // indexed="64" is special "system foreground" color
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(64),
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    #[test]
    fn test_resolve_indexed_colors_common() {
        // Test common indexed colors used in fills
        let test_cases = [
            (0, "#000000"),  // Black
            (1, "#FFFFFF"),  // White
            (2, "#FF0000"),  // Red
            (3, "#00FF00"),  // Green
            (4, "#0000FF"),  // Blue
            (5, "#FFFF00"),  // Yellow
            (22, "#C0C0C0"), // Silver
            (23, "#808080"), // Gray
        ];

        for (indexed, expected) in test_cases {
            let color = ColorSpec {
                rgb: None,
                theme: None,
                tint: None,
                indexed: Some(indexed),
                auto: false,
            };

            let resolved = resolve_color(&color, &default_theme_colors(), None);
            assert_eq!(
                resolved,
                Some(expected.to_string()),
                "Failed for indexed: {indexed}"
            );
        }
    }

    // =========================================================================
    // Auto color tests
    // =========================================================================

    #[test]
    fn test_resolve_auto_color() {
        // auto="1" means use automatic color (defaults to black)
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: None,
            auto: true,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#000000".to_string()));
    }

    // =========================================================================
    // Color priority tests
    // =========================================================================

    #[test]
    fn test_color_priority_rgb_over_theme() {
        // RGB should take priority over theme if both are specified
        let color = ColorSpec {
            rgb: Some("FF0000".to_string()),
            theme: Some(4),
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string())); // RGB wins
    }

    #[test]
    fn test_color_priority_rgb_over_indexed() {
        // RGB should take priority over indexed
        let color = ColorSpec {
            rgb: Some("00FF00".to_string()),
            theme: None,
            tint: None,
            indexed: Some(2), // Red
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#00FF00".to_string())); // RGB wins
    }

    #[test]
    fn test_color_priority_theme_over_indexed() {
        // Theme should take priority over indexed
        let color = ColorSpec {
            rgb: None,
            theme: Some(4), // accent1 = #4472C4
            tint: None,
            indexed: Some(2), // Red
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#4472C4".to_string())); // Theme wins
    }

    #[test]
    fn test_color_priority_indexed_over_auto() {
        // Indexed should take priority over auto
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(2), // Red
            auto: true,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert_eq!(resolved, Some("#FF0000".to_string())); // Indexed wins
    }

    // =========================================================================
    // Edge case tests
    // =========================================================================

    #[test]
    fn test_empty_color_spec() {
        // No color specified should return None
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_none());
    }

    #[test]
    fn test_invalid_theme_index() {
        // Theme index beyond available colors
        let color = ColorSpec {
            rgb: None,
            theme: Some(100), // Invalid
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_none());
    }

    #[test]
    fn test_invalid_indexed_color() {
        // Indexed color beyond 64-color palette (but not 64)
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(100), // Invalid
            auto: false,
        };

        let resolved = resolve_color(&color, &default_theme_colors(), None);
        assert!(resolved.is_none());
    }

    #[test]
    fn test_empty_theme_colors_uses_defaults() {
        // When theme_colors is empty, should use DEFAULT_THEME_COLORS
        let color = ColorSpec {
            rgb: None,
            theme: Some(4), // accent1
            tint: None,
            indexed: None,
            auto: false,
        };

        let resolved = resolve_color(&color, &[], None);
        // Should fall back to DEFAULT_THEME_COLORS
        assert_eq!(resolved, Some("#4472C4".to_string()));
    }

    // =========================================================================
    // Custom indexed color tests
    // =========================================================================

    #[test]
    fn test_custom_indexed_colors() {
        // Test with default palette (None)
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(2), // Red in default palette
            auto: false,
        };

        let resolved = resolve_color(&color, &[], None);
        assert_eq!(resolved, Some("#FF0000".to_string()));

        // Test with custom palette
        let custom_palette = vec![
            "#000000".to_string(),
            "#FFFFFF".to_string(),
            "#00FF00".to_string(), // Custom green instead of red at index 2
        ];

        let resolved_custom = resolve_color(&color, &[], Some(&custom_palette));
        assert_eq!(resolved_custom, Some("#00FF00".to_string()));
    }

    #[test]
    fn test_fallback_to_default_when_custom_too_short() {
        let color = ColorSpec {
            rgb: None,
            theme: None,
            tint: None,
            indexed: Some(10), // Index beyond custom palette
            auto: false,
        };

        // Custom palette with only 3 colors
        let custom_palette = vec![
            "#000000".to_string(),
            "#FFFFFF".to_string(),
            "#00FF00".to_string(),
        ];

        // Should fall back to default palette
        let resolved = resolve_color(&color, &[], Some(&custom_palette));
        assert!(resolved.is_some());
    }
}

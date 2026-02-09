//! Style and value resolution - resolves cell values, styles, and borders.

use crate::color::resolve_color;
use crate::numfmt::{format_number_compiled, CompiledFormat};
use crate::types::{
    Border, BorderStyle, CellRawValue, CellType, GradientFill, GradientStop, HAlign, PatternType,
    RawBorderSide, Style, StyleSheet, Theme, VAlign,
};

use super::worksheet::CellTypeTag;
use super::{now_ms, NumFmtInfo, ParseOptions, SheetParseMetrics};

/// Resolve cell value and type
#[allow(clippy::needless_option_as_deref, clippy::too_many_arguments)]
pub(super) fn resolve_cell_value(
    raw_value: Option<&str>,
    cell_type: CellTypeTag,
    shared_strings: &[String],
    style_idx: Option<u32>,
    numfmt_lookup: &[NumFmtInfo],
    date1904: bool,
    options: ParseOptions,
    mut metrics: Option<&mut SheetParseMetrics>,
) -> (
    Option<String>,
    Option<CellRawValue>,
    Option<String>,
    CellType,
) {
    match cell_type {
        CellTypeTag::Shared => {
            // Shared string
            let idx: usize = raw_value.and_then(|v| v.parse().ok()).unwrap_or(0);
            if options.eager_values {
                let value = shared_strings.get(idx).cloned();
                (value, None, None, CellType::String)
            } else {
                let idx_u32 = u32::try_from(idx).unwrap_or(0);
                (
                    None,
                    Some(CellRawValue::SharedString(idx_u32)),
                    None,
                    CellType::String,
                )
            }
        }
        CellTypeTag::Str | CellTypeTag::Inline => {
            // Inline string
            let value = raw_value.map(ToString::to_string);
            if options.eager_values {
                (value, None, None, CellType::String)
            } else {
                (
                    None,
                    value.map(CellRawValue::String),
                    None,
                    CellType::String,
                )
            }
        }
        CellTypeTag::Bool => {
            // Boolean
            let bool_value = match raw_value {
                Some("1" | "true") => Some(true),
                Some("0" | "false") => Some(false),
                _ => None,
            };
            if options.eager_values {
                let value = match bool_value {
                    Some(true) => Some("TRUE".to_string()),
                    Some(false) => Some("FALSE".to_string()),
                    None => raw_value.map(ToString::to_string),
                };
                (value, None, None, CellType::Boolean)
            } else if let Some(value) = bool_value {
                (
                    None,
                    Some(CellRawValue::Boolean(value)),
                    None,
                    CellType::Boolean,
                )
            } else {
                (
                    None,
                    raw_value.map(|v| CellRawValue::String(v.to_string())),
                    None,
                    CellType::Boolean,
                )
            }
        }
        CellTypeTag::Error => {
            // Error
            let value = raw_value.map(ToString::to_string);
            if options.eager_values {
                (value, None, None, CellType::Error)
            } else {
                (None, value.map(CellRawValue::Error), None, CellType::Error)
            }
        }
        CellTypeTag::Default => {
            // Number or date (default)
            let Some(v) = raw_value else {
                return (None, None, None, CellType::String);
            };

            let parsed = if let Some(m) = metrics.as_deref_mut() {
                m.value_parse_calls = m.value_parse_calls.saturating_add(1);
                let parse_start = now_ms();
                let parsed = v.parse::<f64>();
                m.value_parse_ms += now_ms() - parse_start;
                parsed
            } else {
                v.parse::<f64>()
            };

            let Ok(num) = parsed else {
                let value = v.to_string();
                if options.eager_values {
                    return (Some(value), None, None, CellType::String);
                }
                return (
                    None,
                    Some(CellRawValue::String(value)),
                    None,
                    CellType::String,
                );
            };

            let info = style_idx.and_then(|idx| numfmt_lookup.get(idx as usize));

            if let Some(m) = metrics.as_deref_mut() {
                if let Some(info) = info {
                    if info.is_builtin {
                        m.numfmt_builtin = m.numfmt_builtin.saturating_add(1);
                    } else if info.is_custom {
                        m.numfmt_custom = m.numfmt_custom.saturating_add(1);
                    }
                    if info.is_general {
                        m.numfmt_general = m.numfmt_general.saturating_add(1);
                    }
                } else {
                    m.numfmt_general = m.numfmt_general.saturating_add(1);
                }
            }

            if let Some(info) = info {
                match &info.compiled {
                    CompiledFormat::General => {
                        let raw_string = v.to_string();
                        if options.eager_values {
                            (Some(raw_string), None, None, CellType::Number)
                        } else {
                            (
                                None,
                                Some(CellRawValue::Number(num)),
                                Some(raw_string),
                                CellType::Number,
                            )
                        }
                    }
                    CompiledFormat::Date(_) => {
                        if options.eager_values {
                            if let Some(m) = metrics.as_deref_mut() {
                                m.format_number_calls = m.format_number_calls.saturating_add(1);
                                m.format_number_date_calls =
                                    m.format_number_date_calls.saturating_add(1);
                            }
                            let format_start = now_ms();
                            let formatted = format_number_compiled(num, &info.compiled, date1904);
                            if let Some(m) = metrics.as_deref_mut() {
                                let elapsed = now_ms() - format_start;
                                m.format_number_ms += elapsed;
                                m.format_number_date_ms += elapsed;
                            }
                            (Some(formatted), None, None, CellType::Date)
                        } else {
                            (None, Some(CellRawValue::Date(num)), None, CellType::Date)
                        }
                    }
                    _ => {
                        if options.eager_values {
                            if let Some(m) = metrics.as_deref_mut() {
                                m.format_number_calls = m.format_number_calls.saturating_add(1);
                                m.format_number_number_calls =
                                    m.format_number_number_calls.saturating_add(1);
                            }
                            let format_start = now_ms();
                            let formatted = format_number_compiled(num, &info.compiled, date1904);
                            if let Some(m) = metrics.as_deref_mut() {
                                let elapsed = now_ms() - format_start;
                                m.format_number_ms += elapsed;
                                m.format_number_number_ms += elapsed;
                            }
                            (Some(formatted), None, None, CellType::Number)
                        } else {
                            (
                                None,
                                Some(CellRawValue::Number(num)),
                                None,
                                CellType::Number,
                            )
                        }
                    }
                }
            } else if options.eager_values {
                (Some(v.to_string()), None, None, CellType::Number)
            } else {
                (
                    None,
                    Some(CellRawValue::Number(num)),
                    Some(v.to_string()),
                    CellType::Number,
                )
            }
        }
    }
}

/// Get the default style (just the default font) for cells without explicit styling
pub(super) fn get_default_style(stylesheet: &StyleSheet, theme: &Theme) -> Option<Style> {
    let default_font = stylesheet.default_font.as_ref()?;
    let indexed_colors = stylesheet.indexed_colors.as_ref();
    let theme_colors = &theme.colors;

    let mut style = Style::default();

    // Resolve font family: use theme font if scheme is specified
    let font_family = match default_font.scheme.as_deref() {
        Some("minor") => theme.minor_font.as_ref().or(default_font.name.as_ref()),
        Some("major") => theme.major_font.as_ref().or(default_font.name.as_ref()),
        _ => default_font.name.as_ref(),
    };

    if let Some(family) = font_family {
        style.font_family = Some(family.clone());
    }

    style.font_size = default_font.size;
    style.bold = if default_font.bold { Some(true) } else { None };
    style.italic = if default_font.italic {
        Some(true)
    } else {
        None
    };
    style.underline = default_font.underline;
    style.strikethrough = if default_font.strikethrough {
        Some(true)
    } else {
        None
    };
    style.vert_align = default_font.vert_align;

    if let Some(ref color) = default_font.color {
        style.font_color = resolve_color(color, theme_colors, indexed_colors);
    }

    Some(style)
}

/// Resolve a style index to a full Style object
pub(super) fn resolve_style(idx: u32, stylesheet: &StyleSheet, theme: &Theme) -> Option<Style> {
    let xf = stylesheet.cell_xfs.get(idx as usize)?;
    let mut style = Style::default();

    let indexed_colors = stylesheet.indexed_colors.as_ref();
    let theme_colors = &theme.colors;

    // BUGFIX BUG-002: Implement inheritance from cellStyleXfs
    // If this cellXf has an xfId, get the parent cellStyleXf for inheritance
    let parent_xf = xf
        .xf_id
        .and_then(|xf_id| stylesheet.cell_style_xfs.get(xf_id as usize));

    // Start with the default font from the "Normal" style (applies to all cells)
    // This provides the baseline font family, size, and color for the workbook
    if let Some(ref default_font) = stylesheet.default_font {
        // Resolve font family: use theme font if scheme is specified
        let font_family = match default_font.scheme.as_deref() {
            Some("minor") => theme.minor_font.as_ref().or(default_font.name.as_ref()),
            Some("major") => theme.major_font.as_ref().or(default_font.name.as_ref()),
            _ => default_font.name.as_ref(),
        };

        if let Some(family) = font_family {
            style.font_family = Some(family.clone());
        }

        style.font_size = default_font.size;
        style.bold = if default_font.bold { Some(true) } else { None };
        style.italic = if default_font.italic {
            Some(true)
        } else {
            None
        };
        style.underline = default_font.underline;
        style.strikethrough = if default_font.strikethrough {
            Some(true)
        } else {
            None
        };
        style.vert_align = default_font.vert_align;

        if let Some(ref color) = default_font.color {
            style.font_color = resolve_color(color, theme_colors, indexed_colors);
        }
    }

    // Font - inherit from parent if applyFont is false, otherwise use child
    // Per ECMA-376, if applyFont is true, use the cellXf's fontId
    // If applyFont is false (or not set), inherit from the parent cellStyleXf's fontId
    let font_id = if xf.apply_font {
        xf.font_id
    } else {
        parent_xf.and_then(|p| p.font_id).or(xf.font_id)
    };

    if let Some(font_id) = font_id {
        if let Some(font) = stylesheet.fonts.get(font_id as usize) {
            // Resolve font family: use theme font if scheme is specified
            let font_family = match font.scheme.as_deref() {
                Some("minor") => theme.minor_font.as_ref().or(font.name.as_ref()),
                Some("major") => theme.major_font.as_ref().or(font.name.as_ref()),
                _ => font.name.as_ref(),
            };

            if let Some(family) = font_family {
                style.font_family = Some(family.clone());
            }

            // Only override properties that are explicitly set in this font
            // This preserves defaults/inherited values for unset properties
            if font.size.is_some() {
                style.font_size = font.size;
            }
            if font.bold {
                style.bold = Some(true);
            }
            if font.italic {
                style.italic = Some(true);
            }
            if font.underline.is_some() {
                style.underline = font.underline;
            }
            if font.strikethrough {
                style.strikethrough = Some(true);
            }
            if font.vert_align.is_some() {
                style.vert_align = font.vert_align;
            }

            if let Some(ref color) = font.color {
                style.font_color = resolve_color(color, theme_colors, indexed_colors);
            }
        }
    }

    // Fill - inherit from parent if applyFill is false
    let fill_id = if xf.apply_fill {
        xf.fill_id
    } else {
        parent_xf.and_then(|p| p.fill_id).or(xf.fill_id)
    };

    if let Some(fill_id) = fill_id {
        if let Some(fill) = stylesheet.fills.get(fill_id as usize) {
            // Check for gradient fill first
            if let Some(ref gradient) = fill.gradient {
                // Resolve gradient fill
                let resolved_stops: Vec<GradientStop> = gradient
                    .stops
                    .iter()
                    .map(|stop| GradientStop {
                        position: stop.position,
                        color: resolve_color(&stop.color, theme_colors, indexed_colors)
                            .unwrap_or_else(|| "#000000".to_string()),
                    })
                    .collect();

                style.gradient = Some(GradientFill {
                    gradient_type: gradient
                        .gradient_type
                        .clone()
                        .unwrap_or_else(|| "linear".to_string()),
                    degree: gradient.degree,
                    left: gradient.left,
                    right: gradient.right,
                    top: gradient.top,
                    bottom: gradient.bottom,
                    stops: resolved_stops,
                });
            } else {
                // Pattern fill handling
                let pattern_type = parse_pattern_type(fill.pattern_type.as_deref());

                match pattern_type {
                    Some(PatternType::None) | None => {
                        // No fill - leave bg_color as None
                    }
                    Some(PatternType::Solid) => {
                        // Solid fill: fgColor is the background color
                        if let Some(ref color) = fill.fg_color {
                            style.bg_color = resolve_color(color, theme_colors, indexed_colors);
                        }
                    }
                    Some(pt) => {
                        // Pattern fill: set pattern type, fg_color (pattern), and bg_color (background)
                        style.pattern_type = Some(pt);
                        if let Some(ref color) = fill.fg_color {
                            style.fg_color = resolve_color(color, theme_colors, indexed_colors);
                        }
                        if let Some(ref color) = fill.bg_color {
                            style.bg_color = resolve_color(color, theme_colors, indexed_colors);
                        }
                    }
                }
            }
        }
    }

    // Border - inherit from parent if applyBorder is false
    let border_id = if xf.apply_border {
        xf.border_id
    } else {
        parent_xf.and_then(|p| p.border_id).or(xf.border_id)
    };

    if let Some(border_id) = border_id {
        if let Some(border) = stylesheet.borders.get(border_id as usize) {
            style.border_top = resolve_border(border.top.as_ref(), theme_colors, indexed_colors);
            style.border_right =
                resolve_border(border.right.as_ref(), theme_colors, indexed_colors);
            style.border_bottom =
                resolve_border(border.bottom.as_ref(), theme_colors, indexed_colors);
            style.border_left = resolve_border(border.left.as_ref(), theme_colors, indexed_colors);
            style.border_diagonal =
                resolve_border(border.diagonal.as_ref(), theme_colors, indexed_colors);
            style.diagonal_up = if border.diagonal_up { Some(true) } else { None };
            style.diagonal_down = if border.diagonal_down {
                Some(true)
            } else {
                None
            };
        }
    }

    // Alignment - inherit from parent if applyAlignment is false
    let alignment = if xf.apply_alignment {
        xf.alignment.as_ref()
    } else {
        parent_xf
            .and_then(|p| p.alignment.as_ref())
            .or(xf.alignment.as_ref())
    };

    if let Some(align) = alignment {
        style.align_h = match align.horizontal.as_deref() {
            Some("left") => Some(HAlign::Left),
            Some("center") => Some(HAlign::Center),
            Some("right") => Some(HAlign::Right),
            Some("justify") => Some(HAlign::Justify),
            Some("fill") => Some(HAlign::Fill),
            Some("general") => Some(HAlign::General),
            Some("centerContinuous") => Some(HAlign::CenterContinuous),
            Some("distributed") => Some(HAlign::Distributed),
            _ => None,
        };

        style.align_v = match align.vertical.as_deref() {
            Some("top") => Some(VAlign::Top),
            Some("center") => Some(VAlign::Center),
            Some("bottom") => Some(VAlign::Bottom),
            Some("justify") => Some(VAlign::Justify),
            Some("distributed") => Some(VAlign::Distributed),
            _ => None,
        };

        style.wrap = if align.wrap_text { Some(true) } else { None };
        style.indent = align.indent;
        style.rotation = align.text_rotation;
        style.shrink_to_fit = if align.shrink_to_fit {
            Some(true)
        } else {
            None
        };
    }

    // Protection - inherit from parent if applyProtection is false
    let protection = if xf.apply_protection {
        xf.protection.as_ref()
    } else {
        parent_xf
            .and_then(|p| p.protection.as_ref())
            .or(xf.protection.as_ref())
    };

    // Note: In Excel, cells are locked by default (locked=true).
    // We only set locked/hidden when there's explicit protection info.
    if let Some(prot) = protection {
        // Only include locked if it's explicitly set to false (unlocked)
        // since locked=true is the default
        if !prot.locked {
            style.locked = Some(false);
        }
        // Only include hidden if it's true (formulas hidden)
        if prot.hidden {
            style.hidden = Some(true);
        }
    }

    // Only return style if it has any properties set
    if style.font_family.is_some()
        || style.font_size.is_some()
        || style.font_color.is_some()
        || style.bold.is_some()
        || style.italic.is_some()
        || style.underline.is_some()
        || style.strikethrough.is_some()
        || style.vert_align.is_some()
        || style.bg_color.is_some()
        || style.pattern_type.is_some()
        || style.fg_color.is_some()
        || style.border_top.is_some()
        || style.border_right.is_some()
        || style.border_bottom.is_some()
        || style.border_left.is_some()
        || style.border_diagonal.is_some()
        || style.diagonal_up.is_some()
        || style.diagonal_down.is_some()
        || style.align_h.is_some()
        || style.align_v.is_some()
        || style.wrap.is_some()
        || style.indent.is_some()
        || style.rotation.is_some()
        || style.shrink_to_fit.is_some()
        || style.locked.is_some()
        || style.hidden.is_some()
    {
        Some(style)
    } else {
        None
    }
}

/// Parse a pattern type string into a PatternType enum
fn parse_pattern_type(pattern_type: Option<&str>) -> Option<PatternType> {
    match pattern_type? {
        "none" => Some(PatternType::None),
        "solid" => Some(PatternType::Solid),
        "gray125" => Some(PatternType::Gray125),
        "gray0625" => Some(PatternType::Gray0625),
        "darkGray" => Some(PatternType::DarkGray),
        "mediumGray" => Some(PatternType::MediumGray),
        "lightGray" => Some(PatternType::LightGray),
        "darkHorizontal" => Some(PatternType::DarkHorizontal),
        "darkVertical" => Some(PatternType::DarkVertical),
        "darkDown" => Some(PatternType::DarkDown),
        "darkUp" => Some(PatternType::DarkUp),
        "darkGrid" => Some(PatternType::DarkGrid),
        "darkTrellis" => Some(PatternType::DarkTrellis),
        "lightHorizontal" => Some(PatternType::LightHorizontal),
        "lightVertical" => Some(PatternType::LightVertical),
        "lightDown" => Some(PatternType::LightDown),
        "lightUp" => Some(PatternType::LightUp),
        "lightGrid" => Some(PatternType::LightGrid),
        "lightTrellis" => Some(PatternType::LightTrellis),
        _ => None,
    }
}

/// Resolve a border side
fn resolve_border(
    side: Option<&RawBorderSide>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<Border> {
    let side = side?;

    let style = match side.style.as_str() {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => return None,
    };

    let color = side
        .color
        .as_ref()
        .and_then(|c| resolve_color(c, theme_colors, indexed_colors))
        .unwrap_or_else(|| "#000000".to_string());

    Some(Border { style, color })
}

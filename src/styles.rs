//! Parsing of xl/styles.xml
//!
//! This file contains fonts, fills, borders, and cell formats (xf).

use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::BufRead;

use crate::error::Result;
use crate::types::{
    CellXf, ColorSpec, DxfStyle, NamedStyle, RawAlignment, RawBorder, RawBorderSide, RawFill,
    RawFont, RawGradientFill, RawGradientStop, RawProtection, StyleSheet, UnderlineStyle,
    VertAlign,
};
use crate::xml_helpers::parse_color_attrs;

/// Parse styles.xml content
#[allow(clippy::too_many_lines)]
#[allow(clippy::cognitive_complexity)]
pub fn parse_styles<R: BufRead>(reader: R) -> Result<StyleSheet> {
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut stylesheet = StyleSheet::default();
    let mut buf = Vec::new();

    // State tracking
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;
    let mut in_cell_styles = false;
    let mut in_num_fmts = false;
    let mut in_colors = false;
    let mut in_indexed_colors = false;
    let mut in_dxfs = false;
    let mut in_dxf = false;

    let mut current_font: Option<RawFont> = None;
    let mut current_fill: Option<RawFill> = None;
    let mut current_border: Option<RawBorder> = None;
    let mut current_xf: Option<CellXf> = None;
    let mut current_border_side: Option<String> = None;
    let mut indexed_colors: Vec<String> = Vec::new();
    let mut current_dxf: Option<DxfStyle> = None;
    let mut current_gradient: Option<RawGradientFill> = None;
    let mut in_gradient_stop = false;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(ref event @ (Event::Start(ref e) | Event::Empty(ref e))) => {
                let is_empty = matches!(event, Event::Empty(_));
                let name = e.local_name();
                let name_str = std::str::from_utf8(name.as_ref()).unwrap_or("");

                match name_str {
                    "numFmts" => in_num_fmts = true,
                    "fonts" => in_fonts = true,
                    "fills" => in_fills = true,
                    "borders" => in_borders = true,
                    "cellXfs" => in_cell_xfs = true,
                    "cellStyleXfs" => in_cell_style_xfs = true,
                    "cellStyles" => in_cell_styles = true,
                    "colors" => in_colors = true,
                    "indexedColors" if in_colors => in_indexed_colors = true,
                    "dxfs" => in_dxfs = true,
                    "dxf" if in_dxfs => {
                        in_dxf = true;
                        current_dxf = Some(DxfStyle::default());
                    }

                    "rgbColor" if in_indexed_colors => {
                        // Parse <rgbColor rgb="FFRRGGBB"/>
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"rgb" {
                                if let Ok(rgb) = std::str::from_utf8(&attr.value) {
                                    // Excel uses ARGB format (8 chars), strip alpha to get RGB
                                    let color = if rgb.len() == 8 {
                                        format!("#{}", &rgb[2..])
                                    } else {
                                        format!("#{rgb}")
                                    };
                                    indexed_colors.push(color);
                                }
                            }
                        }
                    }

                    "numFmt" if in_num_fmts => {
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    id = std::str::from_utf8(&attr.value)
                                        .unwrap_or("0")
                                        .parse()
                                        .unwrap_or(0);
                                }
                                b"formatCode" => {
                                    code =
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                }
                                _ => {}
                            }
                        }
                        stylesheet.num_fmts.push((id, code));
                    }

                    "font" if in_fonts => {
                        current_font = Some(RawFont::default());
                    }

                    "sz" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.size = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }

                    "name" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.name = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                            }
                        }
                    }

                    "scheme" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.scheme = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                            }
                        }
                    }

                    "b" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.bold = true;
                        }
                    }

                    "i" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.italic = true;
                        }
                    }

                    "u" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            // Parse underline style from val attribute
                            // Default to Single if no val attribute (e.g., <u/>)
                            let mut underline_style = UnderlineStyle::Single;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    underline_style = match std::str::from_utf8(&attr.value)
                                        .unwrap_or("single")
                                    {
                                        "single" => UnderlineStyle::Single,
                                        "double" => UnderlineStyle::Double,
                                        "singleAccounting" => UnderlineStyle::SingleAccounting,
                                        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
                                        "none" => UnderlineStyle::None,
                                        _ => UnderlineStyle::Single,
                                    };
                                }
                            }
                            font.underline = Some(underline_style);
                        }
                    }

                    "vertAlign" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    font.vert_align = match std::str::from_utf8(&attr.value)
                                        .unwrap_or("baseline")
                                    {
                                        "subscript" => Some(VertAlign::Subscript),
                                        "superscript" => Some(VertAlign::Superscript),
                                        "baseline" => Some(VertAlign::Baseline),
                                        _ => None,
                                    };
                                }
                            }
                        }
                    }

                    "strike" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.strikethrough = true;
                        }
                    }

                    "color" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.color = Some(parse_color_attrs(e));
                        }
                    }

                    "fill" if in_fills => {
                        current_fill = Some(RawFill::default());
                    }

                    "patternFill" if current_fill.is_some() => {
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"patternType" {
                                    fill.pattern_type = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                            }
                        }
                    }

                    "gradientFill" if current_fill.is_some() => {
                        // Parse gradient fill attributes
                        let mut gradient = RawGradientFill::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"type" => {
                                    gradient.gradient_type = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                                b"degree" => {
                                    gradient.degree = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"left" => {
                                    gradient.left = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"right" => {
                                    gradient.right = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"top" => {
                                    gradient.top = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"bottom" => {
                                    gradient.bottom = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }

                        // Default to "linear" if no type specified
                        if gradient.gradient_type.is_none() {
                            gradient.gradient_type = Some("linear".to_string());
                        }

                        current_gradient = Some(gradient);
                    }

                    "stop" if current_gradient.is_some() => {
                        // Parse stop position
                        in_gradient_stop = true;
                        let mut position = 0.0;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"position" {
                                position = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0.0);
                            }
                        }
                        // We'll store position temporarily and add the color when we see it
                        // Use a placeholder stop that will be filled in when we see the color
                        if let Some(ref mut gradient) = current_gradient {
                            gradient.stops.push(RawGradientStop {
                                position,
                                color: ColorSpec {
                                    rgb: None,
                                    theme: None,
                                    tint: None,
                                    indexed: None,
                                    auto: false,
                                },
                            });
                        }
                    }

                    "color" if in_gradient_stop && current_gradient.is_some() => {
                        // Parse color within a gradient stop
                        let color = parse_color_attrs(e);
                        if let Some(ref mut gradient) = current_gradient {
                            if let Some(last_stop) = gradient.stops.last_mut() {
                                last_stop.color = color;
                            }
                        }
                    }

                    "fgColor" if current_fill.is_some() && current_gradient.is_none() => {
                        if let Some(ref mut fill) = current_fill {
                            fill.fg_color = Some(parse_color_attrs(e));
                        }
                    }

                    "bgColor" if current_fill.is_some() && current_gradient.is_none() => {
                        if let Some(ref mut fill) = current_fill {
                            fill.bg_color = Some(parse_color_attrs(e));
                        }
                    }

                    "border" if in_borders => {
                        let mut border = RawBorder::default();

                        // Parse diagonalUp and diagonalDown attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"diagonalUp" => {
                                    border.diagonal_up =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"diagonalDown" => {
                                    border.diagonal_down =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                _ => {}
                            }
                        }

                        // For self-closing <border/> tags, add immediately
                        if is_empty {
                            stylesheet.borders.push(border);
                        } else {
                            current_border = Some(border);
                        }
                    }

                    "left" | "right" | "top" | "bottom" | "diagonal"
                        if current_border.is_some() =>
                    {
                        current_border_side = Some(name_str.to_string());

                        // Get style attribute
                        let mut style = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                style = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                        }

                        if !style.is_empty() {
                            let side = RawBorderSide { style, color: None };
                            if let Some(ref mut border) = current_border {
                                match name_str {
                                    "left" => border.left = Some(side),
                                    "right" => border.right = Some(side),
                                    "top" => border.top = Some(side),
                                    "bottom" => border.bottom = Some(side),
                                    "diagonal" => border.diagonal = Some(side),
                                    _ => {}
                                }
                            }
                        }
                    }

                    "color" if current_border_side.is_some() && current_border.is_some() => {
                        let color = parse_color_attrs(e);
                        if let Some(ref mut border) = current_border {
                            if let Some(ref side_name) = current_border_side {
                                let side = match side_name.as_str() {
                                    "right" => &mut border.right,
                                    "top" => &mut border.top,
                                    "bottom" => &mut border.bottom,
                                    "diagonal" => &mut border.diagonal,
                                    // "left" and any unknown default to left
                                    _ => &mut border.left,
                                };
                                if let Some(ref mut s) = side {
                                    s.color = Some(color);
                                }
                            }
                        }
                    }

                    "xf" if in_cell_xfs || in_cell_style_xfs => {
                        let mut xf = CellXf::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"fontId" => {
                                    xf.font_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"fillId" => {
                                    xf.fill_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"borderId" => {
                                    xf.border_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"numFmtId" => {
                                    xf.num_fmt_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"xfId" => {
                                    xf.xf_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"applyFont" => {
                                    xf.apply_font =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"applyFill" => {
                                    xf.apply_fill =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"applyBorder" => {
                                    xf.apply_border =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"applyAlignment" => {
                                    xf.apply_alignment =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"applyNumberFormat" => {
                                    xf.apply_number_format =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"applyProtection" => {
                                    xf.apply_protection =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                _ => {}
                            }
                        }

                        // For self-closing <xf .../> tags, add immediately since there's no End event
                        if is_empty {
                            if in_cell_xfs {
                                stylesheet.cell_xfs.push(xf);
                            } else if in_cell_style_xfs {
                                stylesheet.cell_style_xfs.push(xf);
                            }
                        } else {
                            // This was a Start tag, wait for End event
                            current_xf = Some(xf);
                        }
                    }

                    "cellStyle" if in_cell_styles => {
                        let mut name = String::new();
                        let mut xf_id: u32 = 0;
                        let mut builtin_id: Option<u32> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name =
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                }
                                b"xfId" => {
                                    xf_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"builtinId" => {
                                    builtin_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }

                        stylesheet.named_styles.push(NamedStyle {
                            name,
                            xf_id,
                            builtin_id,
                        });
                    }

                    "alignment" if current_xf.is_some() => {
                        let mut align = RawAlignment::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    align.horizontal = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                                b"vertical" => {
                                    align.vertical = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(std::string::ToString::to_string);
                                }
                                b"wrapText" => {
                                    align.wrap_text =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"shrinkToFit" => {
                                    align.shrink_to_fit =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"indent" => {
                                    align.indent = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"textRotation" => {
                                    align.text_rotation = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"readingOrder" => {
                                    align.reading_order = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }

                        if let Some(ref mut xf) = current_xf {
                            xf.alignment = Some(align);
                        }
                    }

                    "protection" if current_xf.is_some() => {
                        // Parse protection element inside xf
                        // Default: locked=true, hidden=false
                        let mut protection = RawProtection {
                            locked: true, // Excel default
                            hidden: false,
                        };

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"locked" => {
                                    // locked="0" means unlocked, locked="1" or absent means locked
                                    protection.locked =
                                        std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                                }
                                b"hidden" => {
                                    protection.hidden =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                _ => {}
                            }
                        }

                        if let Some(ref mut xf) = current_xf {
                            xf.protection = Some(protection);
                        }
                    }

                    // DXF font parsing (for conditional formatting styles)
                    "font" if in_dxf => {
                        // Font element within dxf - we'll parse child elements
                    }
                    "b" if in_dxf => {
                        // Bold within dxf font
                        if let Some(ref mut dxf) = current_dxf {
                            // Check for val="0" which means not bold
                            let mut is_bold = true;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    is_bold =
                                        std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                                }
                            }
                            dxf.bold = Some(is_bold);
                        }
                    }
                    "i" if in_dxf => {
                        // Italic within dxf font
                        if let Some(ref mut dxf) = current_dxf {
                            let mut is_italic = true;
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    is_italic =
                                        std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                                }
                            }
                            dxf.italic = Some(is_italic);
                        }
                    }
                    "u" if in_dxf => {
                        // Underline within dxf font
                        if let Some(ref mut dxf) = current_dxf {
                            dxf.underline = Some(true);
                        }
                    }
                    "strike" if in_dxf => {
                        // Strikethrough within dxf font
                        if let Some(ref mut dxf) = current_dxf {
                            dxf.strikethrough = Some(true);
                        }
                    }
                    "color" if in_dxf => {
                        // Font color within dxf
                        if let Some(ref mut dxf) = current_dxf {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"rgb" => {
                                        if let Ok(rgb) = std::str::from_utf8(&attr.value) {
                                            let color = if rgb.len() == 8 {
                                                format!("#{}", &rgb[2..])
                                            } else {
                                                format!("#{rgb}")
                                            };
                                            dxf.font_color = Some(color);
                                        }
                                    }
                                    b"theme" => {
                                        // Could resolve theme color here if needed
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    // DXF fill parsing
                    "fill" if in_dxf => {
                        // Fill element within dxf
                    }
                    "patternFill" if in_dxf => {
                        // Pattern fill within dxf
                    }
                    "bgColor" if in_dxf => {
                        // Background color within dxf fill
                        if let Some(ref mut dxf) = current_dxf {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    if let Ok(rgb) = std::str::from_utf8(&attr.value) {
                                        let color = if rgb.len() == 8 {
                                            format!("#{}", &rgb[2..])
                                        } else {
                                            format!("#{rgb}")
                                        };
                                        dxf.fill_color = Some(color);
                                    }
                                }
                            }
                        }
                    }
                    "fgColor" if in_dxf => {
                        // Foreground color within dxf fill (often used for solid fills)
                        if let Some(ref mut dxf) = current_dxf {
                            // Only set fill_color if not already set by bgColor
                            if dxf.fill_color.is_none() {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"rgb" {
                                        if let Ok(rgb) = std::str::from_utf8(&attr.value) {
                                            let color = if rgb.len() == 8 {
                                                format!("#{}", &rgb[2..])
                                            } else {
                                                format!("#{rgb}")
                                            };
                                            dxf.fill_color = Some(color);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    _ => {}
                }
            }

            Ok(Event::End(ref e)) => {
                let name = e.local_name();
                let name_str = std::str::from_utf8(name.as_ref()).unwrap_or("");

                match name_str {
                    "numFmts" => in_num_fmts = false,
                    "fonts" => in_fonts = false,
                    "fills" => in_fills = false,
                    "borders" => in_borders = false,
                    "cellXfs" => in_cell_xfs = false,
                    "cellStyleXfs" => in_cell_style_xfs = false,
                    "dxfs" => in_dxfs = false,
                    "dxf" => {
                        if let Some(dxf) = current_dxf.take() {
                            stylesheet.dxf_styles.push(dxf);
                        }
                        in_dxf = false;
                    }
                    "cellStyles" => in_cell_styles = false,
                    "colors" => in_colors = false,
                    "indexedColors" => in_indexed_colors = false,

                    "font" => {
                        if let Some(font) = current_font.take() {
                            stylesheet.fonts.push(font);
                        }
                    }

                    "fill" => {
                        // If we have a gradient, attach it to the fill before pushing
                        if let Some(gradient) = current_gradient.take() {
                            if let Some(ref mut fill) = current_fill {
                                fill.gradient = Some(gradient);
                            }
                        }
                        if let Some(fill) = current_fill.take() {
                            stylesheet.fills.push(fill);
                        }
                    }

                    "gradientFill" => {
                        // Gradient fill has ended; attach it to the current fill
                        if let Some(gradient) = current_gradient.take() {
                            if let Some(ref mut fill) = current_fill {
                                fill.gradient = Some(gradient);
                            }
                        }
                    }

                    "stop" => {
                        in_gradient_stop = false;
                    }

                    "border" => {
                        if let Some(border) = current_border.take() {
                            stylesheet.borders.push(border);
                        }
                    }

                    "xf" if in_cell_xfs => {
                        if let Some(xf) = current_xf.take() {
                            stylesheet.cell_xfs.push(xf);
                        }
                    }

                    "xf" if in_cell_style_xfs => {
                        if let Some(xf) = current_xf.take() {
                            stylesheet.cell_style_xfs.push(xf);
                        }
                    }

                    "left" | "right" | "top" | "bottom" | "diagonal" => {
                        current_border_side = None;
                    }

                    _ => {}
                }
            }

            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => {}
        }

        buf.clear();
    }

    // Store custom indexed colors if any were parsed
    if !indexed_colors.is_empty() {
        stylesheet.indexed_colors = Some(indexed_colors);
    }

    // Extract the default font from the "Normal" style (cellStyleXfs[0])
    // This is the workbook's default font that should apply to all cells
    if let Some(normal_xf) = stylesheet.cell_style_xfs.first() {
        if let Some(font_id) = normal_xf.font_id {
            if let Some(font) = stylesheet.fonts.get(font_id as usize) {
                stylesheet.default_font = Some(font.clone());
            }
        }
    }

    Ok(stylesheet)
}

// Include fill tests module
// #[cfg(test)]
// #[path = "fill_tests.rs"]
// mod fill_tests;

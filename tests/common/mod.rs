//! Common test utilities and assertion helpers.
//!
//! This module provides helper functions for testing the xlview parser,
//! including style extraction, comparison, and assertion utilities.
#![allow(
    dead_code,
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

// Re-export fixtures for convenience
pub use super::fixtures::*;

// ============================================================================
// Workbook Parsing Helper
// ============================================================================

/// Parse XLSX bytes into a workbook JSON value.
///
/// This is a test helper that panics on parse failure.
#[must_use]
pub fn parse_xlsx_to_json(data: &[u8]) -> serde_json::Value {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).expect("Failed to open ZIP archive");

    // We need to re-implement parsing here since the main library
    // returns JsValue which isn't available in tests.
    // Instead, we parse to JSON directly.

    // For now, we'll use a simplified approach that reads the raw XML
    // and extracts key information for testing.

    let mut result = serde_json::json!({
        "sheets": [],
        "theme": {
            "colors": []
        }
    });

    // Parse workbook.xml to get sheet names
    let sheet_names = parse_workbook_sheets(&mut archive);

    // Parse styles.xml
    let styles_info = parse_styles_info(&mut archive);

    // Parse shared strings
    let shared_strings = parse_shared_strings_list(&mut archive);

    // Parse theme colors
    let theme_colors = parse_theme_colors(&mut archive);
    result["theme"]["colors"] = serde_json::json!(theme_colors);

    // Parse each sheet
    let mut sheets = Vec::new();
    for (i, name) in sheet_names.iter().enumerate() {
        let sheet_path = format!("xl/worksheets/sheet{}.xml", i + 1);
        if let Some(sheet_data) = parse_sheet_data(
            &mut archive,
            &sheet_path,
            &shared_strings,
            &styles_info,
            &theme_colors,
        ) {
            let mut sheet_json = sheet_data;
            sheet_json["name"] = serde_json::json!(name);
            sheets.push(sheet_json);
        }
    }

    result["sheets"] = serde_json::json!(sheets);
    result
}

/// Parse workbook.xml to get sheet names.
fn parse_workbook_sheets<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Vec<String> {
    let mut names = Vec::new();

    let Ok(file) = archive.by_name("xl/workbook.xml") else {
        return names;
    };

    let reader = std::io::BufReader::new(file);
    let mut xml = quick_xml::Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e) | quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"sheet" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"name" {
                            if let Ok(name) = std::str::from_utf8(&attr.value) {
                                names.push(name.to_string());
                            }
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    names
}

/// Parsed styles information.
#[derive(Debug, Default)]
pub struct StylesInfo {
    pub fonts: Vec<FontInfo>,
    pub fills: Vec<FillInfo>,
    pub borders: Vec<BorderInfo>,
    pub cell_xfs: Vec<CellXfInfo>,
    pub num_fmts: Vec<(u32, String)>,
}

#[derive(Debug, Default, Clone)]
pub struct FontInfo {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
}

#[derive(Debug, Default, Clone)]
pub struct FillInfo {
    pub pattern_type: Option<String>,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[derive(Debug, Default, Clone)]
pub struct BorderInfo {
    pub left: Option<BorderSideInfo>,
    pub right: Option<BorderSideInfo>,
    pub top: Option<BorderSideInfo>,
    pub bottom: Option<BorderSideInfo>,
}

#[derive(Debug, Clone)]
pub struct BorderSideInfo {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Default, Clone)]
pub struct CellXfInfo {
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
    pub num_fmt_id: Option<u32>,
    pub alignment: Option<AlignmentInfo>,
}

#[derive(Debug, Default, Clone)]
pub struct AlignmentInfo {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub indent: Option<u32>,
    pub rotation: Option<i32>,
}

/// Parse styles.xml into structured info.
fn parse_styles_info<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> StylesInfo {
    let mut info = StylesInfo::default();

    let Ok(file) = archive.by_name("xl/styles.xml") else {
        return info;
    };

    let reader = std::io::BufReader::new(file);
    let mut xml = quick_xml::Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_cell_xfs = false;
    let mut in_num_fmts = false;

    let mut current_font: Option<FontInfo> = None;
    let mut current_fill: Option<FillInfo> = None;
    let mut current_border: Option<BorderInfo> = None;
    let mut current_xf: Option<CellXfInfo> = None;
    let mut current_border_side: Option<String> = None;

    loop {
        let event = xml.read_event_into(&mut buf);
        let is_empty = matches!(&event, Ok(quick_xml::events::Event::Empty(_)));

        match event {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "numFmts" => in_num_fmts = true,
                    "fonts" => in_fonts = true,
                    "fills" => in_fills = true,
                    "borders" => in_borders = true,
                    "cellXfs" => in_cell_xfs = true,

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
                        info.num_fmts.push((id, code));
                    }

                    "font" if in_fonts => {
                        current_font = Some(FontInfo::default());
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
                                        .map(|s| s.to_string());
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
                            font.underline = true;
                        }
                    }

                    "strike" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.strikethrough = true;
                        }
                    }

                    "color" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.color = parse_color_from_attrs(e);
                        }
                    }

                    "fill" if in_fills => {
                        current_fill = Some(FillInfo::default());
                    }

                    "patternFill" if current_fill.is_some() => {
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"patternType" {
                                    fill.pattern_type = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s.to_string());
                                }
                            }
                        }
                    }

                    "fgColor" if current_fill.is_some() => {
                        if let Some(ref mut fill) = current_fill {
                            fill.fg_color = parse_color_from_attrs(e);
                        }
                    }

                    "bgColor" if current_fill.is_some() => {
                        if let Some(ref mut fill) = current_fill {
                            fill.bg_color = parse_color_from_attrs(e);
                        }
                    }

                    "border" if in_borders => {
                        current_border = Some(BorderInfo::default());
                    }

                    "left" | "right" | "top" | "bottom" if current_border.is_some() => {
                        current_border_side = Some(name.to_string());

                        let mut style = String::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                style = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                        }

                        if !style.is_empty() {
                            let side = BorderSideInfo { style, color: None };
                            if let Some(ref mut border) = current_border {
                                match name {
                                    "left" => border.left = Some(side),
                                    "right" => border.right = Some(side),
                                    "top" => border.top = Some(side),
                                    "bottom" => border.bottom = Some(side),
                                    _ => {}
                                }
                            }
                        }
                    }

                    "color" if current_border_side.is_some() && current_border.is_some() => {
                        let color = parse_color_from_attrs(e);
                        if let Some(ref mut border) = current_border {
                            if let Some(ref side_name) = current_border_side {
                                let side = match side_name.as_str() {
                                    "left" => &mut border.left,
                                    "right" => &mut border.right,
                                    "top" => &mut border.top,
                                    "bottom" => &mut border.bottom,
                                    _ => &mut border.left,
                                };
                                if let Some(ref mut s) = side {
                                    s.color = color;
                                }
                            }
                        }
                    }

                    "xf" if in_cell_xfs => {
                        let mut xf = CellXfInfo::default();

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
                                _ => {}
                            }
                        }

                        current_xf = Some(xf);

                        // For Empty elements (self-closing), immediately push
                        if is_empty {
                            if let Some(xf) = current_xf.take() {
                                info.cell_xfs.push(xf);
                            }
                        }
                    }

                    "alignment" if current_xf.is_some() => {
                        let mut align = AlignmentInfo::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    align.horizontal = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s.to_string());
                                }
                                b"vertical" => {
                                    align.vertical = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s.to_string());
                                }
                                b"wrapText" => {
                                    align.wrap_text =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                b"indent" => {
                                    align.indent = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"textRotation" => {
                                    align.rotation = std::str::from_utf8(&attr.value)
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

                    _ => {}
                }
            }

            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "numFmts" => in_num_fmts = false,
                    "fonts" => in_fonts = false,
                    "fills" => in_fills = false,
                    "borders" => in_borders = false,
                    "cellXfs" => in_cell_xfs = false,

                    "font" => {
                        if let Some(font) = current_font.take() {
                            info.fonts.push(font);
                        }
                    }

                    "fill" => {
                        if let Some(fill) = current_fill.take() {
                            info.fills.push(fill);
                        }
                    }

                    "border" => {
                        if let Some(border) = current_border.take() {
                            info.borders.push(border);
                        }
                    }

                    "xf" if in_cell_xfs => {
                        if let Some(xf) = current_xf.take() {
                            info.cell_xfs.push(xf);
                        }
                    }

                    "left" | "right" | "top" | "bottom" => {
                        current_border_side = None;
                    }

                    _ => {}
                }
            }

            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    info
}

/// Parse color from XML attributes.
fn parse_color_from_attrs(e: &quick_xml::events::BytesStart) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == b"rgb" {
            return std::str::from_utf8(&attr.value).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Parse shared strings from xl/sharedStrings.xml.
fn parse_shared_strings_list<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Vec<String> {
    let mut strings = Vec::new();

    let Ok(file) = archive.by_name("xl/sharedStrings.xml") else {
        return strings;
    };

    let reader = std::io::BufReader::new(file);
    let mut xml = quick_xml::Reader::from_reader(reader);
    xml.trim_text(false);

    let mut buf = Vec::new();
    let mut current_string = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                match name {
                    "si" => {
                        in_si = true;
                        current_string.clear();
                    }
                    "t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) if in_t => {
                if let Ok(text) = e.unescape() {
                    current_string.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                match name {
                    "si" => {
                        strings.push(current_string.clone());
                        in_si = false;
                    }
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    strings
}

/// Parse theme colors from xl/theme/theme1.xml.
fn parse_theme_colors<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Vec<String> {
    let default_colors = vec![
        "#000000".to_string(),
        "#FFFFFF".to_string(),
        "#44546A".to_string(),
        "#E7E6E6".to_string(),
        "#4472C4".to_string(),
        "#ED7D31".to_string(),
        "#A5A5A5".to_string(),
        "#FFC000".to_string(),
        "#5B9BD5".to_string(),
        "#70AD47".to_string(),
        "#0563C1".to_string(),
        "#954F72".to_string(),
    ];

    let Ok(file) = archive.by_name("xl/theme/theme1.xml") else {
        return default_colors;
    };

    let reader = std::io::BufReader::new(file);
    let mut xml = quick_xml::Reader::from_reader(reader);
    xml.trim_text(true);

    let mut colors = default_colors;
    let mut buf = Vec::new();
    let mut in_clr_scheme = false;
    let mut color_index = 0usize;

    let color_elements = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e) | quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "clrScheme" {
                    in_clr_scheme = true;
                }

                if in_clr_scheme && color_elements.contains(&name) {
                    color_index = color_elements.iter().position(|&n| n == name).unwrap_or(0);
                }

                if in_clr_scheme && (name == "srgbClr" || name == "sysClr") {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"val" || attr.key.as_ref() == b"lastClr" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                if val.len() == 6 && color_index < colors.len() {
                                    colors[color_index] = format!("#{}", val);
                                }
                            }
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");
                if name == "clrScheme" {
                    in_clr_scheme = false;
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    colors
}

/// Parse sheet data into JSON.
fn parse_sheet_data<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    shared_strings: &[String],
    styles: &StylesInfo,
    theme_colors: &[String],
) -> Option<serde_json::Value> {
    let file = archive.by_name(path).ok()?;

    let reader = std::io::BufReader::new(file);
    let mut xml = quick_xml::Reader::from_reader(reader);
    xml.trim_text(false);

    let mut cells: Vec<serde_json::Value> = Vec::new();
    let mut merges: Vec<serde_json::Value> = Vec::new();
    let mut buf = Vec::new();

    loop {
        let event = xml.read_event_into(&mut buf);
        let is_empty_element = matches!(&event, Ok(quick_xml::events::Event::Empty(_)));

        match event {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "c" {
                    let mut cell_ref = String::new();
                    let mut cell_type = String::new();
                    let mut style_idx: Option<u32> = None;

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"r" => {
                                cell_ref =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"t" => {
                                cell_type =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"s" => {
                                style_idx = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                            }
                            _ => {}
                        }
                    }

                    let (col, row) = parse_cell_ref_internal(&cell_ref);

                    // Read cell value - only for non-empty (Start) elements
                    let mut value: Option<String> = None;
                    let mut cell_buf = Vec::new();

                    // Only read children if this was a Start event (not Empty/self-closing)
                    if !is_empty_element {
                        loop {
                            cell_buf.clear();
                            match xml.read_event_into(&mut cell_buf) {
                                Ok(quick_xml::events::Event::Start(ref inner)) => {
                                    let inner_local = inner.local_name();
                                    let inner_name =
                                        std::str::from_utf8(inner_local.as_ref()).unwrap_or("");
                                    if inner_name == "v" || inner_name == "t" {
                                        let mut text_buf = Vec::new();
                                        if let Ok(quick_xml::events::Event::Text(text)) =
                                            xml.read_event_into(&mut text_buf)
                                        {
                                            value = text.unescape().ok().map(|s| s.to_string());
                                        }
                                    }
                                }
                                Ok(quick_xml::events::Event::End(ref inner)) => {
                                    let inner_local = inner.local_name();
                                    if inner_local.as_ref() == b"c" {
                                        break;
                                    }
                                }
                                Ok(quick_xml::events::Event::Eof) | Err(_) => break,
                                _ => {}
                            }
                        }
                    }

                    // Resolve value
                    let display_value = match cell_type.as_str() {
                        "s" => {
                            let idx: usize =
                                value.as_ref().and_then(|v| v.parse().ok()).unwrap_or(0);
                            shared_strings.get(idx).cloned()
                        }
                        "b" => {
                            // Convert boolean 1/0 to TRUE/FALSE
                            match value.as_deref() {
                                Some("1") | Some("true") => Some("TRUE".to_string()),
                                Some("0") | Some("false") => Some("FALSE".to_string()),
                                _ => value,
                            }
                        }
                        _ => value,
                    };

                    // Resolve style
                    let style =
                        style_idx.and_then(|idx| resolve_style_to_json(idx, styles, theme_colors));

                    let mut cell_json = serde_json::json!({
                        "r": row,
                        "c": col,
                        "cell": {
                            "t": match cell_type.as_str() {
                                "s" | "str" | "inlineStr" => "s",
                                "b" => "b",
                                "e" => "e",
                                _ => "n",
                            }
                        }
                    });

                    if let Some(v) = display_value {
                        cell_json["cell"]["v"] = serde_json::json!(v);
                    }

                    if let Some(s) = style {
                        cell_json["cell"]["s"] = s;
                    }

                    cells.push(cell_json);
                }

                if name == "mergeCell" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"ref" {
                            if let Ok(ref_str) = std::str::from_utf8(&attr.value) {
                                if let Some(merge) = parse_merge_ref_internal(ref_str) {
                                    merges.push(merge);
                                }
                            }
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Some(serde_json::json!({
        "cells": cells,
        "merges": merges,
    }))
}

/// Parse cell reference.
fn parse_cell_ref_internal(ref_str: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut in_letters = true;

    for c in ref_str.chars() {
        if in_letters && c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            in_letters = false;
            if c.is_ascii_digit() {
                row = row * 10 + (c as u32 - '0' as u32);
            }
        }
    }

    (col.saturating_sub(1), row.saturating_sub(1))
}

/// Parse merge reference.
fn parse_merge_ref_internal(ref_str: &str) -> Option<serde_json::Value> {
    let parts: Vec<&str> = ref_str.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_col, start_row) = parse_cell_ref_internal(parts[0]);
    let (end_col, end_row) = parse_cell_ref_internal(parts[1]);

    Some(serde_json::json!({
        "startRow": start_row,
        "startCol": start_col,
        "endRow": end_row,
        "endCol": end_col,
    }))
}

/// Resolve style index to JSON.
fn resolve_style_to_json(
    idx: u32,
    styles: &StylesInfo,
    theme_colors: &[String],
) -> Option<serde_json::Value> {
    let xf = styles.cell_xfs.get(idx as usize)?;
    let mut style = serde_json::Map::new();

    // Font
    if let Some(font_id) = xf.font_id {
        if let Some(font) = styles.fonts.get(font_id as usize) {
            if let Some(ref name) = font.name {
                style.insert("fontFamily".to_string(), serde_json::json!(name));
            }
            if let Some(size) = font.size {
                style.insert("fontSize".to_string(), serde_json::json!(size));
            }
            if font.bold {
                style.insert("bold".to_string(), serde_json::json!(true));
            }
            if font.italic {
                style.insert("italic".to_string(), serde_json::json!(true));
            }
            if font.underline {
                style.insert("underline".to_string(), serde_json::json!(true));
            }
            if font.strikethrough {
                style.insert("strikethrough".to_string(), serde_json::json!(true));
            }
            if let Some(ref color) = font.color {
                let resolved = resolve_color_value(color, theme_colors);
                style.insert("fontColor".to_string(), serde_json::json!(resolved));
            }
        }
    }

    // Fill
    if let Some(fill_id) = xf.fill_id {
        if let Some(fill) = styles.fills.get(fill_id as usize) {
            if fill.pattern_type.as_deref() == Some("solid") {
                if let Some(ref color) = fill.fg_color {
                    let resolved = resolve_color_value(color, theme_colors);
                    style.insert("bgColor".to_string(), serde_json::json!(resolved));
                }
            }
        }
    }

    // Border
    if let Some(border_id) = xf.border_id {
        if let Some(border) = styles.borders.get(border_id as usize) {
            if let Some(ref side) = border.top {
                style.insert(
                    "borderTop".to_string(),
                    border_side_to_json(side, theme_colors),
                );
            }
            if let Some(ref side) = border.right {
                style.insert(
                    "borderRight".to_string(),
                    border_side_to_json(side, theme_colors),
                );
            }
            if let Some(ref side) = border.bottom {
                style.insert(
                    "borderBottom".to_string(),
                    border_side_to_json(side, theme_colors),
                );
            }
            if let Some(ref side) = border.left {
                style.insert(
                    "borderLeft".to_string(),
                    border_side_to_json(side, theme_colors),
                );
            }
        }
    }

    // Alignment
    if let Some(ref align) = xf.alignment {
        if let Some(ref h) = align.horizontal {
            style.insert("alignH".to_string(), serde_json::json!(h));
        }
        if let Some(ref v) = align.vertical {
            let v_mapped = match v.as_str() {
                "center" => "middle",
                other => other,
            };
            style.insert("alignV".to_string(), serde_json::json!(v_mapped));
        }
        if align.wrap_text {
            style.insert("wrap".to_string(), serde_json::json!(true));
        }
        if let Some(indent) = align.indent {
            style.insert("indent".to_string(), serde_json::json!(indent));
        }
        if let Some(rotation) = align.rotation {
            style.insert("rotation".to_string(), serde_json::json!(rotation));
        }
    }

    if style.is_empty() {
        None
    } else {
        Some(serde_json::Value::Object(style))
    }
}

/// Convert border side to JSON.
fn border_side_to_json(side: &BorderSideInfo, theme_colors: &[String]) -> serde_json::Value {
    let color = side
        .color
        .as_ref()
        .map(|c| resolve_color_value(c, theme_colors))
        .unwrap_or_else(|| "#000000".to_string());

    serde_json::json!({
        "style": side.style,
        "color": color,
    })
}

/// Resolve color value (ARGB) to #RRGGBB.
fn resolve_color_value(color: &str, _theme_colors: &[String]) -> String {
    let color = color.trim_start_matches('#');
    if color.len() == 8 {
        format!("#{}", &color[2..])
    } else {
        format!("#{}", color)
    }
}

// ============================================================================
// Assertion Helpers
// ============================================================================

/// Assert that a cell exists at the given position with the expected value.
pub fn assert_cell_value(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
    expected: &str,
) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let value = cell["cell"]["v"]
        .as_str()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no value", row, col));

    assert_eq!(
        value, expected,
        "Cell value mismatch at row={}, col={}",
        row, col
    );
}

/// Assert that a cell has bold formatting.
pub fn assert_cell_bold(workbook: &serde_json::Value, sheet: usize, row: u32, col: u32) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let bold = cell["cell"]["s"]["bold"].as_bool().unwrap_or(false);
    assert!(bold, "Cell at row={}, col={} is not bold", row, col);
}

/// Assert that a cell has italic formatting.
pub fn assert_cell_italic(workbook: &serde_json::Value, sheet: usize, row: u32, col: u32) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let italic = cell["cell"]["s"]["italic"].as_bool().unwrap_or(false);
    assert!(italic, "Cell at row={}, col={} is not italic", row, col);
}

/// Assert that a cell has specific font size.
pub fn assert_cell_font_size(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
    expected: f64,
) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let size = cell["cell"]["s"]["fontSize"]
        .as_f64()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no font size", row, col));

    assert!(
        (size - expected).abs() < 0.01,
        "Cell font size mismatch at row={}, col={}: expected {}, got {}",
        row,
        col,
        expected,
        size
    );
}

/// Assert that a cell has specific background color.
pub fn assert_cell_bg_color(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
    expected: &str,
) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let color = cell["cell"]["s"]["bgColor"]
        .as_str()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no background color", row, col));

    let expected_normalized = normalize_color_for_compare(expected);
    let actual_normalized = normalize_color_for_compare(color);

    assert_eq!(
        actual_normalized, expected_normalized,
        "Cell background color mismatch at row={}, col={}: expected {}, got {}",
        row, col, expected, color
    );
}

/// Assert that a cell has specific font color.
pub fn assert_cell_font_color(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
    expected: &str,
) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let color = cell["cell"]["s"]["fontColor"]
        .as_str()
        .unwrap_or_else(|| panic!("Cell at row={}, col={} has no font color", row, col));

    let expected_normalized = normalize_color_for_compare(expected);
    let actual_normalized = normalize_color_for_compare(color);

    assert_eq!(
        actual_normalized, expected_normalized,
        "Cell font color mismatch at row={}, col={}: expected {}, got {}",
        row, col, expected, color
    );
}

/// Assert that a cell has horizontal alignment.
pub fn assert_cell_align_h(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
    expected: &str,
) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let align = cell["cell"]["s"]["alignH"].as_str().unwrap_or_else(|| {
        panic!(
            "Cell at row={}, col={} has no horizontal alignment",
            row, col
        )
    });

    assert_eq!(
        align, expected,
        "Cell horizontal alignment mismatch at row={}, col={}",
        row, col
    );
}

/// Assert that a cell has text wrapping enabled.
pub fn assert_cell_wrap(workbook: &serde_json::Value, sheet: usize, row: u32, col: u32) {
    let cell = get_cell(workbook, sheet, row, col)
        .unwrap_or_else(|| panic!("Cell at row={}, col={} not found", row, col));

    let wrap = cell["cell"]["s"]["wrap"].as_bool().unwrap_or(false);
    assert!(
        wrap,
        "Cell at row={}, col={} does not have text wrap",
        row, col
    );
}

/// Assert that a merge range exists.
pub fn assert_merge_exists(
    workbook: &serde_json::Value,
    sheet: usize,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
) {
    let merges = &workbook["sheets"][sheet]["merges"];
    let found = merges
        .as_array()
        .map(|arr| {
            arr.iter().any(|m| {
                m["startRow"].as_u64() == Some(u64::from(start_row))
                    && m["startCol"].as_u64() == Some(u64::from(start_col))
                    && m["endRow"].as_u64() == Some(u64::from(end_row))
                    && m["endCol"].as_u64() == Some(u64::from(end_col))
            })
        })
        .unwrap_or(false);

    assert!(
        found,
        "Merge range {}:{} to {}:{} not found in sheet {}",
        start_row, start_col, end_row, end_col, sheet
    );
}

/// Assert that a specific number of sheets exist.
pub fn assert_sheet_count(workbook: &serde_json::Value, expected: usize) {
    let count = workbook["sheets"].as_array().map(|a| a.len()).unwrap_or(0);
    assert_eq!(
        count, expected,
        "Sheet count mismatch: expected {}, got {}",
        expected, count
    );
}

/// Assert that a sheet has the expected name.
pub fn assert_sheet_name(workbook: &serde_json::Value, sheet: usize, expected: &str) {
    let name = workbook["sheets"][sheet]["name"]
        .as_str()
        .unwrap_or_else(|| panic!("Sheet {} has no name", sheet));

    assert_eq!(name, expected, "Sheet name mismatch at index {}", sheet);
}

// ============================================================================
// Helper Functions
// ============================================================================

/// Get a cell from the parsed workbook.
pub fn get_cell(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
) -> Option<serde_json::Value> {
    let cells = workbook["sheets"][sheet]["cells"].as_array()?;

    cells
        .iter()
        .find(|c| {
            c["r"].as_u64() == Some(u64::from(row)) && c["c"].as_u64() == Some(u64::from(col))
        })
        .cloned()
}

/// Get style from a cell.
pub fn get_cell_style(
    workbook: &serde_json::Value,
    sheet: usize,
    row: u32,
    col: u32,
) -> Option<serde_json::Value> {
    let cell = get_cell(workbook, sheet, row, col)?;
    cell["cell"]["s"]
        .as_object()
        .map(|_| cell["cell"]["s"].clone())
}

/// Normalize color for comparison.
fn normalize_color_for_compare(color: &str) -> String {
    color.trim_start_matches('#').to_uppercase()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_minimal_xlsx() {
        let xlsx = minimal_xlsx();
        let workbook = parse_xlsx_to_json(&xlsx);

        assert_sheet_count(&workbook, 1);
        assert_sheet_name(&workbook, 0, "Sheet1");
    }

    #[test]
    fn test_parse_xlsx_with_text() {
        let xlsx = xlsx_with_text("Hello, World!");
        let workbook = parse_xlsx_to_json(&xlsx);

        assert_cell_value(&workbook, 0, 0, 0, "Hello, World!");
    }

    #[test]
    fn test_parse_xlsx_with_styled_cell() {
        let style = StyleBuilder::new()
            .bold()
            .font_size(14.0)
            .bg_color("#FFFF00")
            .build();

        let xlsx = xlsx_with_styled_cell("Styled", style);
        let workbook = parse_xlsx_to_json(&xlsx);

        assert_cell_value(&workbook, 0, 0, 0, "Styled");
        assert_cell_bold(&workbook, 0, 0, 0);
        assert_cell_font_size(&workbook, 0, 0, 0, 14.0);
        assert_cell_bg_color(&workbook, 0, 0, 0, "#FFFF00");
    }
}

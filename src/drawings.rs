//! Drawings and images parsing module
//!
//! This module handles parsing of embedded images and charts from XLSX files.
//! It parses drawing*.xml files to get image positions and references,
//! resolves relationships to find image paths, and reads image data from xl/media/.
//!
//! # XLSX Drawing Structure
//!
//! Drawings are stored in `xl/drawings/drawing*.xml` files. Each sheet can have
//! at most one drawing file, referenced via `xl/worksheets/_rels/sheet*.xml.rels`.
//!
//! The drawing XML contains anchor elements that define positioning:
//! - `twoCellAnchor`: Anchored to two cells (resizes with cells)
//! - `oneCellAnchor`: Anchored to one cell with absolute size
//! - `absoluteAnchor`: Absolute position (not anchored to cells)
//!
//! Each anchor contains a picture (`pic`), chart (`graphicFrame`), or shape (`sp`).
//! Images reference their data via relationship IDs (`r:embed`), which are resolved
//! through `xl/drawings/_rels/drawing*.xml.rels` to paths in `xl/media/`.

use base64::{engine::general_purpose::STANDARD as BASE64, Engine};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use zip::ZipArchive;

use crate::types::{Drawing, EmbeddedImage, ImageFormat};

/// Parse drawings from drawing XML and return Drawing objects
///
/// This is the main entry point for parsing a drawing file. It reads the drawing XML,
/// parses all anchor elements, and returns a vector of Drawing objects ready to be
/// stored in the Sheet struct.
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `drawing_path` - Path to the drawing, e.g., "xl/drawings/drawing1.xml"
///
/// # Returns
/// A tuple of (drawings, image_rels) where:
/// - drawings: Vec of Drawing objects with position info
/// - image_rels: HashMap mapping rId -> image path for loading image data
pub fn parse_drawing_file<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    drawing_path: &str,
) -> (Vec<Drawing>, HashMap<String, String>) {
    let mut drawings = Vec::new();

    // Normalize path - remove leading slash if present
    let normalized_path = drawing_path.trim_start_matches('/');

    // First, get the image relationships
    let image_rels = get_image_relationships(archive, normalized_path);

    let Ok(file) = archive.by_name(normalized_path) else {
        return (drawings, image_rels);
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    // Current parsing state
    let mut current_drawing: Option<DrawingBuilder> = None;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_ext = false; // extent element
    let mut in_pic = false;
    let mut in_chart = false;
    let mut in_sp = false; // shape
    let mut in_blip_fill = false;
    let mut in_nv_pic_pr = false; // non-visual picture properties
    let mut in_c_nv_pr = false; // common non-visual properties
    let mut current_element: Option<String> = None;
    // Shape-specific state
    let mut in_sp_pr = false; // shape properties
    let mut in_ln = false; // line element
    let mut in_tx_body = false; // text body
    let mut in_xfrm = false; // inside transform element (for precise positioning)
    let mut shape_text_parts: Vec<String> = Vec::new();
    // Hyperlink state - stores pending rId to resolve later
    let mut pending_hyperlink_rid: Option<String> = None;
    let mut pending_hyperlink_tooltip: Option<String> = None;
    // Group shape state
    let mut in_grp_sp = false; // Inside a group shape
    let mut in_nv_grp_sp_pr = false; // Inside group non-visual properties
                                     // Shape non-visual properties state (for text box detection)
    let mut in_nv_sp_pr = false; // Inside shape non-visual properties

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "twoCellAnchor" => {
                        let mut builder = DrawingBuilder::new("twoCellAnchor");
                        // Parse editAs attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"editAs" {
                                builder.edit_as = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                        current_drawing = Some(builder);
                    }
                    "oneCellAnchor" => {
                        current_drawing = Some(DrawingBuilder::new("oneCellAnchor"));
                    }
                    "absoluteAnchor" => {
                        current_drawing = Some(DrawingBuilder::new("absoluteAnchor"));
                    }
                    "from" => in_from = true,
                    "to" => in_to = true,
                    "ext" => {
                        in_ext = true;
                        // Parse extent dimensions
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"cx" => {
                                        drawing.extent_cx = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"cy" => {
                                        drawing.extent_cy = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "pos" => {
                        // Absolute position (for absoluteAnchor)
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"x" => {
                                        drawing.pos_x = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"y" => {
                                        drawing.pos_y = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "pic" => {
                        in_pic = true;
                        if let Some(ref mut drawing) = current_drawing {
                            drawing.drawing_type = Some("picture".to_string());
                        }
                    }
                    "graphicFrame" => {
                        in_chart = true;
                        if let Some(ref mut drawing) = current_drawing {
                            drawing.drawing_type = Some("chart".to_string());
                        }
                    }
                    "sp" => {
                        in_sp = true;
                        // Only set drawing_type if not inside a group (child shapes don't create new drawings)
                        if !in_grp_sp {
                            if let Some(ref mut drawing) = current_drawing {
                                drawing.drawing_type = Some("shape".to_string());
                            }
                        }
                    }
                    // Connector shape
                    "cxnSp" => {
                        in_sp = true; // Connectors are treated like shapes
                        if let Some(ref mut drawing) = current_drawing {
                            drawing.drawing_type = Some("connector".to_string());
                        }
                    }
                    // Group shape
                    "grpSp" => {
                        in_grp_sp = true;
                        if let Some(ref mut drawing) = current_drawing {
                            drawing.drawing_type = Some("group".to_string());
                        }
                    }
                    // Group non-visual properties
                    "nvGrpSpPr" => {
                        in_nv_grp_sp_pr = true;
                    }
                    "nvPicPr" => {
                        in_nv_pic_pr = true;
                    }
                    // Shape non-visual properties (contains cNvSpPr which has txBox attribute)
                    "nvSpPr" => {
                        in_nv_sp_pr = true;
                    }
                    // Shape-specific non-visual properties - check for txBox attribute
                    "cNvSpPr" if in_nv_sp_pr => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"txBox" {
                                let is_text_box = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(|s| s == "1" || s == "true")
                                    .unwrap_or(false);
                                if is_text_box {
                                    if let Some(ref mut drawing) = current_drawing {
                                        drawing.drawing_type = Some("textbox".to_string());
                                    }
                                }
                            }
                        }
                    }
                    "cNvPr" => {
                        in_c_nv_pr = true;
                        // Parse name and description from attributes
                        // Only update for top-level elements or group properties, not child shapes in groups
                        let should_update = !in_grp_sp || in_nv_grp_sp_pr;
                        if should_update {
                            if let Some(ref mut drawing) = current_drawing {
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"name" => {
                                            drawing.name = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        b"descr" => {
                                            drawing.description = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        b"title" => {
                                            drawing.title = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }
                    }
                    // Hyperlink in drawing
                    "hlinkClick" if in_c_nv_pr => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id"
                                || key == b"id"
                                || (key.len() > 3 && key.ends_with(b":id"))
                            {
                                pending_hyperlink_rid = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            } else if key == b"tooltip" {
                                pending_hyperlink_tooltip = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                    }
                    "blipFill" => in_blip_fill = true,
                    "blip" if in_blip_fill || in_pic => {
                        // Parse r:embed attribute for image relationship ID
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                let key = attr.key.as_ref();
                                // Check for r:embed or embed (namespace variations)
                                if key == b"r:embed"
                                    || key == b"embed"
                                    || (key.len() > 6 && key.ends_with(b":embed"))
                                {
                                    drawing.image_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "chart" if in_chart => {
                        // Parse r:id attribute for chart relationship ID
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                let key = attr.key.as_ref();
                                if key == b"r:id"
                                    || key == b"id"
                                    || (key.len() > 3 && key.ends_with(b":id"))
                                {
                                    drawing.chart_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "xfrm" => {
                        // Transform element - parse rotation and flip attributes
                        in_xfrm = true;
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"rot" => {
                                        drawing.rotation = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"flipH" => {
                                        drawing.flip_h = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(|s| s == "1" || s == "true");
                                    }
                                    b"flipV" => {
                                        drawing.flip_v = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(|s| s == "1" || s == "true");
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "col" | "colOff" | "row" | "rowOff" | "t" => {
                        current_element = Some(name.to_string());
                    }
                    // Shape properties
                    "spPr" if in_sp => {
                        in_sp_pr = true;
                    }
                    "ln" if in_sp_pr => {
                        in_ln = true;
                    }
                    "txBody" if in_sp => {
                        in_tx_body = true;
                        shape_text_parts.clear();
                    }
                    // Preset geometry (shape type) - can be start element with children
                    "prstGeom" if in_sp_pr => {
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"prst" {
                                    drawing.shape_type = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "ext" => {
                        // Parse extent dimensions (self-closing element)
                        // When inside xfrm, this is the precise transform extent
                        // Otherwise, it's the anchor-level extent
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"cx" => {
                                        let val = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                        if in_xfrm {
                                            drawing.xfrm_cx = val;
                                        } else {
                                            drawing.extent_cx = val;
                                        }
                                    }
                                    b"cy" => {
                                        let val = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                        if in_xfrm {
                                            drawing.xfrm_cy = val;
                                        } else {
                                            drawing.extent_cy = val;
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    // Transform offset - precise position from Excel (inside xfrm)
                    "off" if in_xfrm => {
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"x" => {
                                        drawing.xfrm_x = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"y" => {
                                        drawing.xfrm_y = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "pos" => {
                        // Absolute position (self-closing)
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"x" => {
                                        drawing.pos_x = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"y" => {
                                        drawing.pos_y = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "cNvPr" => {
                        // Parse name and description from attributes (self-closing)
                        // Only update for top-level elements or group properties, not child shapes in groups
                        let should_update = !in_grp_sp || in_nv_grp_sp_pr;
                        if should_update {
                            if let Some(ref mut drawing) = current_drawing {
                                for attr in e.attributes().flatten() {
                                    match attr.key.as_ref() {
                                        b"name" => {
                                            drawing.name = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        b"descr" => {
                                            drawing.description = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        b"title" => {
                                            drawing.title = std::str::from_utf8(&attr.value)
                                                .ok()
                                                .map(ToString::to_string);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }
                    }
                    // Shape-specific non-visual properties (self-closing) - check for txBox attribute
                    "cNvSpPr" => {
                        if in_nv_sp_pr {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"txBox" {
                                    let is_text_box = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(|s| s == "1" || s == "true")
                                        .unwrap_or(false);
                                    if is_text_box {
                                        if let Some(ref mut drawing) = current_drawing {
                                            drawing.drawing_type = Some("textbox".to_string());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Hyperlink in drawing (self-closing)
                    "hlinkClick" => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id"
                                || key == b"id"
                                || (key.len() > 3 && key.ends_with(b":id"))
                            {
                                pending_hyperlink_rid = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            } else if key == b"tooltip" {
                                pending_hyperlink_tooltip = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                    }
                    "blip" => {
                        // Parse r:embed attribute for image relationship ID (self-closing)
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                let key = attr.key.as_ref();
                                if key == b"r:embed"
                                    || key == b"embed"
                                    || (key.len() > 6 && key.ends_with(b":embed"))
                                {
                                    drawing.image_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "chart" => {
                        // Parse r:id attribute for chart relationship ID (self-closing)
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                let key = attr.key.as_ref();
                                if key == b"r:id"
                                    || key == b"id"
                                    || (key.len() > 3 && key.ends_with(b":id"))
                                {
                                    drawing.chart_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "xfrm" => {
                        // Transform element (self-closing)
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"rot" => {
                                        drawing.rotation = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                    }
                                    b"flipH" => {
                                        drawing.flip_h = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(|s| s == "1" || s == "true");
                                    }
                                    b"flipV" => {
                                        drawing.flip_v = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(|s| s == "1" || s == "true");
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    // Preset geometry (shape type) - self-closing
                    "prstGeom" if in_sp_pr => {
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"prst" {
                                    drawing.shape_type = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    // Solid fill color - self-closing srgbClr
                    "srgbClr" if in_sp_pr && !in_ln => {
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    let hex = std::str::from_utf8(&attr.value).unwrap_or("");
                                    drawing.fill_color = Some(format!("#{}", hex));
                                }
                            }
                        }
                    }
                    // Line fill color - self-closing srgbClr inside ln
                    "srgbClr" if in_ln => {
                        if let Some(ref mut drawing) = current_drawing {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    let hex = std::str::from_utf8(&attr.value).unwrap_or("");
                                    drawing.line_color = Some(format!("#{}", hex));
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let (Some(ref element), Some(ref mut drawing)) =
                    (&current_element, &mut current_drawing)
                {
                    if let Ok(text) = e.unescape() {
                        match element.as_str() {
                            "col" if in_from => {
                                drawing.from_col = text.parse().ok();
                            }
                            "row" if in_from => {
                                drawing.from_row = text.parse().ok();
                            }
                            "colOff" if in_from => {
                                drawing.from_col_off = text.parse().ok();
                            }
                            "rowOff" if in_from => {
                                drawing.from_row_off = text.parse().ok();
                            }
                            "col" if in_to => {
                                drawing.to_col = text.parse().ok();
                            }
                            "row" if in_to => {
                                drawing.to_row = text.parse().ok();
                            }
                            "colOff" if in_to => {
                                drawing.to_col_off = text.parse().ok();
                            }
                            "rowOff" if in_to => {
                                drawing.to_row_off = text.parse().ok();
                            }
                            // Text content in shapes
                            "t" if in_tx_body => {
                                shape_text_parts.push(text.to_string());
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" => {
                        if let Some(builder) = current_drawing.take() {
                            if let Some(drawing) = builder.build() {
                                drawings.push(drawing);
                            }
                        }
                    }
                    "from" => in_from = false,
                    "to" => in_to = false,
                    "ext" => in_ext = false,
                    "pic" => in_pic = false,
                    "graphicFrame" => in_chart = false,
                    "sp" | "cxnSp" => in_sp = false,
                    "grpSp" => in_grp_sp = false,
                    "nvGrpSpPr" => in_nv_grp_sp_pr = false,
                    "nvPicPr" => in_nv_pic_pr = false,
                    "nvSpPr" => in_nv_sp_pr = false,
                    "cNvPr" => in_c_nv_pr = false,
                    "blipFill" => in_blip_fill = false,
                    "col" | "colOff" | "row" | "rowOff" | "t" => current_element = None,
                    // Shape-specific closing tags
                    "spPr" => in_sp_pr = false,
                    "ln" => in_ln = false,
                    "xfrm" => in_xfrm = false,
                    "txBody" => {
                        // Collect text content from shape
                        if in_tx_body && !shape_text_parts.is_empty() {
                            if let Some(ref mut drawing) = current_drawing {
                                drawing.text_content = Some(shape_text_parts.join(" "));
                            }
                        }
                        in_tx_body = false;
                        shape_text_parts.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Suppress unused variable warnings
    let _ = in_ext;
    let _ = in_sp;
    let _ = in_nv_pic_pr;
    let _ = in_nv_sp_pr;
    let _ = in_c_nv_pr;
    let _ = in_sp_pr;
    let _ = in_ln;
    let _ = in_tx_body;
    let _ = in_grp_sp;
    let _ = in_nv_grp_sp_pr;
    let _ = pending_hyperlink_rid;
    let _ = pending_hyperlink_tooltip;

    (drawings, image_rels)
}

/// Get drawing file path from sheet relationships
///
/// Looks in xl/worksheets/_rels/sheetN.xml.rels for a relationship
/// with type containing "drawing" and returns the target path.
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `sheet_path` - Path to the sheet, e.g., "xl/worksheets/sheet1.xml"
///
/// # Returns
/// The full path to the drawing file, e.g., "xl/drawings/drawing1.xml"
pub fn get_drawing_path<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_path: &str,
) -> Option<String> {
    // Convert sheet path to rels path
    // e.g., "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
    let sheet_path = sheet_path.trim_start_matches('/');
    let rels_path = construct_rels_path(sheet_path);

    let Ok(file) = archive.by_name(&rels_path) else {
        return None;
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut target = String::new();
                    let mut rel_type = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Target" => {
                                target = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Type" => {
                                rel_type =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            _ => {}
                        }
                    }

                    // Check if this is a drawing relationship
                    if rel_type.contains("drawing") && !target.is_empty() {
                        // Resolve relative path
                        let base_dir = if let Some(pos) = sheet_path.rfind('/') {
                            &sheet_path[..pos]
                        } else {
                            ""
                        };

                        let full_path = resolve_relative_path(base_dir, &target);
                        return Some(full_path);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

/// Parse image relationships from drawing rels
///
/// Parses xl/drawings/_rels/drawingN.xml.rels to get the mapping
/// from relationship IDs (rId) to image file paths.
fn get_image_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let mut rels = HashMap::new();

    let rels_path = construct_rels_path(drawing_path);

    let Ok(file) = archive.by_name(&rels_path) else {
        return rels;
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = String::new();
                    let mut target = String::new();
                    let mut rel_type = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => {
                                id = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Target" => {
                                target = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Type" => {
                                rel_type =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            _ => {}
                        }
                    }

                    // Include image relationships (type contains "image")
                    if !id.is_empty() && !target.is_empty() && rel_type.contains("image") {
                        // Resolve relative path
                        let base_dir = if let Some(pos) = drawing_path.rfind('/') {
                            &drawing_path[..pos]
                        } else {
                            ""
                        };

                        let full_path = resolve_relative_path(base_dir, &target);
                        rels.insert(id, full_path);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    rels
}

/// Read image data from the archive and return as EmbeddedImage
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `image_path` - Path to the image, e.g., "xl/media/image1.png"
///
/// # Returns
/// An EmbeddedImage with base64-encoded data, or None if the image cannot be read
pub fn read_image<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    image_path: &str,
) -> Option<EmbeddedImage> {
    let normalized_path = image_path.trim_start_matches('/');

    let mut file = archive.by_name(normalized_path).ok()?;

    let mut data = Vec::new();
    file.read_to_end(&mut data).ok()?;

    if data.is_empty() {
        return None;
    }

    // Detect image format from magic bytes first, then fall back to extension
    let format = ImageFormat::from_magic_bytes(&data);
    let format = if format == ImageFormat::Unknown {
        // Try extension
        let ext = normalized_path.rsplit('.').next().unwrap_or("");
        ImageFormat::from_extension(ext)
    } else {
        format
    };

    let mime_type = format.mime_type().to_string();
    let base64_data = BASE64.encode(&data);

    // Extract filename from path
    let filename = normalized_path.rsplit('/').next().map(ToString::to_string);

    Some(EmbeddedImage {
        id: normalized_path.to_string(),
        mime_type,
        data: base64_data,
        filename,
        width: None,  // Could parse from image header if needed
        height: None, // Could parse from image header if needed
    })
}

/// Collect all unique image paths from drawings
pub fn collect_image_paths(
    drawings: &[Drawing],
    image_rels: &HashMap<String, String>,
) -> Vec<String> {
    let mut paths: Vec<String> = drawings
        .iter()
        .filter_map(|d| d.image_id.as_ref())
        .filter_map(|id| image_rels.get(id))
        .cloned()
        .collect();

    // Remove duplicates while preserving order
    paths.sort();
    paths.dedup();
    paths
}

/// Read all images referenced by drawings
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `image_paths` - List of image paths to read
///
/// # Returns
/// A vector of EmbeddedImage objects
pub fn read_all_images<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    image_paths: &[String],
) -> Vec<EmbeddedImage> {
    image_paths
        .iter()
        .filter_map(|path| read_image(archive, path))
        .collect()
}

/// Construct the relationships file path from a file path
/// e.g., "xl/drawings/drawing1.xml" -> "xl/drawings/_rels/drawing1.xml.rels"
fn construct_rels_path(file_path: &str) -> String {
    if let Some(pos) = file_path.rfind('/') {
        let dir = &file_path[..pos];
        let filename = &file_path[pos + 1..];
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{file_path}.rels")
    }
}

/// Resolve a relative path against a base directory
///
/// Handles paths like "../media/image1.png" relative to "xl/drawings"
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    // If path is absolute (starts with /), just remove the leading slash
    if let Some(stripped) = relative.strip_prefix('/') {
        return stripped.to_string();
    }

    // Split base directory into components
    let mut components: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

    // Process relative path
    for part in relative.split('/') {
        match part {
            ".." => {
                components.pop();
            }
            "." | "" => {}
            _ => components.push(part),
        }
    }

    components.join("/")
}

/// Builder for constructing Drawing objects during parsing
#[derive(Debug, Default)]
struct DrawingBuilder {
    anchor_type: String,
    drawing_type: Option<String>,
    name: Option<String>,
    description: Option<String>,
    title: Option<String>,
    from_col: Option<u32>,
    from_row: Option<u32>,
    from_col_off: Option<i64>,
    from_row_off: Option<i64>,
    to_col: Option<u32>,
    to_row: Option<u32>,
    to_col_off: Option<i64>,
    to_row_off: Option<i64>,
    pos_x: Option<i64>,
    pos_y: Option<i64>,
    extent_cx: Option<i64>,
    extent_cy: Option<i64>,
    edit_as: Option<String>,
    image_id: Option<String>,
    chart_id: Option<String>,
    shape_type: Option<String>,
    fill_color: Option<String>,
    line_color: Option<String>,
    text_content: Option<String>,
    rotation: Option<i64>,
    flip_h: Option<bool>,
    flip_v: Option<bool>,
    xfrm_x: Option<i64>,
    xfrm_y: Option<i64>,
    xfrm_cx: Option<i64>,
    xfrm_cy: Option<i64>,
}

impl DrawingBuilder {
    fn new(anchor_type: &str) -> Self {
        Self {
            anchor_type: anchor_type.to_string(),
            ..Default::default()
        }
    }

    fn build(self) -> Option<Drawing> {
        // Need at least anchor_type and drawing_type
        let drawing_type = self.drawing_type?;

        Some(Drawing {
            anchor_type: self.anchor_type,
            drawing_type,
            name: self.name,
            description: self.description,
            title: self.title,
            from_col: self.from_col,
            from_row: self.from_row,
            from_col_off: self.from_col_off,
            from_row_off: self.from_row_off,
            to_col: self.to_col,
            to_row: self.to_row,
            to_col_off: self.to_col_off,
            to_row_off: self.to_row_off,
            pos_x: self.pos_x,
            pos_y: self.pos_y,
            extent_cx: self.extent_cx,
            extent_cy: self.extent_cy,
            edit_as: self.edit_as,
            image_id: self.image_id,
            chart_id: self.chart_id,
            shape_type: self.shape_type,
            fill_color: self.fill_color,
            line_color: self.line_color,
            text_content: self.text_content,
            rotation: self.rotation,
            flip_h: self.flip_h,
            flip_v: self.flip_v,
            hyperlink: None,
            xfrm_x: self.xfrm_x,
            xfrm_y: self.xfrm_y,
            xfrm_cx: self.xfrm_cx,
            xfrm_cy: self.xfrm_cy,
        })
    }
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
    fn test_resolve_relative_path() {
        // Basic relative path
        assert_eq!(
            resolve_relative_path("xl/drawings", "image1.png"),
            "xl/drawings/image1.png"
        );

        // Parent directory
        assert_eq!(
            resolve_relative_path("xl/drawings", "../media/image1.png"),
            "xl/media/image1.png"
        );

        // Multiple parent directories
        assert_eq!(
            resolve_relative_path("xl/drawings/sub", "../../media/image1.png"),
            "xl/media/image1.png"
        );

        // Absolute path
        assert_eq!(
            resolve_relative_path("xl/drawings", "/xl/media/image1.png"),
            "xl/media/image1.png"
        );

        // Current directory
        assert_eq!(
            resolve_relative_path("xl/drawings", "./image1.png"),
            "xl/drawings/image1.png"
        );
    }

    #[test]
    fn test_construct_rels_path() {
        assert_eq!(
            construct_rels_path("xl/drawings/drawing1.xml"),
            "xl/drawings/_rels/drawing1.xml.rels"
        );

        assert_eq!(
            construct_rels_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );

        assert_eq!(
            construct_rels_path("workbook.xml"),
            "_rels/workbook.xml.rels"
        );
    }

    #[test]
    fn test_drawing_builder() {
        // Complete drawing
        let builder = DrawingBuilder {
            anchor_type: "twoCellAnchor".to_string(),
            drawing_type: Some("picture".to_string()),
            name: Some("Picture 1".to_string()),
            description: Some("A test image".to_string()),
            title: None,
            from_col: Some(0),
            from_row: Some(0),
            from_col_off: Some(0),
            from_row_off: Some(0),
            to_col: Some(5),
            to_row: Some(10),
            to_col_off: Some(304800),
            to_row_off: Some(190500),
            pos_x: None,
            pos_y: None,
            extent_cx: None,
            extent_cy: None,
            edit_as: Some("oneCell".to_string()),
            image_id: Some("rId1".to_string()),
            chart_id: None,
            shape_type: None,
            fill_color: None,
            line_color: None,
            text_content: None,
            rotation: None,
            flip_h: None,
            flip_v: None,
            xfrm_x: None,
            xfrm_y: None,
            xfrm_cx: None,
            xfrm_cy: None,
        };

        let drawing = builder.build().unwrap();
        assert_eq!(drawing.anchor_type, "twoCellAnchor");
        assert_eq!(drawing.drawing_type, "picture");
        assert_eq!(drawing.name, Some("Picture 1".to_string()));
        assert_eq!(drawing.from_col, Some(0));
        assert_eq!(drawing.to_col, Some(5));
        assert_eq!(drawing.to_col_off, Some(304800));
        assert_eq!(drawing.image_id, Some("rId1".to_string()));

        // Missing drawing type
        let incomplete = DrawingBuilder {
            anchor_type: "twoCellAnchor".to_string(),
            drawing_type: None,
            ..Default::default()
        };
        assert!(incomplete.build().is_none());
    }

    #[test]
    fn test_image_format_detection() {
        // PNG magic bytes
        let png_data = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(ImageFormat::from_magic_bytes(&png_data), ImageFormat::Png);

        // JPEG magic bytes
        let jpeg_data = [0xFF, 0xD8, 0xFF, 0xE0];
        assert_eq!(ImageFormat::from_magic_bytes(&jpeg_data), ImageFormat::Jpeg);

        // Extension detection
        assert_eq!(ImageFormat::from_extension("png"), ImageFormat::Png);
        assert_eq!(ImageFormat::from_extension("PNG"), ImageFormat::Png);
        assert_eq!(ImageFormat::from_extension("jpg"), ImageFormat::Jpeg);
        assert_eq!(ImageFormat::from_extension("jpeg"), ImageFormat::Jpeg);
        assert_eq!(ImageFormat::from_extension("gif"), ImageFormat::Gif);
        assert_eq!(ImageFormat::from_extension("unknown"), ImageFormat::Unknown);
    }

    #[test]
    fn test_mime_types() {
        assert_eq!(ImageFormat::Png.mime_type(), "image/png");
        assert_eq!(ImageFormat::Jpeg.mime_type(), "image/jpeg");
        assert_eq!(ImageFormat::Gif.mime_type(), "image/gif");
        assert_eq!(ImageFormat::Unknown.mime_type(), "application/octet-stream");
    }
}

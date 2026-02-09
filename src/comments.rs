//! Comments/Notes parsing module
//! This module handles parsing of cell comments from XLSX files.
//!
//! Excel comments (called "notes" in newer versions) are stored in separate XML files
//! within the XLSX package. Each sheet can have its own comments file (e.g., xl/comments1.xml)
//! linked via the sheet's relationship file (xl/worksheets/_rels/sheet1.xml.rels).

use crate::color::resolve_color;
use crate::types::{Comment, RichTextRun, RunStyle, VerticalAlign};
use crate::xml_helpers::parse_color_attrs;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{BufReader, Read, Seek};
use zip::ZipArchive;

/// Parse comments from comments XML file
/// Returns a vector of Comment structs
///
/// # XML Structure
/// ```xml
/// <comments>
///   <authors>
///     <author>John Doe</author>
///   </authors>
///   <commentList>
///     <comment ref="A1" authorId="0">
///       <text>
///         <r><t>Comment text</t></r>
///       </text>
///     </comment>
///   </commentList>
/// </comments>
/// ```
pub fn parse_comments<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    comments_path: &str,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Vec<Comment> {
    let Ok(file) = archive.by_name(comments_path) else {
        return Vec::new();
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(false);

    let mut comments = Vec::new();
    let mut authors: Vec<String> = Vec::new();
    let mut buf = Vec::new();

    // Parsing state
    let mut in_authors = false;
    let mut in_author = false;
    let mut in_comment_list = false;
    let mut in_comment = false;
    let mut in_text = false;
    let mut in_r = false;
    let mut in_t = false;
    let mut in_rpr = false;

    // Current comment being parsed
    let mut current_author = String::new();
    let mut current_cell_ref = String::new();
    let mut current_author_id: Option<u32> = None;
    let mut current_text_parts: Vec<String> = Vec::new();
    let mut current_rich_runs: Vec<RichTextRun> = Vec::new();
    let mut current_run_text = String::new();
    let mut current_run_style: Option<RunStyle> = None;
    let mut has_any_styling = false;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "authors" => {
                        in_authors = true;
                    }
                    "author" if in_authors => {
                        in_author = true;
                        current_author.clear();
                    }
                    "commentList" => {
                        in_comment_list = true;
                    }
                    "comment" if in_comment_list => {
                        in_comment = true;
                        current_cell_ref.clear();
                        current_author_id = None;
                        current_text_parts.clear();
                        current_rich_runs.clear();
                        has_any_styling = false;

                        // Parse attributes
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"ref" => {
                                    current_cell_ref =
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                }
                                b"authorId" => {
                                    current_author_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                _ => {}
                            }
                        }
                    }
                    "text" if in_comment => {
                        in_text = true;
                    }
                    "r" if in_text => {
                        in_r = true;
                        current_run_text.clear();
                        current_run_style = None;
                    }
                    "rPr" if in_r => {
                        in_rpr = true;
                        current_run_style = Some(RunStyle::default());
                    }
                    "t" if in_text => {
                        in_t = true;
                    }
                    // Font styling elements within rPr
                    "b" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            style.bold = Some(true);
                            has_any_styling = true;
                        }
                    }
                    "i" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            style.italic = Some(true);
                            has_any_styling = true;
                        }
                    }
                    "u" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            style.underline = Some(true);
                            has_any_styling = true;
                        }
                    }
                    "strike" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            style.strikethrough = Some(true);
                            has_any_styling = true;
                        }
                    }
                    "sz" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    style.font_size = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                    if style.font_size.is_some() {
                                        has_any_styling = true;
                                    }
                                }
                            }
                        }
                    }
                    "rFont" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    style.font_family = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                    if style.font_family.is_some() {
                                        has_any_styling = true;
                                    }
                                }
                            }
                        }
                    }
                    "color" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            let color_spec = parse_color_attrs(e);
                            style.font_color =
                                resolve_color(&color_spec, theme_colors, indexed_colors);
                            if style.font_color.is_some() {
                                has_any_styling = true;
                            }
                        }
                    }
                    "vertAlign" if in_rpr => {
                        if let Some(ref mut style) = current_run_style {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or("");
                                    style.vert_align = match val {
                                        "superscript" => Some(VerticalAlign::Superscript),
                                        "subscript" => Some(VerticalAlign::Subscript),
                                        _ => Some(VerticalAlign::Baseline),
                                    };
                                    if style.vert_align.is_some() {
                                        has_any_styling = true;
                                    }
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

                // Handle self-closing tags for font styling
                if in_rpr {
                    match name {
                        "b" => {
                            if let Some(ref mut style) = current_run_style {
                                // Check for val="0" which means not bold
                                let mut is_bold = true;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        is_bold =
                                            std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                                    }
                                }
                                if is_bold {
                                    style.bold = Some(true);
                                    has_any_styling = true;
                                }
                            }
                        }
                        "i" => {
                            if let Some(ref mut style) = current_run_style {
                                let mut is_italic = true;
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        is_italic =
                                            std::str::from_utf8(&attr.value).unwrap_or("1") != "0";
                                    }
                                }
                                if is_italic {
                                    style.italic = Some(true);
                                    has_any_styling = true;
                                }
                            }
                        }
                        "u" => {
                            if let Some(ref mut style) = current_run_style {
                                style.underline = Some(true);
                                has_any_styling = true;
                            }
                        }
                        "strike" => {
                            if let Some(ref mut style) = current_run_style {
                                style.strikethrough = Some(true);
                                has_any_styling = true;
                            }
                        }
                        "sz" => {
                            if let Some(ref mut style) = current_run_style {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        style.font_size = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .and_then(|s| s.parse().ok());
                                        if style.font_size.is_some() {
                                            has_any_styling = true;
                                        }
                                    }
                                }
                            }
                        }
                        "rFont" => {
                            if let Some(ref mut style) = current_run_style {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        style.font_family = std::str::from_utf8(&attr.value)
                                            .ok()
                                            .map(ToString::to_string);
                                        if style.font_family.is_some() {
                                            has_any_styling = true;
                                        }
                                    }
                                }
                            }
                        }
                        "color" => {
                            if let Some(ref mut style) = current_run_style {
                                let color_spec = parse_color_attrs(e);
                                style.font_color =
                                    resolve_color(&color_spec, theme_colors, indexed_colors);
                                if style.font_color.is_some() {
                                    has_any_styling = true;
                                }
                            }
                        }
                        "vertAlign" => {
                            if let Some(ref mut style) = current_run_style {
                                for attr in e.attributes().flatten() {
                                    if attr.key.as_ref() == b"val" {
                                        let val = std::str::from_utf8(&attr.value).unwrap_or("");
                                        style.vert_align = match val {
                                            "superscript" => Some(VerticalAlign::Superscript),
                                            "subscript" => Some(VerticalAlign::Subscript),
                                            _ => Some(VerticalAlign::Baseline),
                                        };
                                        if style.vert_align.is_some() {
                                            has_any_styling = true;
                                        }
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_author {
                    if let Ok(text) = e.unescape() {
                        current_author.push_str(&text);
                    }
                } else if in_t {
                    if let Ok(text) = e.unescape() {
                        if in_r {
                            current_run_text.push_str(&text);
                        } else {
                            // Plain text directly in <text> element (not in a run)
                            current_text_parts.push(text.to_string());
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "authors" => {
                        in_authors = false;
                    }
                    "author" => {
                        in_author = false;
                        authors.push(current_author.clone());
                        current_author.clear();
                    }
                    "commentList" => {
                        in_comment_list = false;
                    }
                    "comment" => {
                        in_comment = false;

                        // Get author name from author ID
                        let author = current_author_id
                            .and_then(|id| authors.get(id as usize))
                            .cloned();

                        // Build the plain text from either rich runs or plain text parts
                        let text = if !current_rich_runs.is_empty() {
                            current_rich_runs
                                .iter()
                                .map(|r| r.text.as_str())
                                .collect::<String>()
                        } else {
                            current_text_parts.join("")
                        };

                        // Only include rich_text if there was actual styling
                        let rich_text = if has_any_styling && !current_rich_runs.is_empty() {
                            Some(current_rich_runs.clone())
                        } else {
                            None
                        };

                        comments.push(Comment {
                            cell_ref: current_cell_ref.clone(),
                            author,
                            text,
                            rich_text,
                        });
                    }
                    "text" => {
                        in_text = false;
                    }
                    "r" => {
                        in_r = false;
                        // Save the current run
                        if !current_run_text.is_empty() {
                            current_rich_runs.push(RichTextRun {
                                text: current_run_text.clone(),
                                style: current_run_style.take(),
                            });
                        }
                        current_run_text.clear();
                    }
                    "rPr" => {
                        in_rpr = false;
                    }
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    comments
}

/// Get comments file path from sheet relationships
///
/// # XML Structure
/// ```xml
/// <Relationships>
///   <Relationship Id="rId1"
///     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
///     Target="../comments1.xml"/>
/// </Relationships>
/// ```
pub fn get_comments_path<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_path: &str,
) -> Option<String> {
    // Construct the relationship file path from the sheet path
    // e.g., "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
    let rels_path = if let Some(pos) = sheet_path.rfind('/') {
        let dir = &sheet_path[..pos + 1];
        let file = &sheet_path[pos + 1..];
        format!("{dir}_rels/{file}.rels")
    } else {
        format!("_rels/{sheet_path}.rels")
    };

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

                    // Check if this is a comments relationship
                    if rel_type.contains("comments") && !target.is_empty() {
                        // Resolve the target path relative to the sheet's directory
                        let sheet_dir = if let Some(pos) = sheet_path.rfind('/') {
                            &sheet_path[..pos + 1]
                        } else {
                            ""
                        };

                        // Handle relative paths (e.g., "../comments1.xml")
                        let full_path = resolve_relative_path(sheet_dir, &target);
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

/// Resolve a relative path from a base directory
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    if let Some(stripped) = relative.strip_prefix('/') {
        // Absolute path from root
        stripped.to_string()
    } else if let Some(stripped) = relative.strip_prefix("../") {
        // Go up one directory
        let parent = if let Some(pos) = base_dir.trim_end_matches('/').rfind('/') {
            &base_dir[..pos + 1]
        } else {
            ""
        };
        resolve_relative_path(parent, stripped)
    } else {
        // Relative to current directory
        format!("{base_dir}{relative}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_relative_path_parent() {
        assert_eq!(
            resolve_relative_path("xl/worksheets/", "../comments1.xml"),
            "xl/comments1.xml"
        );
    }

    #[test]
    fn test_resolve_relative_path_same_dir() {
        assert_eq!(
            resolve_relative_path("xl/worksheets/", "comments1.xml"),
            "xl/worksheets/comments1.xml"
        );
    }

    #[test]
    fn test_resolve_relative_path_absolute() {
        assert_eq!(
            resolve_relative_path("xl/worksheets/", "/xl/comments1.xml"),
            "xl/comments1.xml"
        );
    }

    #[test]
    fn test_resolve_relative_path_double_parent() {
        assert_eq!(
            resolve_relative_path("xl/worksheets/subdir/", "../../comments1.xml"),
            "xl/comments1.xml"
        );
    }
}

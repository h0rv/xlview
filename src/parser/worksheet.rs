//! Worksheet parsing - parses individual sheet XML into Sheet structs.

use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{BufReader, Read};
use zip::ZipArchive;

use crate::auto_filter::parse_auto_filter;
use crate::cell_ref::{parse_cell_ref_bytes_or_default, parse_cell_ref_or_default, parse_sqref};
use crate::conditional::parse_conditional_formatting;
use crate::data_validation::parse_data_validation;
use crate::error::Result;
use crate::sparklines::parse_ext_sparklines;
use crate::types::{
    Cell, CellData, CellType, ColWidth, MergeRange, PaneState, RowHeight, Sheet, StyleRef,
    StyleSheet, Theme,
};
use crate::xml_helpers::parse_color_attrs;

use super::{f64_to_u32_clamped, now_ms, NumFmtInfo, ParseOptions, SheetParseMetrics};

use crate::color::resolve_color;

use super::styles::resolve_cell_value;

/// Sheet metadata from workbook.xml
pub(super) struct SheetInfo {
    pub name: String,
    pub path: String,
    pub state: crate::types::SheetState,
}

/// Cell type tag from the `t` attribute of a `<c>` element.
#[derive(Copy, Clone)]
pub(super) enum CellTypeTag {
    Shared,
    Inline,
    Str,
    Bool,
    Error,
    Default,
}

pub(super) fn parse_cell_type_tag(value: &[u8]) -> CellTypeTag {
    match value {
        b"s" => CellTypeTag::Shared,
        b"b" => CellTypeTag::Bool,
        b"e" => CellTypeTag::Error,
        b"str" => CellTypeTag::Str,
        b"inlineStr" => CellTypeTag::Inline,
        _ => CellTypeTag::Default,
    }
}

pub(super) fn parse_u32_bytes(value: &[u8]) -> Option<u32> {
    let mut num: u32 = 0;
    let mut seen = false;
    for &b in value {
        if !b.is_ascii_digit() {
            return None;
        }
        seen = true;
        num = num.saturating_mul(10).saturating_add(u32::from(b - b'0'));
    }
    if seen {
        Some(num)
    } else {
        None
    }
}

/// Parse a merge range like "A1:B2"
fn parse_merge_ref(ref_str: &str) -> Option<MergeRange> {
    let (start_part, end_part) = ref_str.split_once(':')?;

    let (start_col, start_row) = parse_cell_ref_or_default(start_part);
    let (end_col, end_row) = parse_cell_ref_or_default(end_part);

    Some(MergeRange {
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

/// Parse a dimension range like "A1:B2" into (start_row, start_col, end_row, end_col).
fn parse_dimension_ref(ref_str: &str) -> Option<(u32, u32, u32, u32)> {
    let (start_part, end_part) = ref_str.split_once(':').unwrap_or((ref_str, ref_str));

    let (start_col, start_row) = parse_cell_ref_or_default(start_part);
    let (end_col, end_row) = parse_cell_ref_or_default(end_part);

    Some((start_row, start_col, end_row, end_col))
}

/// Parse a single worksheet
#[allow(
    clippy::too_many_lines,
    clippy::too_many_arguments,
    clippy::needless_option_as_deref
)]
pub(super) fn parse_sheet<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    info: &SheetInfo,
    shared_strings: &[String],
    stylesheet: &StyleSheet,
    theme: &Theme,
    date1904: bool,
    digit_width: f64,
    options: ParseOptions,
    numfmt_lookup: &[NumFmtInfo],
    resolved_styles: &[Option<StyleRef>],
    default_style: &Option<StyleRef>,
    mut metrics: Option<&mut SheetParseMetrics>,
) -> Result<Sheet> {
    let file = archive.by_name(&info.path)?;

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(false);

    let mut sheet = Sheet {
        name: info.name.clone(),
        state: info.state,
        tab_color: None,
        cells: Vec::new(),
        cells_by_row: Vec::new(),
        merges: Vec::new(),
        col_widths: Vec::new(),
        row_heights: Vec::new(),
        default_col_width: 64.0,  // ~8.43 characters
        default_row_height: 20.0, // ~15 points
        hidden_cols: Vec::new(),
        hidden_rows: Vec::new(),
        max_row: 0,
        max_col: 0,
        frozen_rows: 0,
        frozen_cols: 0,
        split_row: None,
        split_col: None,
        pane_state: None,
        is_protected: false,
        data_validations: Vec::new(),
        auto_filter: None,
        outline_level_row: Vec::new(),
        outline_level_col: Vec::new(),
        outline_summary_below: true,
        outline_summary_right: true,
        hyperlinks: Vec::new(),
        comments: Vec::new(),
        comments_by_cell: HashMap::new(),
        print_area: None,
        row_breaks: Vec::new(),
        col_breaks: Vec::new(),
        print_titles_rows: None,
        print_titles_cols: None,
        page_margins: None,
        page_setup: None,
        header_footer: None,
        sparkline_groups: Vec::new(),
        drawings: Vec::new(),
        charts: Vec::new(),
        conditional_formatting: Vec::new(),
        conditional_formatting_cache: Vec::new(),
    };

    let mut buf = Vec::new();
    let mut cell_buf = Vec::new();
    let mut text_buf = Vec::new();
    let mut current_row: u32 = 0;
    let mut in_sheet_pr = false;
    let mut in_sheet_view = false;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(ref event @ (Event::Start(_) | Event::Empty(_))) => {
                let (Event::Start(ref e) | Event::Empty(ref e)) = event else {
                    continue;
                };
                let is_start_event = matches!(event, Event::Start(_));
                let local_name = e.local_name();

                match local_name.as_ref() {
                    b"dimension" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                if let Ok(ref_str) = std::str::from_utf8(&attr.value) {
                                    if let Some((start_row, start_col, end_row, end_col)) =
                                        parse_dimension_ref(ref_str)
                                    {
                                        let rows = end_row.saturating_sub(start_row) + 1;
                                        let cols = end_col.saturating_sub(start_col) + 1;
                                        let area = rows.saturating_mul(cols);
                                        let reserve = area.min(200_000) as usize;
                                        if reserve > 0 {
                                            sheet.cells.reserve(reserve);
                                        }
                                        let dim_max_row = end_row.saturating_add(1);
                                        let dim_max_col = end_col.saturating_add(1);
                                        if dim_max_row > sheet.max_row {
                                            sheet.max_row = dim_max_row;
                                        }
                                        if dim_max_col > sheet.max_col {
                                            sheet.max_col = dim_max_col;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"sheetFormatPr" => {
                        // Parse default column width and row height
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"defaultColWidth" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        if let Ok(w) = s.parse::<f64>() {
                                            // Convert from character width to pixels using the workbook's default font
                                            sheet.default_col_width =
                                                ((w * digit_width + 5.0) / digit_width * 256.0)
                                                    .floor()
                                                    / 256.0
                                                    * digit_width;
                                        }
                                    }
                                }
                                b"defaultRowHeight" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        if let Ok(h) = s.parse::<f64>() {
                                            // Convert from points to pixels
                                            sheet.default_row_height = h * (96.0 / 72.0);
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }

                    b"sheetPr" => {
                        in_sheet_pr = true;
                    }

                    b"tabColor" if in_sheet_pr => {
                        // Parse tab color from sheetPr/tabColor
                        let color_spec = parse_color_attrs(e);
                        sheet.tab_color = resolve_color(
                            &color_spec,
                            &theme.colors,
                            stylesheet.indexed_colors.as_ref(),
                        );
                    }

                    b"sheetView" => {
                        in_sheet_view = true;
                    }

                    b"pane" if in_sheet_view => {
                        // Parse frozen/split pane information
                        // <pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
                        let mut x_split: Option<f64> = None;
                        let mut y_split: Option<f64> = None;
                        let mut state_str = String::new();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"xSplit" => {
                                    x_split = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"ySplit" => {
                                    y_split = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"state" => {
                                    state_str =
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                }
                                _ => {}
                            }
                        }

                        // Determine pane state and set appropriate fields
                        match state_str.as_str() {
                            "frozen" => {
                                sheet.pane_state = Some(PaneState::Frozen);
                                // For frozen panes, xSplit/ySplit are column/row counts
                                if let Some(x) = x_split {
                                    sheet.frozen_cols = f64_to_u32_clamped(x);
                                }
                                if let Some(y) = y_split {
                                    sheet.frozen_rows = f64_to_u32_clamped(y);
                                }
                            }
                            "frozenSplit" => {
                                sheet.pane_state = Some(PaneState::FrozenSplit);
                                // For frozenSplit, xSplit/ySplit are column/row counts
                                if let Some(x) = x_split {
                                    sheet.frozen_cols = f64_to_u32_clamped(x);
                                }
                                if let Some(y) = y_split {
                                    sheet.frozen_rows = f64_to_u32_clamped(y);
                                }
                            }
                            "split" => {
                                sheet.pane_state = Some(PaneState::Split);
                                // For split panes, xSplit/ySplit are positions in twips
                                sheet.split_col = x_split;
                                sheet.split_row = y_split;
                            }
                            _ => {
                                // No state attribute but has split values - treat as split
                                if x_split.is_some() || y_split.is_some() {
                                    sheet.pane_state = Some(PaneState::Split);
                                    sheet.split_col = x_split;
                                    sheet.split_row = y_split;
                                }
                            }
                        }
                    }

                    b"row" => {
                        if let Some(m) = metrics.as_deref_mut() {
                            m.row_count = m.row_count.saturating_add(1);
                        }
                        let mut row_height: Option<f64> = None;
                        let mut row_hidden = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    current_row = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"ht" => {
                                    row_height = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"hidden" => {
                                    row_hidden =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                _ => {}
                            }
                        }

                        if let Some(ht) = row_height {
                            sheet.row_heights.push(RowHeight {
                                row: current_row.saturating_sub(1),
                                height: ht * (96.0 / 72.0), // Convert points to pixels (DPI ratio)
                            });
                        }

                        if row_hidden {
                            sheet.hidden_rows.push(current_row.saturating_sub(1));
                        }

                        if current_row > sheet.max_row {
                            sheet.max_row = current_row;
                        }
                    }

                    b"c" => {
                        // Cell element - parse cell attributes first
                        let mut col: u32 = 0;
                        let mut row: u32 = 0;
                        let mut cell_type = CellTypeTag::Default;
                        let mut style_idx: Option<u32> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    let (c, r) = parse_cell_ref_bytes_or_default(&attr.value);
                                    col = c;
                                    row = r;
                                }
                                b"t" => {
                                    cell_type = parse_cell_type_tag(&attr.value);
                                }
                                b"s" => {
                                    style_idx = parse_u32_bytes(&attr.value);
                                }
                                _ => {}
                            }
                        }

                        // Read cell value from child elements
                        // Only do this for Start events (non-empty cells)
                        // Empty/self-closing cells like <c r="A1"/> have no child elements
                        let mut value: Option<String> = None;
                        if is_start_event {
                            loop {
                                cell_buf.clear();
                                match xml.read_event_into(&mut cell_buf) {
                                    Ok(Event::Start(ref inner)) => {
                                        let inner_name = inner.local_name();
                                        let inner_name = inner_name.as_ref();

                                        if inner_name == b"v" || inner_name == b"t" {
                                            // Value or inline text (direct child of <c>)
                                            text_buf.clear();
                                            if let Ok(Event::Text(text)) =
                                                xml.read_event_into(&mut text_buf)
                                            {
                                                let raw = text.as_ref();
                                                let needs_unescape = raw.contains(&b'&');
                                                if let Some(m) = metrics.as_deref_mut() {
                                                    m.text_unescape_calls =
                                                        m.text_unescape_calls.saturating_add(1);
                                                    let unescape_start = now_ms();
                                                    value = if needs_unescape {
                                                        text.unescape().ok().map(|s| s.to_string())
                                                    } else {
                                                        std::str::from_utf8(raw)
                                                            .ok()
                                                            .map(ToString::to_string)
                                                    };
                                                    m.text_unescape_ms += now_ms() - unescape_start;
                                                } else {
                                                    value = if needs_unescape {
                                                        text.unescape().ok().map(|s| s.to_string())
                                                    } else {
                                                        std::str::from_utf8(raw)
                                                            .ok()
                                                            .map(ToString::to_string)
                                                    };
                                                }
                                            }
                                        } else if inner_name == b"is" {
                                            // Inline string container <is><t>text</t></is>
                                            // Read nested elements to find <t>
                                            loop {
                                                text_buf.clear();
                                                match xml.read_event_into(&mut text_buf) {
                                                    Ok(Event::Start(ref is_inner)) => {
                                                        if is_inner.local_name().as_ref() == b"t" {
                                                            // Found <t> inside <is>
                                                            let mut t_buf = Vec::new();
                                                            if let Ok(Event::Text(text)) =
                                                                xml.read_event_into(&mut t_buf)
                                                            {
                                                                let raw = text.as_ref();
                                                                let needs_unescape =
                                                                    raw.contains(&b'&');
                                                                if let Some(m) =
                                                                    metrics.as_deref_mut()
                                                                {
                                                                    m.text_unescape_calls = m
                                                                        .text_unescape_calls
                                                                        .saturating_add(1);
                                                                }
                                                                value = if needs_unescape {
                                                                    text.unescape()
                                                                        .ok()
                                                                        .map(|s| s.to_string())
                                                                } else {
                                                                    std::str::from_utf8(raw)
                                                                        .ok()
                                                                        .map(ToString::to_string)
                                                                };
                                                            }
                                                        }
                                                    }
                                                    Ok(Event::End(ref is_inner)) => {
                                                        if is_inner.local_name().as_ref() == b"is" {
                                                            break;
                                                        }
                                                    }
                                                    Ok(Event::Eof) | Err(_) => break,
                                                    _ => {}
                                                }
                                            }
                                        }
                                    }
                                    Ok(Event::End(ref inner)) => {
                                        if inner.local_name().as_ref() == b"c" {
                                            break;
                                        }
                                    }
                                    Ok(Event::Eof) | Err(_) => break,
                                    _ => {}
                                }
                            }
                        }

                        // Resolve cell type and value
                        let (final_value, raw_value, cached_display, final_type) =
                            resolve_cell_value(
                                value.as_deref(),
                                cell_type,
                                shared_strings,
                                style_idx,
                                numfmt_lookup,
                                date1904,
                                options,
                                metrics.as_deref_mut(),
                            );

                        // Resolve style - if no style index, use default font style
                        let style = if let Some(idx) = style_idx {
                            if let Some(m) = metrics.as_deref_mut() {
                                m.style_count = m.style_count.saturating_add(1);
                            }
                            if options.eager_values {
                                resolved_styles.get(idx as usize).and_then(|s| s.clone())
                            } else {
                                None
                            }
                        } else {
                            // Cell has no explicit style, apply default font from Normal style
                            if let Some(m) = metrics.as_deref_mut() {
                                m.default_style_count = m.default_style_count.saturating_add(1);
                            }
                            if options.eager_values {
                                default_style.clone()
                            } else {
                                None
                            }
                        };

                        if col + 1 > sheet.max_col {
                            sheet.max_col = col + 1;
                        }

                        if let Some(m) = metrics.as_deref_mut() {
                            m.cell_count = m.cell_count.saturating_add(1);
                            match final_type {
                                CellType::String => {
                                    m.string_cells = m.string_cells.saturating_add(1)
                                }
                                CellType::Number => {
                                    m.number_cells = m.number_cells.saturating_add(1)
                                }
                                CellType::Boolean => m.bool_cells = m.bool_cells.saturating_add(1),
                                CellType::Error => m.error_cells = m.error_cells.saturating_add(1),
                                CellType::Date => m.date_cells = m.date_cells.saturating_add(1),
                            }
                            match cell_type {
                                CellTypeTag::Shared => {
                                    m.shared_string_cells = m.shared_string_cells.saturating_add(1);
                                }
                                CellTypeTag::Inline | CellTypeTag::Str => {
                                    m.inline_string_cells = m.inline_string_cells.saturating_add(1);
                                }
                                _ => {}
                            }
                        }

                        sheet.cells.push(CellData {
                            r: row,
                            c: col,
                            cell: Cell {
                                v: final_value,
                                t: final_type,
                                s: style,
                                style_idx,
                                raw: raw_value,
                                cached_display,
                                rich_text: None,
                                cached_rich_text: None,
                                has_comment: None,
                                hyperlink: None,
                            },
                        });
                    }

                    b"col" => {
                        // Column definition
                        let mut min: u32 = 0;
                        let mut max: u32 = 0;
                        let mut width: f64 = 8.43;
                        let mut hidden = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"max" => {
                                    max = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                                b"width" => {
                                    width = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(8.43);
                                }
                                b"hidden" => {
                                    hidden = std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                                _ => {}
                            }
                        }

                        // Store the Excel character width directly
                        // Viewer will convert to pixels: width * ~7.0 (digit width varies by font)
                        let span = max.saturating_sub(min) + 1;
                        if let Some(m) = metrics.as_deref_mut() {
                            m.col_count = m.col_count.saturating_add(u64::from(span));
                        }

                        for col in min..=max {
                            sheet.col_widths.push(ColWidth {
                                col: col.saturating_sub(1),
                                width,
                            });

                            if hidden {
                                sheet.hidden_cols.push(col.saturating_sub(1));
                            }
                        }
                    }

                    b"mergeCell" => {
                        // Merged cell range
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                if let Ok(ref_str) = std::str::from_utf8(&attr.value) {
                                    if let Some(merge) = parse_merge_ref(ref_str) {
                                        if let Some(m) = metrics.as_deref_mut() {
                                            m.merge_count = m.merge_count.saturating_add(1);
                                        }
                                        sheet.merges.push(merge);
                                    }
                                }
                            }
                        }
                    }

                    b"sheetProtection" => {
                        // Parse sheet protection status
                        // <sheetProtection sheet="1" objects="1" scenarios="1" password="XXXX"/>
                        // The presence of this element with sheet="1" means the sheet is protected
                        let mut found_sheet_attr = false;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"sheet" {
                                found_sheet_attr = true;
                                sheet.is_protected =
                                    std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                break;
                            }
                        }
                        // If element exists but no sheet attribute, consider it protected
                        // (conservative approach - presence of sheetProtection usually means protection)
                        if !found_sheet_attr {
                            sheet.is_protected = true;
                        }
                    }

                    b"autoFilter" => {
                        // Parse auto-filter from sheet XML
                        // <autoFilter ref="A1:D100">
                        //   <filterColumn colId="0">
                        //     <filters><filter val="Value1"/></filters>
                        //   </filterColumn>
                        // </autoFilter>
                        sheet.auto_filter = parse_auto_filter(e, &mut xml);
                        if let Some(filter) = sheet.auto_filter.as_ref() {
                            let end_row = filter.end_row.saturating_add(1);
                            let end_col = filter.end_col.saturating_add(1);
                            if end_row > sheet.max_row {
                                sheet.max_row = end_row;
                            }
                            if end_col > sheet.max_col {
                                sheet.max_col = end_col;
                            }
                        }
                    }

                    b"dataValidation" => {
                        // Parse data validation rule
                        // <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="A1:A100">
                        //   <formula1>"Option1,Option2,Option3"</formula1>
                        // </dataValidation>
                        if let Some(validation) = parse_data_validation(e, &mut xml) {
                            for (_start_row, _start_col, end_row, end_col) in
                                parse_sqref(&validation.sqref)
                            {
                                let end_row = end_row.saturating_add(1);
                                let end_col = end_col.saturating_add(1);
                                if end_row > sheet.max_row {
                                    sheet.max_row = end_row;
                                }
                                if end_col > sheet.max_col {
                                    sheet.max_col = end_col;
                                }
                            }
                            sheet.data_validations.push(validation);
                        }
                    }

                    b"conditionalFormatting" => {
                        // Parse conditional formatting rules
                        // <conditionalFormatting sqref="A1:A10">
                        //   <cfRule type="colorScale" priority="1">
                        //     <colorScale>
                        //       <cfvo type="min"/>
                        //       <cfvo type="max"/>
                        //       <color rgb="FFF8696B"/>
                        //       <color rgb="FF63BE7B"/>
                        //     </colorScale>
                        //   </cfRule>
                        // </conditionalFormatting>
                        if let Some(cf) = parse_conditional_formatting(
                            e,
                            &mut xml,
                            &theme.colors,
                            stylesheet.indexed_colors.as_ref(),
                        ) {
                            for (_start_row, _start_col, end_row, end_col) in parse_sqref(&cf.sqref)
                            {
                                let end_row = end_row.saturating_add(1);
                                let end_col = end_col.saturating_add(1);
                                if end_row > sheet.max_row {
                                    sheet.max_row = end_row;
                                }
                                if end_col > sheet.max_col {
                                    sheet.max_col = end_col;
                                }
                            }
                            sheet.conditional_formatting.push(cf);
                        }
                    }

                    b"ext" => {
                        // Parse extension elements (sparklines, etc.)
                        let groups = parse_ext_sparklines(
                            e,
                            &mut xml,
                            &theme.colors,
                            stylesheet.indexed_colors.as_ref(),
                        );
                        sheet.sparkline_groups.extend(groups);
                    }

                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"sheetPr" => in_sheet_pr = false,
                    b"sheetView" => in_sheet_view = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    sheet.rebuild_cell_index();
    Ok(sheet)
}

//! Main XLSX parser
//!
//! Orchestrates the parsing of all components from the ZIP archive.

mod relationships;
pub(crate) mod styles;
mod worksheet;

use quick_xml::Reader;
use std::io::{BufReader, Cursor};
use zip::ZipArchive;

use serde::Serialize;
use std::collections::HashMap;
use std::sync::Arc;

use crate::cell_ref::parse_cell_ref as parse_cell_ref_str;
use crate::comments::{get_comments_path, parse_comments};
use crate::drawings::{get_drawing_path, parse_drawing_file};
use crate::error::Result;
use crate::hyperlinks::{parse_hyperlink_rels, parse_hyperlinks, resolve_hyperlinks};
use crate::numfmt::{compile_format_code, get_builtin_format, CompiledFormat};
use crate::types::{HyperlinkDef, Sheet, StyleRef, Workbook};

use relationships::{
    collect_and_read_images, get_sheet_info, parse_charts_from_drawing, parse_shared_strings,
    parse_stylesheet, parse_theme, parse_workbook_relationships,
};
use styles::{get_default_style, resolve_style};
use worksheet::parse_sheet;

#[cfg(target_arch = "wasm32")]
fn now_ms() -> f64 {
    if let Some(window) = web_sys::window() {
        if let Some(perf) = window.performance() {
            return perf.now();
        }
    }
    js_sys::Date::now()
}

#[cfg(not(target_arch = "wasm32"))]
fn now_ms() -> f64 {
    use std::time::Instant;
    thread_local! {
        static START: Instant = Instant::now();
    }
    START.with(|s| s.elapsed().as_secs_f64() * 1000.0)
}

/// Detailed timing metrics for XLSX parsing.
#[derive(Clone, Debug, Default, Serialize)]
pub struct ParseMetrics {
    pub parse_ms: f64,
    pub relationships_ms: f64,
    pub theme_ms: f64,
    pub shared_strings_ms: f64,
    pub styles_ms: f64,
    pub workbook_info_ms: f64,
    pub sheets_ms: f64,
    pub charts_resolve_ms: f64,
    pub images_ms: f64,
    pub dxf_ms: f64,
    pub style_resolve_ms: f64,
    pub format_number_ms: f64,
    pub format_number_date_ms: f64,
    pub format_number_number_ms: f64,
    pub value_parse_ms: f64,
    pub text_unescape_ms: f64,
    pub sheets_count: u64,
    pub shared_strings_count: u64,
    pub shared_strings_chars: u64,
    pub styles_fonts: u64,
    pub styles_fills: u64,
    pub styles_borders: u64,
    pub styles_cell_xfs: u64,
    pub styles_cell_style_xfs: u64,
    pub styles_num_fmts: u64,
    pub styles_named_styles: u64,
    pub styles_dxf: u64,
    pub styles_indexed_colors: u64,
    pub total_cells: u64,
    pub total_rows: u64,
    pub total_cols: u64,
    pub total_merges: u64,
    pub total_styles: u64,
    pub total_default_styles: u64,
    pub total_style_cache_hits: u64,
    pub total_style_cache_misses: u64,
    pub total_string_cells: u64,
    pub total_number_cells: u64,
    pub total_bool_cells: u64,
    pub total_error_cells: u64,
    pub total_date_cells: u64,
    pub total_shared_string_cells: u64,
    pub total_inline_string_cells: u64,
    pub total_numfmt_builtin: u64,
    pub total_numfmt_custom: u64,
    pub total_numfmt_general: u64,
    pub total_format_number_calls: u64,
    pub total_format_number_date_calls: u64,
    pub total_format_number_number_calls: u64,
    pub total_value_parse_calls: u64,
    pub total_text_unescape_calls: u64,
    pub total_comments: u64,
    pub total_hyperlinks: u64,
    pub total_data_validations: u64,
    pub total_conditional_formats: u64,
    pub total_drawings: u64,
    pub total_charts: u64,
    pub sheet_metrics: Vec<SheetParseMetrics>,
}

/// Per-sheet timing metrics.
#[derive(Clone, Debug, Default, Serialize)]
pub struct SheetParseMetrics {
    pub name: String,
    pub parse_ms: f64,
    pub comments_ms: f64,
    pub hyperlinks_ms: f64,
    pub drawings_ms: f64,
    pub style_resolve_ms: f64,
    pub format_number_ms: f64,
    pub format_number_date_ms: f64,
    pub format_number_number_ms: f64,
    pub value_parse_ms: f64,
    pub text_unescape_ms: f64,
    pub cell_count: u64,
    pub row_count: u64,
    pub col_count: u64,
    pub merge_count: u64,
    pub style_count: u64,
    pub default_style_count: u64,
    pub style_cache_hits: u64,
    pub style_cache_misses: u64,
    pub string_cells: u64,
    pub number_cells: u64,
    pub bool_cells: u64,
    pub error_cells: u64,
    pub date_cells: u64,
    pub shared_string_cells: u64,
    pub inline_string_cells: u64,
    pub numfmt_builtin: u64,
    pub numfmt_custom: u64,
    pub numfmt_general: u64,
    pub format_number_calls: u64,
    pub format_number_date_calls: u64,
    pub format_number_number_calls: u64,
    pub value_parse_calls: u64,
    pub text_unescape_calls: u64,
    pub comment_count: u64,
    pub hyperlink_count: u64,
    pub data_validation_count: u64,
    pub conditional_format_count: u64,
    pub drawing_count: u64,
    pub chart_count: u64,
}

#[derive(Debug)]
struct NumFmtInfo {
    compiled: CompiledFormat,
    is_builtin: bool,
    is_custom: bool,
    is_general: bool,
}

/// Safely convert f64 to u32 with clamping.
/// The clamp ensures the value is in [0, u32::MAX] before casting.
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
fn f64_to_u32_clamped(v: f64) -> u32 {
    v.clamp(0.0, f64::from(u32::MAX)).floor() as u32
}

/// Get the maximum digit width for a given font name at 11pt, 96 DPI
///
/// This value is used in Excel's column width conversion formula.
/// Falls back to 7.0 (Calibri) if the font is unknown.
fn get_digit_width(font_name: &str) -> f64 {
    match font_name {
        "Calibri" => 7.0,
        "Arial" => 6.5,
        "Times New Roman" => 5.7,
        "Verdana" => 7.5,
        "Consolas" => 7.7,
        "Courier New" => 7.3,
        "Tahoma" => 6.8,
        "Georgia" => 6.2,
        _ => 7.0, // Default to Calibri if unknown
    }
}

#[derive(Clone, Copy)]
struct ParseOptions {
    eager_values: bool,
}

impl ParseOptions {
    fn eager() -> Self {
        Self { eager_values: true }
    }

    fn lazy() -> Self {
        Self {
            eager_values: false,
        }
    }
}

/// Parse an XLSX file from bytes (eager values for JS export compatibility).
pub fn parse(data: &[u8]) -> Result<Workbook> {
    parse_internal(data, None, ParseOptions::eager())
}

/// Parse an XLSX file from bytes (lazy values for viewer performance).
pub fn parse_lazy(data: &[u8]) -> Result<Workbook> {
    parse_internal(data, None, ParseOptions::lazy())
}

/// Parse an XLSX file from bytes and return detailed timing metrics (eager values).
pub fn parse_with_metrics(data: &[u8]) -> Result<(Workbook, ParseMetrics)> {
    let mut metrics = ParseMetrics::default();
    let workbook = parse_internal(data, Some(&mut metrics), ParseOptions::eager())?;
    Ok((workbook, metrics))
}

/// Parse an XLSX file from bytes and return detailed timing metrics (lazy values).
pub fn parse_with_metrics_lazy(data: &[u8]) -> Result<(Workbook, ParseMetrics)> {
    let mut metrics = ParseMetrics::default();
    let workbook = parse_internal(data, Some(&mut metrics), ParseOptions::lazy())?;
    Ok((workbook, metrics))
}

#[allow(clippy::needless_option_as_deref)]
fn parse_internal(
    data: &[u8],
    mut metrics: Option<&mut ParseMetrics>,
    options: ParseOptions,
) -> Result<Workbook> {
    let total_start = now_ms();

    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor)?;

    // Parse workbook relationships first to get actual file paths
    let relationships_start = now_ms();
    let relationships = parse_workbook_relationships(&mut archive);
    if let Some(m) = metrics.as_mut() {
        m.relationships_ms = now_ms() - relationships_start;
    }

    // Parse theme (needed for color resolution)
    // Use path from relationships if available, otherwise fallback to default
    let theme_start = now_ms();
    let theme = parse_theme(&mut archive, relationships.theme.as_deref());
    if let Some(m) = metrics.as_mut() {
        m.theme_ms = now_ms() - theme_start;
    }
    let theme_colors = &theme.colors;

    // Parse shared strings using path from relationships
    let shared_strings_start = now_ms();
    let shared_strings =
        parse_shared_strings(&mut archive, relationships.shared_strings.as_deref());
    if let Some(m) = metrics.as_mut() {
        m.shared_strings_count = shared_strings.len() as u64;
        m.shared_strings_chars = shared_strings.iter().map(|s| s.len() as u64).sum();
    }
    if let Some(m) = metrics.as_mut() {
        m.shared_strings_ms = now_ms() - shared_strings_start;
    }

    // Parse styles using path from relationships
    let styles_start = now_ms();
    let stylesheet = parse_stylesheet(&mut archive, relationships.styles.as_deref())?;
    if let Some(m) = metrics.as_mut() {
        m.styles_fonts = stylesheet.fonts.len() as u64;
        m.styles_fills = stylesheet.fills.len() as u64;
        m.styles_borders = stylesheet.borders.len() as u64;
        m.styles_cell_xfs = stylesheet.cell_xfs.len() as u64;
        m.styles_cell_style_xfs = stylesheet.cell_style_xfs.len() as u64;
        m.styles_num_fmts = stylesheet.num_fmts.len() as u64;
        m.styles_named_styles = stylesheet.named_styles.len() as u64;
        m.styles_dxf = stylesheet.dxf_styles.len() as u64;
        m.styles_indexed_colors = stylesheet
            .indexed_colors
            .as_ref()
            .map(|c| c.len() as u64)
            .unwrap_or(0);
    }
    if let Some(m) = metrics.as_mut() {
        m.styles_ms = now_ms() - styles_start;
    }

    // Get sheet names, paths, states, and date1904 flag from workbook.xml
    let workbook_info_start = now_ms();
    let (sheet_info, date1904) = get_sheet_info(&mut archive, &relationships.worksheets)?;
    if let Some(m) = metrics.as_mut() {
        m.sheets_count = sheet_info.len() as u64;
    }
    if let Some(m) = metrics.as_mut() {
        m.workbook_info_ms = now_ms() - workbook_info_start;
    }

    // Determine the default font for column width calculation
    // Try to get it from theme's minor font, then from Normal style, then fallback to Calibri
    let default_font_name = theme
        .minor_font
        .as_ref()
        .or_else(|| {
            // Try to get the font from the Normal style (first cellStyleXf)
            stylesheet
                .cell_style_xfs
                .first()
                .and_then(|xf| xf.font_id)
                .and_then(|font_id| stylesheet.fonts.get(font_id as usize))
                .and_then(|font| font.name.as_ref())
        })
        .map(String::as_str)
        .unwrap_or("Calibri");
    let digit_width = get_digit_width(default_font_name);

    let style_resolve_start = now_ms();
    let resolved_styles: Vec<Option<StyleRef>> = (0..stylesheet.cell_xfs.len())
        .map(|idx| {
            u32::try_from(idx).ok().and_then(|id| {
                resolve_style(id, &stylesheet, &theme).map(|s| StyleRef(Arc::new(s)))
            })
        })
        .collect();
    let default_style = get_default_style(&stylesheet, &theme).map(|s| StyleRef(Arc::new(s)));
    if let Some(m) = metrics.as_mut() {
        m.style_resolve_ms += now_ms() - style_resolve_start;
    }

    let custom_numfmts: HashMap<u32, &str> = stylesheet
        .num_fmts
        .iter()
        .map(|(id, code)| (*id, code.as_str()))
        .collect();
    let mut numfmt_lookup: Vec<NumFmtInfo> = Vec::with_capacity(stylesheet.cell_xfs.len());
    for xf in &stylesheet.cell_xfs {
        let mut is_builtin = false;
        let mut is_custom = false;
        let compiled = if let Some(id) = xf.num_fmt_id {
            if let Some(code) = get_builtin_format(id) {
                is_builtin = true;
                compile_format_code(code)
            } else if let Some(code) = custom_numfmts.get(&id) {
                is_custom = true;
                compile_format_code(code)
            } else {
                CompiledFormat::General
            }
        } else {
            CompiledFormat::General
        };
        let is_general = matches!(compiled, CompiledFormat::General);
        numfmt_lookup.push(NumFmtInfo {
            compiled,
            is_builtin,
            is_custom,
            is_general,
        });
    }
    let numfmt_cache: Vec<CompiledFormat> = numfmt_lookup
        .iter()
        .map(|info| info.compiled.clone())
        .collect();

    // Parse each sheet
    let mut sheets = Vec::new();
    let sheets_start = now_ms();
    for info in sheet_info {
        let sheet_start = now_ms();
        let mut sheet_metrics = SheetParseMetrics {
            name: info.name.clone(),
            parse_ms: 0.0,
            comments_ms: 0.0,
            hyperlinks_ms: 0.0,
            drawings_ms: 0.0,
            style_resolve_ms: 0.0,
            format_number_ms: 0.0,
            format_number_date_ms: 0.0,
            format_number_number_ms: 0.0,
            value_parse_ms: 0.0,
            text_unescape_ms: 0.0,
            cell_count: 0,
            row_count: 0,
            col_count: 0,
            merge_count: 0,
            style_count: 0,
            default_style_count: 0,
            style_cache_hits: 0,
            style_cache_misses: 0,
            string_cells: 0,
            number_cells: 0,
            bool_cells: 0,
            error_cells: 0,
            date_cells: 0,
            shared_string_cells: 0,
            inline_string_cells: 0,
            numfmt_builtin: 0,
            numfmt_custom: 0,
            numfmt_general: 0,
            format_number_calls: 0,
            format_number_date_calls: 0,
            format_number_number_calls: 0,
            value_parse_calls: 0,
            text_unescape_calls: 0,
            comment_count: 0,
            hyperlink_count: 0,
            data_validation_count: 0,
            conditional_format_count: 0,
            drawing_count: 0,
            chart_count: 0,
        };

        let mut sheet = parse_sheet(
            &mut archive,
            &info,
            &shared_strings,
            &stylesheet,
            &theme,
            date1904,
            digit_width,
            options,
            &numfmt_lookup,
            &resolved_styles,
            &default_style,
            metrics.as_deref_mut().map(|_| &mut sheet_metrics),
        )?;
        sheet_metrics.parse_ms = now_ms() - sheet_start;

        // Parse comments for this sheet
        let comments_start = now_ms();
        if let Some(comments_path) = get_comments_path(&mut archive, &info.path) {
            let comments = parse_comments(
                &mut archive,
                &comments_path,
                theme_colors,
                stylesheet.indexed_colors.as_ref(),
            );
            sheet.comments = comments;
            sheet.rebuild_comment_index();

            // Update cells with has_comment flag
            if !sheet.comments.is_empty() {
                for comment in &sheet.comments {
                    if let Some((col, row)) = parse_cell_ref_str(&comment.cell_ref) {
                        if let Some(idx) = sheet.cell_index_at(row, col) {
                            if let Some(cell_data) = sheet.cells.get_mut(idx) {
                                cell_data.cell.has_comment = Some(true);
                            }
                        }
                    }
                }
            }
        }
        sheet_metrics.comments_ms = now_ms() - comments_start;
        sheet_metrics.comment_count = sheet.comments.len() as u64;

        // Parse hyperlinks for this sheet
        // Re-open the sheet file to parse hyperlinks section
        let hyperlinks_start = now_ms();
        let raw_hyperlinks = if let Ok(file) = archive.by_name(&info.path) {
            let reader = BufReader::new(file);
            let mut xml = Reader::from_reader(reader);
            xml.trim_text(true);

            // Parse raw hyperlinks from the sheet XML
            parse_hyperlinks(&mut xml)
        } else {
            Vec::new()
        };

        if !raw_hyperlinks.is_empty() {
            // Parse hyperlink relationships (for external URLs)
            let rels = parse_hyperlink_rels(&mut archive, &info.path);

            // Resolve hyperlinks with relationship data
            let resolved = resolve_hyperlinks(&raw_hyperlinks, &rels);

            // Update cells with hyperlink data
            for (cell_ref, hyperlink) in &resolved {
                if let Some((col, row)) = parse_cell_ref_str(cell_ref) {
                    if let Some(idx) = sheet.cell_index_at(row, col) {
                        if let Some(cell_data) = sheet.cells.get_mut(idx) {
                            cell_data.cell.hyperlink = Some(hyperlink.clone());
                        }
                    }
                }
            }

            // Store hyperlinks at sheet level
            sheet.hyperlinks = resolved
                .into_iter()
                .map(|(cell_ref, hyperlink)| HyperlinkDef {
                    cell_ref,
                    hyperlink,
                })
                .collect();
        }
        sheet_metrics.hyperlinks_ms = now_ms() - hyperlinks_start;
        sheet_metrics.hyperlink_count = sheet.hyperlinks.len() as u64;

        // Parse drawings (images, charts, shapes) for this sheet
        let drawings_start = now_ms();
        if let Some(drawing_path) = get_drawing_path(&mut archive, &info.path) {
            let (mut drawings, image_rels) = parse_drawing_file(&mut archive, &drawing_path);

            // Resolve image_id (rId) to actual image path for each drawing
            for drawing in &mut drawings {
                if let Some(ref rid) = drawing.image_id {
                    if let Some(image_path) = image_rels.get(rid) {
                        // Replace the rId with the resolved image path
                        drawing.image_id = Some(image_path.clone());
                    }
                }
            }

            sheet.drawings = drawings;

            // Parse charts from the drawing
            let charts = parse_charts_from_drawing(&mut archive, &drawing_path);
            sheet.charts = charts;
        }
        sheet_metrics.drawings_ms = now_ms() - drawings_start;
        sheet_metrics.drawing_count = sheet.drawings.len() as u64;
        sheet_metrics.chart_count = sheet.charts.len() as u64;
        sheet_metrics.data_validation_count = sheet.data_validations.len() as u64;
        sheet_metrics.conditional_format_count = sheet.conditional_formatting.len() as u64;

        sheets.push(sheet);

        if let Some(m) = metrics.as_mut() {
            m.total_cells = m.total_cells.saturating_add(sheet_metrics.cell_count);
            m.total_rows = m.total_rows.saturating_add(sheet_metrics.row_count);
            m.total_cols = m.total_cols.saturating_add(sheet_metrics.col_count);
            m.total_merges = m.total_merges.saturating_add(sheet_metrics.merge_count);
            m.total_styles = m.total_styles.saturating_add(sheet_metrics.style_count);
            m.total_default_styles = m
                .total_default_styles
                .saturating_add(sheet_metrics.default_style_count);
            m.total_style_cache_hits = m
                .total_style_cache_hits
                .saturating_add(sheet_metrics.style_cache_hits);
            m.total_style_cache_misses = m
                .total_style_cache_misses
                .saturating_add(sheet_metrics.style_cache_misses);
            m.total_string_cells = m
                .total_string_cells
                .saturating_add(sheet_metrics.string_cells);
            m.total_number_cells = m
                .total_number_cells
                .saturating_add(sheet_metrics.number_cells);
            m.total_bool_cells = m.total_bool_cells.saturating_add(sheet_metrics.bool_cells);
            m.total_error_cells = m
                .total_error_cells
                .saturating_add(sheet_metrics.error_cells);
            m.total_date_cells = m.total_date_cells.saturating_add(sheet_metrics.date_cells);
            m.total_shared_string_cells = m
                .total_shared_string_cells
                .saturating_add(sheet_metrics.shared_string_cells);
            m.total_inline_string_cells = m
                .total_inline_string_cells
                .saturating_add(sheet_metrics.inline_string_cells);
            m.total_numfmt_builtin = m
                .total_numfmt_builtin
                .saturating_add(sheet_metrics.numfmt_builtin);
            m.total_numfmt_custom = m
                .total_numfmt_custom
                .saturating_add(sheet_metrics.numfmt_custom);
            m.total_numfmt_general = m
                .total_numfmt_general
                .saturating_add(sheet_metrics.numfmt_general);
            m.total_format_number_calls = m
                .total_format_number_calls
                .saturating_add(sheet_metrics.format_number_calls);
            m.total_format_number_date_calls = m
                .total_format_number_date_calls
                .saturating_add(sheet_metrics.format_number_date_calls);
            m.total_format_number_number_calls = m
                .total_format_number_number_calls
                .saturating_add(sheet_metrics.format_number_number_calls);
            m.total_value_parse_calls = m
                .total_value_parse_calls
                .saturating_add(sheet_metrics.value_parse_calls);
            m.total_text_unescape_calls = m
                .total_text_unescape_calls
                .saturating_add(sheet_metrics.text_unescape_calls);
            m.style_resolve_ms += sheet_metrics.style_resolve_ms;
            m.format_number_ms += sheet_metrics.format_number_ms;
            m.format_number_date_ms += sheet_metrics.format_number_date_ms;
            m.format_number_number_ms += sheet_metrics.format_number_number_ms;
            m.value_parse_ms += sheet_metrics.value_parse_ms;
            m.text_unescape_ms += sheet_metrics.text_unescape_ms;
            m.total_comments = m.total_comments.saturating_add(sheet_metrics.comment_count);
            m.total_hyperlinks = m
                .total_hyperlinks
                .saturating_add(sheet_metrics.hyperlink_count);
            m.total_data_validations = m
                .total_data_validations
                .saturating_add(sheet_metrics.data_validation_count);
            m.total_conditional_formats = m
                .total_conditional_formats
                .saturating_add(sheet_metrics.conditional_format_count);
            m.total_drawings = m.total_drawings.saturating_add(sheet_metrics.drawing_count);
            m.total_charts = m.total_charts.saturating_add(sheet_metrics.chart_count);
            m.sheet_metrics.push(sheet_metrics);
        }
    }
    if let Some(m) = metrics.as_mut() {
        m.sheets_ms = now_ms() - sheets_start;
    }

    // Resolve chart data references (formulas -> actual cell values)
    // This must happen after all sheets are parsed so cross-sheet refs work
    let charts_start = now_ms();
    resolve_all_chart_data(&mut sheets, &shared_strings);
    if let Some(m) = metrics.as_mut() {
        m.charts_resolve_ms = now_ms() - charts_start;
    }

    // Collect all unique image paths from all sheets and read image data
    let images_start = now_ms();
    let images = collect_and_read_images(&mut archive, &sheets);
    if let Some(m) = metrics.as_mut() {
        m.images_ms = now_ms() - images_start;
    }

    // Parse DXF styles from stylesheet
    let dxf_start = now_ms();
    let dxf_styles = stylesheet.dxf_styles.clone();
    if let Some(m) = metrics.as_mut() {
        m.dxf_ms = now_ms() - dxf_start;
    }

    if let Some(m) = metrics.as_mut() {
        m.parse_ms = now_ms() - total_start;
    }

    Ok(Workbook {
        sheets,
        theme,
        defined_names: Vec::new(),
        date1904,
        images,
        dxf_styles,
        shared_strings,
        resolved_styles,
        default_style,
        numfmt_cache,
    })
}

/// Resolve chart data references for all sheets
///
/// This iterates through all sheets and resolves their chart formula references
/// (like `'Sheet1'!$A$1:$A$5`) to actual cell values.
fn resolve_all_chart_data(sheets: &mut [Sheet], shared_strings: &[String]) {
    use crate::formula::parse_formula_ref;
    use crate::types::ChartDataRef;

    // Fast path: if no charts exist, skip expensive cell lookup construction.
    if sheets.iter().all(|s| s.charts.is_empty()) {
        return;
    }

    #[derive(Clone)]
    struct ChartCellValue {
        num: Option<f64>,
        text: Option<String>,
    }

    fn chart_cell_value(cell: &crate::types::Cell, shared_strings: &[String]) -> ChartCellValue {
        if let Some(raw) = cell.raw.as_ref() {
            match raw {
                crate::types::CellRawValue::Number(n) | crate::types::CellRawValue::Date(n) => {
                    ChartCellValue {
                        num: Some(*n),
                        text: cell.cached_display.clone().or_else(|| Some(n.to_string())),
                    }
                }
                crate::types::CellRawValue::Boolean(b) => ChartCellValue {
                    num: None,
                    text: Some(if *b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }),
                },
                crate::types::CellRawValue::Error(e) => ChartCellValue {
                    num: None,
                    text: Some(e.clone()),
                },
                crate::types::CellRawValue::String(s) => ChartCellValue {
                    num: s.parse::<f64>().ok(),
                    text: Some(s.clone()),
                },
                crate::types::CellRawValue::SharedString(idx) => {
                    let text = shared_strings.get(*idx as usize).cloned();
                    ChartCellValue {
                        num: text.as_deref().and_then(|s| s.parse::<f64>().ok()),
                        text,
                    }
                }
            }
        } else if let Some(v) = cell.v.as_ref() {
            ChartCellValue {
                num: v.parse::<f64>().ok(),
                text: Some(v.clone()),
            }
        } else {
            ChartCellValue {
                num: None,
                text: None,
            }
        }
    }

    // Build a lookup map of sheet name -> cell values for efficient access
    // We store lightweight numeric/text values to avoid formatting during chart resolve.
    let cell_lookup: std::collections::HashMap<
        String,
        std::collections::HashMap<(u32, u32), ChartCellValue>,
    > = sheets
        .iter()
        .map(|s| {
            let mut cells = std::collections::HashMap::with_capacity(s.cells.len());
            for cell_data in &s.cells {
                cells.insert(
                    (cell_data.r, cell_data.c),
                    chart_cell_value(&cell_data.cell, shared_strings),
                );
            }
            (s.name.clone(), cells)
        })
        .collect();

    // Helper to get cell value
    let get_cell_value = |sheet_name: &str, row: u32, col: u32| -> Option<f64> {
        cell_lookup
            .get(sheet_name)
            .and_then(|cells| cells.get(&(row, col)).and_then(|v| v.num))
    };

    // Helper to get cell string
    let get_cell_string = |sheet_name: &str, row: u32, col: u32| -> Option<String> {
        cell_lookup
            .get(sheet_name)
            .and_then(|cells| cells.get(&(row, col)).and_then(|v| v.text.clone()))
    };

    // Helper to resolve numeric values from a formula reference
    let resolve_numeric = |formula: &str, current_sheet: &str| -> Vec<Option<f64>> {
        let Some(fref) = parse_formula_ref(formula) else {
            return Vec::new();
        };
        let sheet_name = fref.sheet_name.unwrap_or(current_sheet);
        let mut values = Vec::new();

        if fref.col_start == fref.col_end {
            for row in fref.row_start..=fref.row_end {
                values.push(get_cell_value(sheet_name, row, fref.col_start));
            }
        } else if fref.row_start == fref.row_end {
            for col in fref.col_start..=fref.col_end {
                values.push(get_cell_value(sheet_name, fref.row_start, col));
            }
        } else {
            for row in fref.row_start..=fref.row_end {
                for col in fref.col_start..=fref.col_end {
                    values.push(get_cell_value(sheet_name, row, col));
                }
            }
        }
        values
    };

    // Helper to resolve string values from a formula reference
    let resolve_strings = |formula: &str, current_sheet: &str| -> Vec<String> {
        let Some(fref) = parse_formula_ref(formula) else {
            return Vec::new();
        };
        let sheet_name = fref.sheet_name.unwrap_or(current_sheet);
        let mut values = Vec::new();

        if fref.col_start == fref.col_end {
            for row in fref.row_start..=fref.row_end {
                if let Some(s) = get_cell_string(sheet_name, row, fref.col_start) {
                    values.push(s);
                }
            }
        } else if fref.row_start == fref.row_end {
            for col in fref.col_start..=fref.col_end {
                if let Some(s) = get_cell_string(sheet_name, fref.row_start, col) {
                    values.push(s);
                }
            }
        } else {
            for row in fref.row_start..=fref.row_end {
                for col in fref.col_start..=fref.col_end {
                    if let Some(s) = get_cell_string(sheet_name, row, col) {
                        values.push(s);
                    }
                }
            }
        }
        values
    };

    // Helper to resolve a data ref
    let resolve_data_ref = |data_ref: &mut ChartDataRef, current_sheet: &str| {
        if let Some(ref formula) = data_ref.formula {
            data_ref.num_values = resolve_numeric(formula, current_sheet);
            if data_ref.str_values.is_empty() {
                data_ref.str_values = resolve_strings(formula, current_sheet);
            }
        }
    };

    // Now iterate and resolve
    for sheet in sheets.iter_mut() {
        let sheet_name = sheet.name.clone();

        for chart in &mut sheet.charts {
            for series in &mut chart.series {
                // Resolve series values (Y-axis data)
                if let Some(ref mut values) = series.values {
                    resolve_data_ref(values, &sheet_name);
                }

                // Resolve series categories (X-axis labels)
                if let Some(ref mut categories) = series.categories {
                    resolve_data_ref(categories, &sheet_name);
                }

                // Resolve X values (for scatter charts)
                if let Some(ref mut x_values) = series.x_values {
                    resolve_data_ref(x_values, &sheet_name);
                }

                // Resolve bubble sizes (for bubble charts)
                if let Some(ref mut bubble_sizes) = series.bubble_sizes {
                    resolve_data_ref(bubble_sizes, &sheet_name);
                }

                // Resolve series name from cell reference
                if series.name.is_none() {
                    if let Some(ref name_ref) = series.name_ref {
                        if let Some(fref) = parse_formula_ref(name_ref) {
                            let sn = fref.sheet_name.unwrap_or(&sheet_name);
                            series.name = get_cell_string(sn, fref.row_start, fref.col_start);
                        }
                    }
                }
            }
        }
    }
}

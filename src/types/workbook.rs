use serde::{Deserialize, Serialize};
use std::collections::HashMap;

use super::*;
use crate::numfmt::CompiledFormat;

/// A comment (note) attached to a cell
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Comment {
    /// The cell reference this comment is attached to (e.g., "A1")
    pub cell_ref: String,
    /// The author of the comment (if available)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// The plain text content of the comment
    pub text: String,
    /// Rich text runs if the comment has formatting
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// A complete Excel workbook
#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub theme: Theme,
    /// Named ranges and defined names from the workbook
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub defined_names: Vec<DefinedName>,
    /// Whether the workbook uses the 1904 date system (Mac default)
    /// If false, uses the 1900 date system (Windows default)
    #[serde(skip_serializing_if = "is_false")]
    pub date1904: bool,
    /// Embedded images from xl/media/ folder (keyed by path)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<EmbeddedImage>,
    /// Differential formatting styles (DXF) for conditional formatting
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub dxf_styles: Vec<DxfStyle>,
    /// Shared string table (lazy cell values; not serialized).
    #[serde(skip)]
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    pub(crate) shared_strings: Vec<String>,
    /// Resolved styles indexed by cell_xf (lazy path; not serialized).
    #[serde(skip)]
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    pub(crate) resolved_styles: Vec<Option<StyleRef>>,
    /// Default style for cells without explicit styling (lazy path; not serialized).
    #[serde(skip)]
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    pub(crate) default_style: Option<StyleRef>,
    /// Compiled number formats indexed by cell_xf (lazy path; not serialized).
    #[serde(skip)]
    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    pub(crate) numfmt_cache: Vec<CompiledFormat>,
}

/// Helper function for serde skip_serializing_if
pub(crate) fn is_false(b: &bool) -> bool {
    !b
}

/// A defined name (named range) in the workbook
///
/// Named ranges can reference cell ranges, single cells, constants, or formulas.
/// Built-in names use the `_xlnm.` prefix:
/// - `_xlnm.Print_Area` - Print area
/// - `_xlnm.Print_Titles` - Print titles
/// - `_xlnm.Database` - Database range
/// - `_xlnm.Criteria` - Criteria range for filtering
/// - `_xlnm.Extract` - Extract range
/// - `_xlnm._FilterDatabase` - AutoFilter database
#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    /// The name (can include special _xlnm. prefix for built-in names)
    pub name: String,
    /// The formula/reference value (e.g., "Sheet1!$A$1:$D$10" or "0.0825")
    pub value: String,
    /// If present, the name is scoped to this sheet (0-indexed)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub local_sheet_id: Option<u32>,
    /// Whether this name is hidden from the Name Manager UI
    pub hidden: bool,
    /// Optional comment/description for this name
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment: Option<String>,
}

/// Sheet visibility state
#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SheetState {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

/// A single worksheet
#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Sheet {
    pub name: String,
    pub state: SheetState,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>, // #RRGGBB
    /// Sparse representation: Vec of (row, col, cell)
    pub cells: Vec<CellData>,
    /// Row index for fast cell lookup/rendering (not serialized).
    #[serde(skip)]
    pub(crate) cells_by_row: Vec<Vec<usize>>,
    pub merges: Vec<MergeRange>,
    pub col_widths: Vec<ColWidth>,
    pub row_heights: Vec<RowHeight>,
    pub default_col_width: f64,
    pub default_row_height: f64,
    pub hidden_cols: Vec<u32>,
    pub hidden_rows: Vec<u32>,
    pub max_row: u32,
    pub max_col: u32,
    /// Number of frozen rows (0 = none)
    pub frozen_rows: u32,
    /// Number of frozen columns (0 = none)
    pub frozen_cols: u32,
    /// Split position in points for rows (for split panes)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub split_row: Option<f64>,
    /// Split position in points for columns (for split panes)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub split_col: Option<f64>,
    /// Pane state (frozen, frozenSplit, or split)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pane_state: Option<PaneState>,
    /// Whether the sheet is protected (cells with locked=true are read-only)
    pub is_protected: bool,
    /// Data validations applied to cell ranges
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub data_validations: Vec<DataValidationRange>,
    /// Auto-filter settings for the sheet
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<AutoFilter>,
    /// Row outline levels (for grouping/collapsing)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub outline_level_row: Vec<OutlineLevel>,
    /// Column outline levels (for grouping/collapsing)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub outline_level_col: Vec<OutlineLevel>,
    /// Summary rows are below detail rows (default true)
    pub outline_summary_below: bool,
    /// Summary columns are right of detail columns (default true)
    pub outline_summary_right: bool,
    /// Hyperlinks defined in this sheet (for bulk access)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<HyperlinkDef>,
    /// Comments/notes attached to cells in this sheet
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
    /// Comment lookup by cell reference (not serialized).
    #[serde(skip)]
    pub(crate) comments_by_cell: HashMap<String, usize>,
    /// Print area range, e.g., "A1:H50"
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_area: Option<String>,
    /// Row indices where manual page breaks occur (0-based)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub row_breaks: Vec<u32>,
    /// Column indices where manual page breaks occur (0-based)
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub col_breaks: Vec<u32>,
    /// Rows to repeat at top of each printed page (start, end) - 0-based
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_titles_rows: Option<(u32, u32)>,
    /// Columns to repeat at left of each printed page (start, end) - 0-based
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_titles_cols: Option<(u32, u32)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<PageMargins>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<PageSetup>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<HeaderFooter>,
    /// Sparkline groups defined in this sheet
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub sparkline_groups: Vec<SparklineGroup>,
    /// Drawings (images, charts, shapes) in this sheet
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub drawings: Vec<Drawing>,
    /// Charts embedded in this sheet
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub charts: Vec<Chart>,
    /// Conditional formatting rules applied to cell ranges
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub conditional_formatting: Vec<ConditionalFormatting>,
    /// Cached conditional formatting metadata (parsed ranges, sorted rule indices)
    #[serde(skip)]
    pub(crate) conditional_formatting_cache: Vec<ConditionalFormattingCache>,
}

impl Sheet {
    pub(crate) fn rebuild_cell_index(&mut self) {
        if self.cells.is_empty() {
            self.cells_by_row = Vec::new();
            return;
        }

        let max_row = self.cells.iter().map(|c| c.r).max().unwrap_or(0) as usize;
        let mut rows: Vec<Vec<usize>> = vec![Vec::new(); max_row + 1];

        for (idx, cell) in self.cells.iter().enumerate() {
            let row = cell.r as usize;
            if let Some(row_cells) = rows.get_mut(row) {
                row_cells.push(idx);
            }
        }

        for row_cells in &mut rows {
            row_cells.sort_by_key(|&i| self.cells.get(i).map(|cell| cell.c).unwrap_or(u32::MAX));
        }

        self.cells_by_row = rows;
    }

    pub(crate) fn rebuild_comment_index(&mut self) {
        self.comments_by_cell.clear();
        if self.comments.is_empty() {
            return;
        }

        self.comments_by_cell.reserve(self.comments.len());
        for (idx, comment) in self.comments.iter().enumerate() {
            self.comments_by_cell
                .entry(comment.cell_ref.clone())
                .or_insert(idx);
        }
    }

    #[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
    pub(crate) fn cell_index_at(&self, row: u32, col: u32) -> Option<usize> {
        if self.cells_by_row.is_empty() {
            return self.cells.iter().position(|c| c.r == row && c.c == col);
        }
        let row_cells = self.cells_by_row.get(row as usize)?;
        let pos = row_cells
            .partition_point(|&i| self.cells.get(i).map(|cell| cell.c < col).unwrap_or(false));
        let idx = row_cells.get(pos).copied()?;
        self.cells
            .get(idx)
            .is_some_and(|cell| cell.c == col)
            .then_some(idx)
    }
}

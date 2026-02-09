//! Formula reference parsing and resolution for chart data
//!
//! Handles parsing Excel formula references like `'Sheet Name'!$A$1:$B$5`
//! and resolving them to actual cell values.

use crate::types::{Cell, CellRawValue, Chart, ChartDataRef, Sheet};

/// Parsed formula reference
#[derive(Debug)]
pub struct FormulaRef<'a> {
    /// Sheet name (None means current sheet)
    pub sheet_name: Option<&'a str>,
    /// Starting row (0-indexed)
    pub row_start: u32,
    /// Ending row (0-indexed)
    pub row_end: u32,
    /// Starting column (0-indexed)
    pub col_start: u32,
    /// Ending column (0-indexed)
    pub col_end: u32,
}

/// Parse a formula reference like "'Sheet Name'!$A$1:$B$5" or "A1:B5"
///
/// Returns the parsed reference with sheet name and cell range
pub fn parse_formula_ref(formula: &str) -> Option<FormulaRef<'_>> {
    let formula = formula.trim();

    // Check if there's a sheet reference (contains '!')
    let (sheet_name, range_part) = if let Some(excl_pos) = formula.rfind('!') {
        let sheet_part = &formula[..excl_pos];
        let range_part = &formula[excl_pos + 1..];

        // Handle quoted sheet names: 'My Sheet' or plain: Sheet1
        let sheet_name = if sheet_part.starts_with('\'') && sheet_part.ends_with('\'') {
            // Remove quotes
            Some(&sheet_part[1..sheet_part.len() - 1])
        } else {
            Some(sheet_part)
        };

        (sheet_name, range_part)
    } else {
        (None, formula)
    };

    // Parse the cell range
    let (row_start, row_end, col_start, col_end) = parse_cell_range(range_part)?;

    Some(FormulaRef {
        sheet_name,
        row_start,
        row_end,
        col_start,
        col_end,
    })
}

/// Parse a cell range string like "A1:C5" or "A1" into (min_row, max_row, min_col, max_col)
/// Returns 0-based row and column indices
#[allow(clippy::indexing_slicing)] // Safe: we check parts.len() before indexing
fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    if let Some((start, end)) = range.split_once(':') {
        let (row1, col1) = parse_cell_ref(start)?;
        let (row2, col2) = parse_cell_ref(end)?;
        Some((
            row1.min(row2),
            row1.max(row2),
            col1.min(col2),
            col1.max(col2),
        ))
    } else {
        let (row, col) = parse_cell_ref(range)?;
        Some((row, row, col, col))
    }
}

/// Parse a cell reference like "A1" or "$B$10" into (row, col) as 0-based indices
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let (col, row) = crate::cell_ref::parse_cell_ref(cell_ref)?;
    Some((row, col))
}

/// Look up a cell value from a sheet
fn get_cell_value(sheet: &Sheet, row: u32, col: u32, shared_strings: &[String]) -> Option<f64> {
    let idx = sheet.cell_index_at(row, col)?;
    sheet
        .cells
        .get(idx)
        .and_then(|cell_data| cell_value_to_f64(&cell_data.cell, shared_strings))
}

/// Convert a cell's value to f64
fn cell_value_to_f64(cell: &Cell, shared_strings: &[String]) -> Option<f64> {
    if let Some(raw) = cell.raw.as_ref() {
        return match raw {
            CellRawValue::Number(n) | CellRawValue::Date(n) => Some(*n),
            CellRawValue::String(s) => s.parse::<f64>().ok(),
            CellRawValue::SharedString(idx) => shared_strings
                .get(*idx as usize)
                .and_then(|s| s.parse::<f64>().ok()),
            CellRawValue::Boolean(_) | CellRawValue::Error(_) => None,
        };
    }
    cell.v.as_ref().and_then(|v| v.parse::<f64>().ok())
}

/// Look up a cell's string value from a sheet
fn get_cell_string(sheet: &Sheet, row: u32, col: u32, shared_strings: &[String]) -> Option<String> {
    let idx = sheet.cell_index_at(row, col)?;
    sheet
        .cells
        .get(idx)
        .and_then(|cell_data| cell_value_to_string(&cell_data.cell, shared_strings))
}

fn cell_value_to_string(cell: &Cell, shared_strings: &[String]) -> Option<String> {
    if let Some(raw) = cell.raw.as_ref() {
        return match raw {
            CellRawValue::SharedString(idx) => shared_strings.get(*idx as usize).cloned(),
            CellRawValue::String(s) => Some(s.clone()),
            CellRawValue::Number(n) | CellRawValue::Date(n) => {
                cell.cached_display.clone().or_else(|| Some(n.to_string()))
            }
            CellRawValue::Boolean(b) => Some(if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            CellRawValue::Error(e) => Some(e.clone()),
        };
    }
    cell.v.clone()
}

/// Resolve a formula reference to numeric values
fn resolve_numeric_values(
    formula_ref: &FormulaRef<'_>,
    sheets: &[Sheet],
    current_sheet_name: &str,
    shared_strings: &[String],
) -> Vec<Option<f64>> {
    // Find the target sheet
    let sheet_name = formula_ref.sheet_name.unwrap_or(current_sheet_name);
    let sheet = sheets.iter().find(|s| s.name == sheet_name);

    let Some(sheet) = sheet else {
        return Vec::new();
    };

    let row_count = formula_ref.row_end.saturating_sub(formula_ref.row_start) + 1;
    let col_count = formula_ref.col_end.saturating_sub(formula_ref.col_start) + 1;
    let cap = (row_count as usize).saturating_mul(col_count as usize);
    let mut values = Vec::with_capacity(cap);

    // Determine if this is a row-wise or column-wise range
    if formula_ref.col_start == formula_ref.col_end {
        // Single column, iterate rows
        for row in formula_ref.row_start..=formula_ref.row_end {
            values.push(get_cell_value(
                sheet,
                row,
                formula_ref.col_start,
                shared_strings,
            ));
        }
    } else if formula_ref.row_start == formula_ref.row_end {
        // Single row, iterate columns
        for col in formula_ref.col_start..=formula_ref.col_end {
            values.push(get_cell_value(
                sheet,
                formula_ref.row_start,
                col,
                shared_strings,
            ));
        }
    } else {
        // 2D range - iterate row by row (Excel's default order)
        for row in formula_ref.row_start..=formula_ref.row_end {
            for col in formula_ref.col_start..=formula_ref.col_end {
                values.push(get_cell_value(sheet, row, col, shared_strings));
            }
        }
    }

    values
}

/// Resolve a formula reference to string values (for categories)
fn resolve_string_values(
    formula_ref: &FormulaRef<'_>,
    sheets: &[Sheet],
    current_sheet_name: &str,
    shared_strings: &[String],
) -> Vec<String> {
    let sheet_name = formula_ref.sheet_name.unwrap_or(current_sheet_name);
    let sheet = sheets.iter().find(|s| s.name == sheet_name);

    let Some(sheet) = sheet else {
        return Vec::new();
    };

    let row_count = formula_ref.row_end.saturating_sub(formula_ref.row_start) + 1;
    let col_count = formula_ref.col_end.saturating_sub(formula_ref.col_start) + 1;
    let cap = (row_count as usize).saturating_mul(col_count as usize);
    let mut values = Vec::with_capacity(cap);

    // Determine if this is a row-wise or column-wise range
    if formula_ref.col_start == formula_ref.col_end {
        // Single column, iterate rows
        for row in formula_ref.row_start..=formula_ref.row_end {
            if let Some(s) = get_cell_string(sheet, row, formula_ref.col_start, shared_strings) {
                values.push(s);
            }
        }
    } else if formula_ref.row_start == formula_ref.row_end {
        // Single row, iterate columns
        for col in formula_ref.col_start..=formula_ref.col_end {
            if let Some(s) = get_cell_string(sheet, formula_ref.row_start, col, shared_strings) {
                values.push(s);
            }
        }
    } else {
        // 2D range - iterate row by row
        for row in formula_ref.row_start..=formula_ref.row_end {
            for col in formula_ref.col_start..=formula_ref.col_end {
                if let Some(s) = get_cell_string(sheet, row, col, shared_strings) {
                    values.push(s);
                }
            }
        }
    }

    values
}

/// Resolve a ChartDataRef's formula to actual values
fn resolve_data_ref(
    data_ref: &mut ChartDataRef,
    sheets: &[Sheet],
    current_sheet_name: &str,
    shared_strings: &[String],
) {
    if let Some(ref formula) = data_ref.formula {
        if let Some(formula_ref) = parse_formula_ref(formula) {
            // Resolve numeric values
            data_ref.num_values =
                resolve_numeric_values(&formula_ref, sheets, current_sheet_name, shared_strings);

            // Also resolve string values if this might be categories
            if data_ref.str_values.is_empty() {
                data_ref.str_values =
                    resolve_string_values(&formula_ref, sheets, current_sheet_name, shared_strings);
            }
        }
    }
}

/// Resolve all chart data references in a workbook
///
/// This should be called after all sheets have been parsed, so that
/// cross-sheet references can be resolved.
pub fn resolve_chart_data(
    charts: &mut [Chart],
    sheets: &[Sheet],
    current_sheet_name: &str,
    shared_strings: &[String],
) {
    for chart in charts {
        for series in &mut chart.series {
            // Resolve series values (Y-axis data)
            if let Some(ref mut values) = series.values {
                resolve_data_ref(values, sheets, current_sheet_name, shared_strings);
            }

            // Resolve series categories (X-axis labels)
            if let Some(ref mut categories) = series.categories {
                resolve_data_ref(categories, sheets, current_sheet_name, shared_strings);
            }

            // Resolve X values (for scatter charts)
            if let Some(ref mut x_values) = series.x_values {
                resolve_data_ref(x_values, sheets, current_sheet_name, shared_strings);
            }

            // Resolve bubble sizes (for bubble charts)
            if let Some(ref mut bubble_sizes) = series.bubble_sizes {
                resolve_data_ref(bubble_sizes, sheets, current_sheet_name, shared_strings);
            }

            // Resolve series name from cell reference
            if series.name.is_none() {
                if let Some(ref name_ref) = series.name_ref {
                    if let Some(formula_ref) = parse_formula_ref(name_ref) {
                        let sheet_name = formula_ref.sheet_name.unwrap_or(current_sheet_name);
                        if let Some(sheet) = sheets.iter().find(|s| s.name == sheet_name) {
                            series.name = get_cell_string(
                                sheet,
                                formula_ref.row_start,
                                formula_ref.col_start,
                                shared_strings,
                            );
                        }
                    }
                }
            }
        }
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
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("$A$1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("$B$10"), Some((9, 1)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
    }

    #[test]
    fn test_parse_cell_range() {
        assert_eq!(parse_cell_range("A1"), Some((0, 0, 0, 0)));
        assert_eq!(parse_cell_range("A1:B5"), Some((0, 4, 0, 1)));
        assert_eq!(parse_cell_range("$A$1:$B$5"), Some((0, 4, 0, 1)));
        assert_eq!(parse_cell_range("B5:A1"), Some((0, 4, 0, 1))); // Reversed
    }

    #[test]
    fn test_parse_formula_ref_simple() {
        let result = parse_formula_ref("A1:B5").unwrap();
        assert_eq!(result.sheet_name, None);
        assert_eq!(result.row_start, 0);
        assert_eq!(result.row_end, 4);
        assert_eq!(result.col_start, 0);
        assert_eq!(result.col_end, 1);
    }

    #[test]
    fn test_parse_formula_ref_with_sheet() {
        let result = parse_formula_ref("Sheet1!A1:B5").unwrap();
        assert_eq!(result.sheet_name, Some("Sheet1"));
        assert_eq!(result.row_start, 0);
        assert_eq!(result.row_end, 4);
    }

    #[test]
    fn test_parse_formula_ref_quoted_sheet() {
        let result = parse_formula_ref("'My Sheet'!$A$1:$B$5").unwrap();
        assert_eq!(result.sheet_name, Some("My Sheet"));
        assert_eq!(result.row_start, 0);
        assert_eq!(result.row_end, 4);
    }

    #[test]
    fn test_parse_formula_ref_single_cell() {
        let result = parse_formula_ref("'Charts'!B1").unwrap();
        assert_eq!(result.sheet_name, Some("Charts"));
        assert_eq!(result.row_start, 0);
        assert_eq!(result.row_end, 0);
        assert_eq!(result.col_start, 1);
        assert_eq!(result.col_end, 1);
    }
}

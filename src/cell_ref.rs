//! Utilities for parsing Excel-style cell references and ranges.

/// Parse a cell reference like "A1" into (col, row) where col and row are 0-indexed.
pub fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut saw_col = false;
    let mut saw_row = false;

    for ch in cell_ref.trim().chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() {
            let upper = ch.to_ascii_uppercase();
            col = col * 26 + (upper as u32 - 'A' as u32 + 1);
            saw_col = true;
        } else if ch.is_ascii_digit() {
            row = row * 10 + (ch as u32 - '0' as u32);
            saw_row = true;
        }
    }

    if !saw_col || !saw_row {
        return None;
    }

    Some((col.saturating_sub(1), row.saturating_sub(1)))
}

/// Parse a cell range like "A1:B10" or "A1" into (start_row, start_col, end_row, end_col).
pub fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    if let Some((start, end)) = range.split_once(':') {
        let (start_col, start_row) = parse_cell_ref(start)?;
        let (end_col, end_row) = parse_cell_ref(end)?;
        Some((start_row, start_col, end_row, end_col))
    } else {
        let (start_col, start_row) = parse_cell_ref(range)?;
        Some((start_row, start_col, start_row, start_col))
    }
}

/// Parse a cell reference from raw bytes (ASCII) into (col, row) where col and row are 0-indexed.
///
/// This is the bytes equivalent of [`parse_cell_ref`] for use when working with
/// raw XML attribute values (e.g., `attr.value` from quick-xml).
pub fn parse_cell_ref_bytes(ref_bytes: &[u8]) -> Option<(u32, u32)> {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut saw_col = false;
    let mut saw_row = false;

    for &b in ref_bytes {
        if b == b'$' {
            continue;
        }
        if b.is_ascii_alphabetic() {
            let upper = if b.is_ascii_lowercase() { b - 32 } else { b };
            col = col * 26 + (u32::from(upper - b'A') + 1);
            saw_col = true;
        } else if b.is_ascii_digit() {
            row = row * 10 + u32::from(b - b'0');
            saw_row = true;
        }
    }

    if !saw_col || !saw_row {
        return None;
    }

    Some((col.saturating_sub(1), row.saturating_sub(1)))
}

/// Parse a cell reference like "A1" into (col, row) with defaults.
///
/// Returns `(0, 0)` if parsing fails. This is a convenience wrapper
/// for callers that don't need to distinguish between invalid input and cell A1.
pub fn parse_cell_ref_or_default(ref_str: &str) -> (u32, u32) {
    parse_cell_ref(ref_str).unwrap_or((0, 0))
}

/// Parse a cell reference from bytes with defaults.
///
/// Returns `(0, 0)` if parsing fails.
pub fn parse_cell_ref_bytes_or_default(ref_bytes: &[u8]) -> (u32, u32) {
    parse_cell_ref_bytes(ref_bytes).unwrap_or((0, 0))
}

/// Parse sqref string into a list of (start_row, start_col, end_row, end_col) ranges.
pub fn parse_sqref(sqref: &str) -> Vec<(u32, u32, u32, u32)> {
    let mut ranges = Vec::new();

    for part in sqref.split_whitespace() {
        if let Some(range) = parse_cell_range(part) {
            ranges.push(range);
        }
    }

    ranges
}

# BUG-018: Default Row Height Hardcoded

**Priority**: MEDIUM
**Status**: Open
**Component**: parser.rs

## Problem

If `sheetFormatPr` omits `defaultRowHeight`, we keep a hardcoded `20px` default. Excel derives the default row height from the workbook's default font (Normal style) and its point size.

## Current Behavior

```rust
// Sheet initialization
default_row_height: 20.0, // ~15 points

// Only updated when defaultRowHeight exists
```

## Expected Behavior

When `defaultRowHeight` is not specified, compute it from the default font size (Normal style) and Excel's row height rules, then convert to pixels.

## Where Data Should Come From

- `xl/styles.xml` Normal style font size
- `xl/worksheets/sheetN.xml` `<sheetFormatPr defaultRowHeight="..."/>` when present

## Impact

- Row heights are incorrect for workbooks with non-default fonts
- Vertical alignment and spacing diverge from Excel

## Proposed Fix

1. Determine the default font size from `cellStyleXfs[0]` / `cellStyles` (Normal).
2. Derive default row height when `defaultRowHeight` is absent.
3. Convert points to pixels with the correct DPI assumptions.

## References

- ECMA-376 Part 1, Section 18.3.1.62 (sheetFormatPr)

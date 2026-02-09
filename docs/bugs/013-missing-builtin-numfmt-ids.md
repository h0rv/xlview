# BUG-013: Missing Built-in Number Format IDs

**Priority**: MEDIUM
**Status**: Open
**Component**: numfmt.rs

## Problem

Excel has 49+ built-in number format IDs (0-49), but only a subset is implemented. Missing formats may cause cells to display raw numbers instead of formatted values.

## Current Behavior

```rust
// numfmt.rs - get_builtin_format function
match id {
    0 => Some("General"),
    1 => Some("0"),
    2 => Some("0.00"),
    // ... partial list
    _ => None,  // Many IDs return None
}
```

Missing IDs include:
- 5-8: Currency with thousands separator
- 27-36: Date formats (locale-specific, Japanese calendar)
- 41-44: Accounting formats without currency symbol
- 45-49: Time formats (mm:ss, [h]:mm:ss, mm:ss.0)

## Expected Behavior

Support all standard built-in format IDs per ECMA-376.

## Built-in Format Reference

| ID | Format Code | Description |
|----|-------------|-------------|
| 0 | General | General |
| 1 | 0 | Integer |
| 2 | 0.00 | 2 decimals |
| 3 | #,##0 | Thousands |
| 4 | #,##0.00 | Thousands + decimals |
| 5 | $#,##0_);($#,##0) | Currency |
| 6 | $#,##0_);[Red]($#,##0) | Currency negative red |
| 7 | $#,##0.00_);($#,##0.00) | Currency 2 dec |
| 8 | $#,##0.00_);[Red]($#,##0.00) | Currency 2 dec neg red |
| 9 | 0% | Percent |
| 10 | 0.00% | Percent 2 dec |
| 11 | 0.00E+00 | Scientific |
| 12 | # ?/? | Fraction |
| 13 | # ??/?? | Fraction 2 digit |
| 14 | m/d/yyyy | Date |
| 15 | d-mmm-yy | Date |
| 16 | d-mmm | Date (day-month) |
| 17 | mmm-yy | Date (month-year) |
| 18 | h:mm AM/PM | Time 12hr |
| 19 | h:mm:ss AM/PM | Time 12hr with sec |
| 20 | h:mm | Time 24hr |
| 21 | h:mm:ss | Time 24hr with sec |
| 22 | m/d/yyyy h:mm | DateTime |
| 27-36 | (locale-specific) | East Asian dates |
| 37 | #,##0_);(#,##0) | Number with parens |
| 38 | #,##0_);[Red](#,##0) | Number neg red |
| 39 | #,##0.00_);(#,##0.00) | Number 2 dec parens |
| 40 | #,##0.00_);[Red](#,##0.00) | Number 2 dec neg red |
| 41-44 | Accounting variants | Accounting no symbol |
| 45 | mm:ss | Minutes:seconds |
| 46 | [h]:mm:ss | Elapsed time |
| 47 | mm:ss.0 | Minutes:seconds.tenths |
| 48 | ##0.0E+0 | Scientific |
| 49 | @ | Text |

## Impact

- Cells with missing format IDs show raw numbers
- Accounting formats may not align properly
- Time duration formats won't work

## Proposed Fix

1. Complete the `get_builtin_format` function with all IDs 0-49
2. Add locale-aware handling for IDs 27-36
3. Implement special formatting for accounting (41-44)

## References

- ECMA-376 Part 1, Section 18.8.30
- [MS-OI29500] Section 2.1.622

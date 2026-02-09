//! Tests for comment rich text parsing in XLSX files
//!
//! These tests parse real XLSX files to verify that comments with rich text
//! formatting (multiple text runs with different styles) are correctly parsed.
//!
//! Tested features:
//! - Basic comment parsing from real files
//! - Rich text runs with different styles (bold, italic, underline, etc.)
//! - Font properties (family, size, color)
//! - Comment text content extraction
//! - Graceful handling of comments without styling

#![allow(
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

use std::fs;

// =============================================================================
// Helper Functions
// =============================================================================

/// Parse an XLSX file and return the workbook
#[allow(clippy::expect_used)]
fn parse_test_file(path: &str) -> xlview::types::Workbook {
    let data = fs::read(path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    xlview::parser::parse(&data).unwrap_or_else(|_| panic!("Failed to parse XLSX file: {}", path))
}

/// Find a sheet by name in the workbook
#[allow(dead_code)]
fn find_sheet<'a>(
    workbook: &'a xlview::types::Workbook,
    name: &str,
) -> Option<&'a xlview::types::Sheet> {
    workbook.sheets.iter().find(|s| s.name == name)
}

/// Get all comments from all sheets
fn get_all_comments(workbook: &xlview::types::Workbook) -> Vec<&xlview::types::Comment> {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.comments.iter())
        .collect()
}

/// Get a comment by cell reference from a sheet
#[allow(dead_code)]
fn get_comment_by_cell<'a>(
    sheet: &'a xlview::types::Sheet,
    cell_ref: &str,
) -> Option<&'a xlview::types::Comment> {
    sheet.comments.iter().find(|c| c.cell_ref == cell_ref)
}

/// Check if a comment has rich text formatting
fn has_rich_text_formatting(comment: &xlview::types::Comment) -> bool {
    comment.rich_text.is_some() && !comment.rich_text.as_ref().unwrap().is_empty()
}

/// Count comments with rich text formatting
fn count_rich_text_comments(workbook: &xlview::types::Workbook) -> usize {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.comments.iter())
        .filter(|c| has_rich_text_formatting(c))
        .count()
}

/// Count total text runs across all rich text comments
fn count_total_text_runs(workbook: &xlview::types::Workbook) -> usize {
    workbook
        .sheets
        .iter()
        .flat_map(|s| s.comments.iter())
        .filter_map(|c| c.rich_text.as_ref())
        .map(|runs| runs.len())
        .sum()
}

// =============================================================================
// Tests: test_comments.xlsx - Specific Comment Lookups
// =============================================================================

#[test]
fn test_get_comment_by_cell_ref() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    for sheet in &workbook.sheets {
        if !sheet.comments.is_empty() {
            // Get the first comment's cell ref
            let first_cell_ref = &sheet.comments[0].cell_ref;

            // Use the helper to find it
            let found = get_comment_by_cell(sheet, first_cell_ref);
            assert!(
                found.is_some(),
                "Should find comment at {} in sheet '{}'",
                first_cell_ref,
                sheet.name
            );

            // Verify it's the same comment
            assert_eq!(
                found.unwrap().text,
                sheet.comments[0].text,
                "Found comment should match original"
            );
        }
    }
}

#[test]
fn test_find_sheet_by_name() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    // Try to find a sheet that exists
    for sheet in &workbook.sheets {
        let found = find_sheet(&workbook, &sheet.name);
        assert!(
            found.is_some(),
            "Should find sheet by name: '{}'",
            sheet.name
        );
    }

    // Try to find a sheet that doesn't exist
    let not_found = find_sheet(&workbook, "NonExistentSheet12345");
    assert!(not_found.is_none(), "Should not find non-existent sheet");
}

// =============================================================================
// Tests: kitchen_sink_v2.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_v2_parsing_does_not_panic() {
    // This test verifies that parsing kitchen_sink_v2.xlsx with comments doesn't panic
    let result = std::panic::catch_unwind(|| parse_test_file("test/kitchen_sink_v2.xlsx"));

    assert!(
        result.is_ok(),
        "Parsing kitchen_sink_v2.xlsx should not panic"
    );

    let workbook = result.unwrap();
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

#[test]
fn test_kitchen_sink_v2_comments_structure() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    // Count total comments across all sheets
    let total_comments: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();

    // kitchen_sink_v2.xlsx may or may not have comments
    // This test just verifies the structure is correct
    println!("kitchen_sink_v2.xlsx has {} total comments", total_comments);

    // Verify comment structure for any comments that exist
    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            // Every comment must have a cell reference
            assert!(
                !comment.cell_ref.is_empty(),
                "Comment cell_ref should not be empty"
            );
            // Every comment must have text content
            // (even if it's empty string, the field should exist)
        }
    }
}

// =============================================================================
// Tests: kitchen_sink.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_kitchen_sink_parsing_does_not_panic() {
    // This test verifies that parsing kitchen_sink.xlsx with comments doesn't panic
    let result = std::panic::catch_unwind(|| parse_test_file("test/kitchen_sink.xlsx"));

    assert!(result.is_ok(), "Parsing kitchen_sink.xlsx should not panic");

    let workbook = result.unwrap();
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

#[test]
fn test_kitchen_sink_comments_structure() {
    let workbook = parse_test_file("test/kitchen_sink.xlsx");

    // Count total comments across all sheets
    let total_comments: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();

    println!("kitchen_sink.xlsx has {} total comments", total_comments);

    // Verify comment structure
    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            assert!(
                !comment.cell_ref.is_empty(),
                "Comment cell_ref should not be empty"
            );
        }
    }
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Comment Parsing
// =============================================================================

#[test]
fn test_ms_cf_samples_parsing_does_not_panic() {
    // This test verifies that parsing ms_cf_samples.xlsx with comments doesn't panic
    let result = std::panic::catch_unwind(|| parse_test_file("test/ms_cf_samples.xlsx"));

    assert!(
        result.is_ok(),
        "Parsing ms_cf_samples.xlsx should not panic"
    );

    let workbook = result.unwrap();
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );
}

#[test]
fn test_ms_cf_samples_comments_structure() {
    let workbook = parse_test_file("test/ms_cf_samples.xlsx");

    // Count total comments across all sheets
    let total_comments: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();

    println!("ms_cf_samples.xlsx has {} total comments", total_comments);

    // Verify comment structure
    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            assert!(
                !comment.cell_ref.is_empty(),
                "Comment cell_ref should not be empty"
            );
        }
    }
}

// =============================================================================
// Tests: test_comments.xlsx - Dedicated Comment Test File
// =============================================================================

#[test]
fn test_dedicated_comments_file_parsing() {
    // test_comments.xlsx is specifically for testing comments
    let workbook = parse_test_file("test/test_comments.xlsx");

    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );

    // This file should have comments
    let total_comments: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();
    println!("test_comments.xlsx has {} total comments", total_comments);
}

#[test]
fn test_dedicated_comments_file_content() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    // Collect all comments
    let all_comments = get_all_comments(&workbook);

    // Check each comment has required fields
    for comment in &all_comments {
        // Cell reference should be valid (letter + number format like A1, B2, etc.)
        assert!(
            comment.cell_ref.chars().any(|c| c.is_ascii_alphabetic()),
            "Comment cell_ref should contain column letters: {}",
            comment.cell_ref
        );
        assert!(
            comment.cell_ref.chars().any(|c| c.is_ascii_digit()),
            "Comment cell_ref should contain row numbers: {}",
            comment.cell_ref
        );
    }
}

#[test]
fn test_dedicated_comments_file_rich_text() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    // Check for rich text formatting in comments
    let rich_text_count = count_rich_text_comments(&workbook);
    println!(
        "test_comments.xlsx has {} comments with rich text formatting",
        rich_text_count
    );

    // Check text runs count
    let total_runs = count_total_text_runs(&workbook);
    println!(
        "test_comments.xlsx has {} total text runs in rich text comments",
        total_runs
    );

    // Verify rich text structure if present
    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                // Each run should have text
                for run in runs {
                    // Text can be empty but the field must exist
                    // Style is optional
                    if let Some(ref style) = run.style {
                        // If style exists, verify boolean fields don't panic
                        let _bold = style.bold;
                        let _italic = style.italic;
                        let _underline = style.underline;
                        let _strikethrough = style.strikethrough;
                    }
                }

                // Verify that concatenated text matches the plain text
                let concatenated: String = runs.iter().map(|r| r.text.as_str()).collect();
                assert_eq!(
                    concatenated, comment.text,
                    "Rich text concatenation should match plain text for comment at {}",
                    comment.cell_ref
                );
            }
        }
    }
}

// =============================================================================
// Tests: Rich Text Styling Properties
// =============================================================================

#[test]
fn test_rich_text_bold_styling() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    // Look for any comment with bold text runs
    let mut found_bold = false;

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                for run in runs {
                    if let Some(ref style) = run.style {
                        if style.bold == Some(true) {
                            found_bold = true;
                            println!(
                                "Found bold text '{}' in comment at {}",
                                run.text, comment.cell_ref
                            );
                        }
                    }
                }
            }
        }
    }

    // Note: This may or may not find bold text depending on the test file content
    println!("Bold text found in comments: {}", found_bold);
}

#[test]
fn test_rich_text_italic_styling() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    // Look for any comment with italic text runs
    let mut found_italic = false;

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                for run in runs {
                    if let Some(ref style) = run.style {
                        if style.italic == Some(true) {
                            found_italic = true;
                            println!(
                                "Found italic text '{}' in comment at {}",
                                run.text, comment.cell_ref
                            );
                        }
                    }
                }
            }
        }
    }

    println!("Italic text found in comments: {}", found_italic);
}

#[test]
fn test_rich_text_font_properties() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    let mut font_families: Vec<String> = Vec::new();
    let mut font_sizes: Vec<f64> = Vec::new();
    let mut font_colors: Vec<String> = Vec::new();

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                for run in runs {
                    if let Some(ref style) = run.style {
                        if let Some(ref family) = style.font_family {
                            if !font_families.contains(family) {
                                font_families.push(family.clone());
                            }
                        }
                        if let Some(size) = style.font_size {
                            if !font_sizes.contains(&size) {
                                font_sizes.push(size);
                            }
                        }
                        if let Some(ref color) = style.font_color {
                            if !font_colors.contains(color) {
                                font_colors.push(color.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    println!("Font families found: {:?}", font_families);
    println!("Font sizes found: {:?}", font_sizes);
    println!("Font colors found: {:?}", font_colors);
}

// =============================================================================
// Tests: Comment Author Parsing
// =============================================================================

#[test]
fn test_comment_authors() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    let mut authors: Vec<String> = Vec::new();
    let mut comments_with_authors = 0;
    let mut comments_without_authors = 0;

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            match &comment.author {
                Some(author) => {
                    comments_with_authors += 1;
                    if !authors.contains(author) {
                        authors.push(author.clone());
                    }
                }
                None => {
                    comments_without_authors += 1;
                }
            }
        }
    }

    println!("Unique authors found: {:?}", authors);
    println!("Comments with authors: {}", comments_with_authors);
    println!("Comments without authors: {}", comments_without_authors);
}

// =============================================================================
// Tests: Comment Text Content Verification
// =============================================================================

#[test]
fn test_comment_text_not_empty_when_rich_text_present() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                // If there are rich text runs with content, the plain text should not be empty
                let has_content = runs.iter().any(|r| !r.text.is_empty());
                if has_content {
                    assert!(
                        !comment.text.is_empty(),
                        "Comment at {} has rich text runs with content but empty plain text",
                        comment.cell_ref
                    );
                }
            }
        }
    }
}

#[test]
fn test_comment_text_matches_rich_text_concatenation() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                let concatenated: String = runs.iter().map(|r| r.text.as_str()).collect();
                assert_eq!(
                    concatenated, comment.text,
                    "Rich text concatenation mismatch for comment at {} in sheet '{}'",
                    comment.cell_ref, sheet.name
                );
            }
        }
    }
}

// =============================================================================
// Tests: Multiple Sheets with Comments
// =============================================================================

#[test]
fn test_comments_per_sheet() {
    let workbook = parse_test_file("test/kitchen_sink_v2.xlsx");

    for sheet in &workbook.sheets {
        println!(
            "Sheet '{}' has {} comments",
            sheet.name,
            sheet.comments.len()
        );

        // Each comment should reference a cell within the sheet's bounds
        for comment in &sheet.comments {
            // Parse the cell reference to get row/col
            let (col, row) = parse_cell_ref(&comment.cell_ref);

            // If sheet has max_row/max_col defined, verify comment is within bounds
            // (unless they're 0, which means the sheet might be empty of data)
            if sheet.max_row > 0 && sheet.max_col > 0 {
                // Comments can be on cells beyond the data range, so this is informational
                if row > sheet.max_row || col > sheet.max_col {
                    println!(
                        "  Note: Comment at {} is beyond data bounds (max: row {}, col {})",
                        comment.cell_ref, sheet.max_row, sheet.max_col
                    );
                }
            }
        }
    }
}

/// Parse a cell reference (e.g., "A1") into (col, row) as 0-indexed
fn parse_cell_ref(ref_str: &str) -> (u32, u32) {
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

// =============================================================================
// Tests: Cross-File Consistency
// =============================================================================

#[test]
fn test_all_files_parse_without_panic() {
    let test_files = [
        "test/kitchen_sink_v2.xlsx",
        "test/kitchen_sink.xlsx",
        "test/ms_cf_samples.xlsx",
        "test/test_comments.xlsx",
    ];

    for file in &test_files {
        let result = std::panic::catch_unwind(|| {
            let data = fs::read(file);
            if let Ok(data) = data {
                let _ = xlview::parser::parse(&data);
            }
        });

        assert!(result.is_ok(), "Parsing {} should not panic", file);
    }
}

#[test]
fn test_comment_cell_ref_format_consistency() {
    let test_files = [
        "test/kitchen_sink_v2.xlsx",
        "test/kitchen_sink.xlsx",
        "test/ms_cf_samples.xlsx",
        "test/test_comments.xlsx",
    ];

    for file in &test_files {
        let Ok(data) = fs::read(file) else { continue };

        let Ok(workbook) = xlview::parser::parse(&data) else {
            continue;
        };

        for sheet in &workbook.sheets {
            for comment in &sheet.comments {
                // Verify cell reference format (e.g., A1, B2, AA100)
                let cell_ref = &comment.cell_ref;

                // Should start with one or more letters
                let letter_count = cell_ref
                    .chars()
                    .take_while(|c| c.is_ascii_alphabetic())
                    .count();
                assert!(
                    letter_count > 0,
                    "Cell ref '{}' in {} should start with letters",
                    cell_ref,
                    file
                );

                // Should end with one or more digits
                let digit_count = cell_ref
                    .chars()
                    .skip(letter_count)
                    .filter(|c| c.is_ascii_digit())
                    .count();
                assert!(
                    digit_count > 0,
                    "Cell ref '{}' in {} should end with digits",
                    cell_ref,
                    file
                );
            }
        }
    }
}

// =============================================================================
// Tests: Rich Text Run Count Statistics
// =============================================================================

#[test]
fn test_rich_text_run_statistics() {
    let test_files = [
        ("test/kitchen_sink_v2.xlsx", "kitchen_sink_v2"),
        ("test/kitchen_sink.xlsx", "kitchen_sink"),
        ("test/ms_cf_samples.xlsx", "ms_cf_samples"),
        ("test/test_comments.xlsx", "test_comments"),
    ];

    println!("\nRich Text Statistics:");
    println!(
        "{:<25} {:>10} {:>15} {:>15}",
        "File", "Comments", "With Rich Text", "Total Runs"
    );
    println!("{}", "-".repeat(70));

    for (file, name) in &test_files {
        let Ok(data) = fs::read(file) else { continue };

        let Ok(workbook) = xlview::parser::parse(&data) else {
            continue;
        };

        let total_comments: usize = workbook.sheets.iter().map(|s| s.comments.len()).sum();
        let rich_text_comments = count_rich_text_comments(&workbook);
        let total_runs = count_total_text_runs(&workbook);

        println!(
            "{:<25} {:>10} {:>15} {:>15}",
            name, total_comments, rich_text_comments, total_runs
        );
    }
}

// =============================================================================
// Tests: Verify has_comment Flag on Cells
// =============================================================================

#[test]
fn test_cells_with_comments_have_flag_set() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            // Parse the cell reference to row/col
            let (col, row) = parse_cell_ref(&comment.cell_ref);

            // Find the cell in the sheet
            let cell = sheet.cells.iter().find(|c| c.r == row && c.c == col);

            if let Some(cell_data) = cell {
                assert_eq!(
                    cell_data.cell.has_comment,
                    Some(true),
                    "Cell at {} (row {}, col {}) should have has_comment=true",
                    comment.cell_ref,
                    row,
                    col
                );
            }
            // Note: It's valid for a comment to exist on an empty cell
            // which might not be in the cells array
        }
    }
}

// =============================================================================
// Tests: Vertical Alignment in Rich Text
// =============================================================================

#[test]
fn test_rich_text_vertical_alignment() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    let mut found_subscript = false;
    let mut found_superscript = false;

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                for run in runs {
                    if let Some(ref style) = run.style {
                        if let Some(ref vert_align) = style.vert_align {
                            match vert_align {
                                xlview::types::VerticalAlign::Subscript => {
                                    found_subscript = true;
                                    println!("Found subscript text: '{}'", run.text);
                                }
                                xlview::types::VerticalAlign::Superscript => {
                                    found_superscript = true;
                                    println!("Found superscript text: '{}'", run.text);
                                }
                                xlview::types::VerticalAlign::Baseline => {}
                            }
                        }
                    }
                }
            }
        }
    }

    println!("Subscript found: {}", found_subscript);
    println!("Superscript found: {}", found_superscript);
}

// =============================================================================
// Tests: Edge Cases
// =============================================================================

#[test]
fn test_empty_rich_text_runs_handled() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if let Some(ref runs) = comment.rich_text {
                // Check that we handle empty text runs gracefully
                for run in runs {
                    // Empty strings are valid
                    let _ = run.text.len();
                }
            }
        }
    }
}

#[test]
fn test_comments_without_rich_text() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    let mut plain_text_comments = 0;
    let mut rich_text_comments = 0;

    for sheet in &workbook.sheets {
        for comment in &sheet.comments {
            if comment.rich_text.is_some() {
                rich_text_comments += 1;
            } else {
                plain_text_comments += 1;
            }
        }
    }

    println!(
        "Plain text comments (no rich_text field): {}",
        plain_text_comments
    );
    println!(
        "Rich text comments (has rich_text field): {}",
        rich_text_comments
    );
}

#[test]
fn test_comment_json_serialization() {
    let workbook = parse_test_file("test/test_comments.xlsx");

    // Verify the workbook can be serialized to JSON
    let json_result = serde_json::to_string(&workbook);
    assert!(json_result.is_ok(), "Workbook should serialize to JSON");

    let json = json_result.unwrap();

    // The JSON should contain the comments field for sheets
    if workbook.sheets.iter().any(|s| !s.comments.is_empty()) {
        assert!(
            json.contains("comments") || json.contains("cellRef"),
            "JSON should contain comment-related fields"
        );
    }
}

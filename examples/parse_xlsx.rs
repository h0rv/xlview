//! Example: Parse an XLSX file and print basic information
//!
//! Run with: cargo run --example parse_xlsx -- path/to/file.xlsx

#![allow(clippy::expect_used, clippy::indexing_slicing)]

use std::env;
use std::fs;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: {} <xlsx-file>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let data = fs::read(path).expect("Failed to read file");

    match xlview::parser::parse(&data) {
        Ok(workbook) => {
            println!("Workbook: {}", path);
            println!("Sheets: {}", workbook.sheets.len());

            for (i, sheet) in workbook.sheets.iter().enumerate() {
                println!("\n  Sheet {}: \"{}\"", i + 1, sheet.name);
                println!("    Cells: {}", sheet.cells.len());
                println!("    Merges: {}", sheet.merges.len());

                if sheet.max_row > 0 || sheet.max_col > 0 {
                    println!(
                        "    Dimensions: {}x{} (rows x cols)",
                        sheet.max_row + 1,
                        sheet.max_col + 1
                    );
                }

                // Print first few cell values
                let preview_cells: Vec<_> = sheet
                    .cells
                    .iter()
                    .filter(|c| c.cell.v.is_some())
                    .take(5)
                    .collect();

                if !preview_cells.is_empty() {
                    println!("    Sample values:");
                    for cell in preview_cells {
                        if let Some(ref v) = cell.cell.v {
                            let truncated = if v.len() > 40 {
                                format!("{}...", &v[..40])
                            } else {
                                v.clone()
                            };
                            println!("      ({}, {}): {}", cell.r, cell.c, truncated);
                        }
                    }
                }
            }
        }
        Err(e) => {
            eprintln!("Failed to parse XLSX: {}", e);
            std::process::exit(1);
        }
    }
}

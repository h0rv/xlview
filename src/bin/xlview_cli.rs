//! CLI tool for xlview - parses XLSX files and outputs JSON
//!
//! Usage:
//!   xlview_cli <input.xlsx>              # Output JSON to stdout
//!   xlview_cli <input.xlsx> -o out.json  # Output JSON to file

#![allow(clippy::exit)]
#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::indexing_slicing)]

use std::env;
use std::fs;
use std::io::{self, Write};
use xlview::parser::parse;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage: xlview_cli <input.xlsx> [-o output.json]");
        std::process::exit(1);
    }

    let input_path = &args[1];
    let output_path = if args.len() > 3 && args[2] == "-o" {
        Some(&args[3])
    } else {
        None
    };

    // Read input file
    let data = match fs::read(input_path) {
        Ok(d) => d,
        Err(e) => {
            eprintln!("Error reading {}: {}", input_path, e);
            std::process::exit(1);
        }
    };

    // Parse XLSX
    let workbook = match parse(&data) {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("Error parsing XLSX: {}", e);
            std::process::exit(1);
        }
    };

    // Serialize to JSON
    let json = match serde_json::to_string_pretty(&workbook) {
        Ok(j) => j,
        Err(e) => {
            eprintln!("Error serializing JSON: {}", e);
            std::process::exit(1);
        }
    };

    // Output
    match output_path {
        Some(path) => {
            if let Err(e) = fs::write(path, &json) {
                eprintln!("Error writing {}: {}", path, e);
                std::process::exit(1);
            }
            eprintln!("Written: {}", path);
        }
        None => {
            io::stdout().write_all(json.as_bytes()).unwrap();
            println!();
        }
    }
}

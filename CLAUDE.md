# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

xlview is a WASM-based Excel (XLSX) viewer with Canvas 2D rendering. Parses XLSX in Rust, compiles to WebAssembly, renders in the browser. View-only (no editing). Zero runtime dependencies.

## Build & Development Commands

```bash
# Build WASM (dev)
wasm-pack build --target web --dev

# Build WASM (release)
wasm-pack build --target web --release

# Run Rust unit tests
cargo test --lib

# Run all tests (CI uses this)
cargo test --all-features

# Run a single test
cargo test --lib test_name

# Run tests in a specific file
cargo test --lib --test file_name

# Clippy (matches CI)
cargo clippy --all-targets --all-features -- -D warnings

# Clippy lib only (faster, used in justfile)
cargo clippy --lib -- -D warnings

# Format check
cargo fmt --all -- --check

# Format
cargo fmt

# Full quality check (fmt + lint + test)
just check

# Serve demo at localhost:8080
python3 -m http.server 8080

# Playwright visual regression tests (needs WASM built + server running)
npm test

# Update visual snapshots
npm run test:update
```

## Strict Lint Rules

The codebase enforces zero unsafe code and strict error handling via `Cargo.toml` lints:
- **Forbidden**: `unsafe` code
- **Denied**: `.unwrap()`, `.expect()`, `panic!()`, array indexing (`[]`), `todo!()`, `unimplemented!()`, `dbg!()`, lossy casts
- **Required**: Use `.get()` for indexing, `?` for error propagation, proper `Result`/`Option` handling

CI runs with `RUSTFLAGS=-Dwarnings` — all warnings are errors.

## Architecture

### Parsing Pipeline (`src/parser.rs` orchestrates)
XLSX files are ZIP archives of XML. The parser extracts and processes in order:
1. Shared strings table
2. Theme/color definitions (`theme_parser.rs`, `color.rs`)
3. Styles: fonts, fills, borders, number formats (`styles.rs`, `numfmt.rs`)
4. Worksheets with cell data, merges, dimensions
5. Relationships: charts, images, comments, hyperlinks

Core types are in `src/types.rs`. Each feature area has its own module (charts, conditional formatting, data validation, sparklines, drawings, etc).

### Rendering Pipeline
- `src/layout/` — Pre-computes cell positions (`sheet_layout.rs`) and manages viewport/scroll state (`viewport.rs`)
- `src/render/` — Backend-agnostic rendering trait (`backend.rs`), with Canvas 2D implementation in `render/canvas/`
  - `canvas/renderer.rs` — Main cell/text/border drawing
  - `canvas/headers.rs` — Row/column headers
  - `canvas/frozen.rs` — Frozen pane support
  - `canvas/indicators.rs` — Comment markers, etc.
- `src/render/blit.rs` — Tile-based off-screen caching (512px tiles)
- `src/render/selection.rs` — Selection overlay

### Viewer (`src/viewer.rs`)
`XlView` is the main WASM-exported struct. It owns the parsed workbook, layout engine, and renderer. Handles scroll, click, keyboard events, sheet switching, and coordinates the full render cycle.

### WASM Exports (`src/lib.rs`)
- `XlView` — Full interactive viewer (canvas-based)
- `parse_xlsx()` / `parse_xlsx_to_js()` — Parse-only APIs returning JSON or JsValue
- `version()` — Library version

### CLI Binary
`src/bin/xlview_cli.rs` — Converts XLSX to JSON for debugging/testing.

## Test Structure

- **Rust unit tests** (`tests/*.rs`, ~60 files): Cover all parsing modules. Run with `cargo test --lib`.
- **Playwright visual tests** (`tests/visual.spec.ts`): Screenshot comparison against baselines in `tests/snapshots/`. Config in `playwright.config.ts`.
- **Test data**: `test/` directory has xlsx fixtures (`minimal.xlsx`, `styled.xlsx`, `kitchen_sink.xlsx`, `large_5000x20.xlsx`, etc).

## CI

GitHub Actions (`.github/workflows/ci.yml`):
- Tests on stable + nightly Rust
- Clippy with `-D warnings` on all targets
- Format check with `rustfmt`

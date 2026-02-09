# xlview

[![CI](https://github.com/h0rv/xlview/actions/workflows/ci.yml/badge.svg)](https://github.com/h0rv/xlview/actions/workflows/ci.yml)
[![Crates.io](https://img.shields.io/crates/v/xlview.svg)](https://crates.io/crates/xlview)
[![npm](https://img.shields.io/npm/v/xlview.svg)](https://www.npmjs.com/package/xlview)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

XLSX viewer for the web. Parses and renders Excel files in the browser using WebAssembly and Canvas 2D. View-only. Zero runtime dependencies.

## Features

- **Styling** — fonts, colors, borders, fills, gradients, conditional formatting
- **Charts** — bar, line, pie, scatter, area, radar, and more
- **Rich content** — images, shapes, comments, hyperlinks, sparklines
- **Layout** — frozen panes, merged cells, hidden rows/columns, grouping
- **Performance** — 100k+ cells at 120fps, tile-cached Canvas 2D rendering
- **Interaction** — native scroll, cell selection, keyboard navigation, copy-to-clipboard

## Install

```bash
npm install xlview
```

Or via CDN:

```html
<script type="module">
  import init, { XlView } from 'https://unpkg.com/xlview/xlview.js';
</script>
```

Or as a Rust crate (parsing only, no rendering):

```bash
cargo add xlview
```

## Quick Start

Drop this into an HTML file:

```html
<div id="container" style="width: 100%; height: 600px; position: relative;">
  <canvas id="base"></canvas>
  <canvas id="overlay"></canvas>
</div>

<script type="module">
  import init, { XlView } from 'xlview';
  await init();

  const base = document.getElementById('base');
  const overlay = document.getElementById('overlay');
  const container = document.getElementById('container');
  const dpr = window.devicePixelRatio || 1;

  // Size canvases to container
  for (const c of [base, overlay]) {
    c.width = container.clientWidth * dpr;
    c.height = container.clientHeight * dpr;
    c.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%';
  }
  overlay.style.pointerEvents = 'none';

  const viewer = XlView.new_with_overlay(base, overlay, dpr);

  // Load a file
  const res = await fetch('spreadsheet.xlsx');
  viewer.load(new Uint8Array(await res.arrayBuffer()));
  viewer.render();

  // Handle resize
  new ResizeObserver(() => {
    const w = container.clientWidth * dpr;
    const h = container.clientHeight * dpr;
    base.width = w; base.height = h;
    overlay.width = w; overlay.height = h;
    viewer.resize(w, h, dpr);
  }).observe(container);
</script>
```

Scroll, click, keyboard, sheet tabs, and selection all work automatically.

## React

```tsx
import { useEffect, useRef } from 'react';
import init, { XlView } from 'xlview';

export function ExcelViewer({ url }: { url: string }) {
  const baseRef = useRef<HTMLCanvasElement>(null);
  const overlayRef = useRef<HTMLCanvasElement>(null);
  const viewerRef = useRef<XlView | null>(null);

  useEffect(() => {
    let active = true;
    (async () => {
      await init();
      if (!active) return;

      const base = baseRef.current!;
      const overlay = overlayRef.current!;
      const dpr = window.devicePixelRatio || 1;
      const w = base.clientWidth * dpr;
      const h = base.clientHeight * dpr;
      base.width = w; base.height = h;
      overlay.width = w; overlay.height = h;

      viewerRef.current = XlView.new_with_overlay(base, overlay, dpr);

      const res = await fetch(url);
      viewerRef.current.load(new Uint8Array(await res.arrayBuffer()));
      viewerRef.current.render();
    })();
    return () => { active = false; viewerRef.current?.free(); };
  }, [url]);

  const style = { position: 'absolute' as const, top: 0, left: 0, width: '100%', height: '100%' };

  return (
    <div style={{ position: 'relative', width: '100%', height: 600 }}>
      <canvas ref={baseRef} style={style} />
      <canvas ref={overlayRef} style={{ ...style, pointerEvents: 'none' }} />
    </div>
  );
}
```

## Parse Only (No Rendering)

Extract cell data without a canvas:

```javascript
import init, { parse_xlsx_to_js } from 'xlview';
await init();

const data = await fetch('file.xlsx').then(r => r.arrayBuffer());
const workbook = parse_xlsx_to_js(new Uint8Array(data));

for (const sheet of workbook.sheets) {
  console.log(sheet.name);
  for (const cell of sheet.cells) {
    console.log(`  ${cell.r},${cell.c}: ${cell.cell.v}`);
  }
}
```

### Rust

```rust
use xlview::parser;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let data = std::fs::read("spreadsheet.xlsx")?;
    let workbook = parser::parse(&data)?;

    for sheet in &workbook.sheets {
        println!("{}", sheet.name);
        for cell in &sheet.cells {
            if let Some(v) = &cell.cell.v {
                println!("  ({},{}): {v}", cell.r, cell.c);
            }
        }
    }
    Ok(())
}
```

## API

### XlView (WASM)

| Method | Description |
|--------|-------------|
| `new(canvas, dpr)` | Create viewer (single canvas, no overlay) |
| `new_with_overlay(base, overlay, dpr)` | Create viewer with selection overlay (recommended) |
| `load(data)` | Load XLSX from `Uint8Array` |
| `render()` | Render current view |
| `resize(w, h, dpr)` | Handle container resize |
| `set_active_sheet(index)` | Switch to sheet by index |
| `sheet_count()` | Number of sheets |
| `sheet_name(index)` | Get sheet name |
| `active_sheet()` | Current sheet index |
| `get_selection()` | Get selected cell range `[r1, c1, r2, c2]` |
| `set_headers_visible(bool)` | Toggle row/column headers |
| `free()` | Release WASM memory |

### Standalone Functions

| Function | Description |
|----------|-------------|
| `parse_xlsx(data)` | Parse XLSX, return JSON string |
| `parse_xlsx_to_js(data)` | Parse XLSX, return JS object |
| `version()` | Library version |

## Browser Support

Chrome 57+, Firefox 52+, Safari 11+, Edge 79+

## Build from Source

```bash
cargo install wasm-pack
wasm-pack build --target web --release
```

## License

[MIT](LICENSE)

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
<script type="module" src="https://unpkg.com/xlview"></script>
```

Or as a Rust crate (parsing only, no rendering):

```bash
cargo add xlview
```

## Quick Start

### Drop-in (1 line)

```html
<script type="module" src="https://unpkg.com/xlview"></script>
<xl-view src="spreadsheet.xlsx" style="width:100%;height:600px"></xl-view>
```

The `<xl-view>` custom element handles canvas setup, resize, DPR, and rendering automatically. Scroll, click, keyboard, sheet tabs, and selection all work out of the box.

### Programmatic

```js
import { mount } from 'xlview';

const viewer = await mount(document.getElementById('container'));
const res = await fetch('spreadsheet.xlsx');
viewer.load(new Uint8Array(await res.arrayBuffer()));
```

`mount()` creates canvases, wires up resize handling, and returns a controller with `load()`, `destroy()`, and the underlying `viewer` instance.

### React

```tsx
import { useEffect, useRef } from 'react';
import { mount, type MountedViewer } from 'xlview';

export function ExcelViewer({ url }: { url: string }) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let mounted: MountedViewer | null = null;
    (async () => {
      mounted = await mount(containerRef.current!);
      const res = await fetch(url);
      mounted.load(new Uint8Array(await res.arrayBuffer()));
    })();
    return () => { mounted?.destroy(); };
  }, [url]);

  return <div ref={containerRef} style={{ width: '100%', height: 600 }} />;
}
```

### Full Control

For custom canvas pipelines, use the WASM API directly:

```js
import init, { XlView } from 'xlview/core';
await init();

const viewer = XlView.newWithOverlay(baseCanvas, overlayCanvas, devicePixelRatio);
viewer.load(data);
viewer.render();
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

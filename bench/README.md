# Browser Benchmarks (xlview)

This folder contains a browser-based benchmark harness for xlview. It measures:
- **Parse** (XLSX → JS workbook)
- **Load** (parse + layout)
- **Render** (single render of current viewport)
- **End-to-end** (load + render)
- **Scroll** (N scroll steps + render)
- **Internal metrics (xlview only)**: parse/layout/total inside WASM
- **Internal render metrics (xlview only)**: prep/draw/total + visible cell count

Designed to add new providers later (see `bench/providers/`).

Current providers:
- `xlview` (WASM)
- `sheetjs` (SheetJS CE + DOM table render baseline)

## Setup

1) Build the WASM bundle (release recommended):
```bash
wasm-pack build --target web --release
```

2) (Optional) Regenerate test XLSX files:
```bash
uv run scripts/generate_test_xlsx.py
```

3) Install bench dependencies with Bun:
```bash
cd bench
bun install
bunx playwright install chromium
```

## Run (headless)
```bash
cd bench
bun run bench
```

## Summarize results
```bash
cd bench
bun run summary -- --compare --last 5
```

## Run (UI)
```bash
cd bench
bun run bench:ui
```

## Options
Set env vars to control scope:
```bash
BENCH_ITERATIONS=5 BENCH_WARMUP=1 BENCH_DATASET=large_5000x20 BENCH_PROVIDER=xlview BENCH_SCROLL_STEPS=10 BENCH_SCROLL_STEP_PX=200 bun run bench
```

- `BENCH_DATASET`: dataset `id` from `bench/manifest.json` (or `all`)
- `BENCH_PROVIDER`: provider key from `bench/providers/index.js`
- `BENCH_SCROLL_STEPS`: number of scroll steps per iteration
- `BENCH_SCROLL_STEP_PX`: pixels per scroll step

## Results
Results are written to:
```
bench/results/bench-<provider>-<timestamp>.json
```

## Summarize results
```bash
cd bench
bun run summary -- --compare --last 5
```

### CSV export
```bash
cd bench
bun run summary -- --csv
```

Custom path:
```bash
cd bench
bun run summary -- --csv --csv-out ./results/bench-summary.csv
```

## Adding a new provider
1) Create `bench/providers/<name>.js` that returns an object with `parse`, `load`, and `render`.
2) Register it in `bench/providers/index.js`.
3) Add any required assets or wiring under `bench/`.

See `bench/providers/xlview.js` as the reference implementation.

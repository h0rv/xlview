#!/usr/bin/env bun
/**
 * Unified E2E test suite for xlview.
 *
 * Tests content correctness, feature regressions, and performance.
 * Runs headless via Playwright — no screenshots, no flaky visual diffs.
 * Uses WASM debug APIs (get_scroll_debug, get_header_config, render_with_metrics)
 * and canvas pixel sampling for ground-truth validation.
 *
 * Run: bun tests/e2e/test_suite.ts
 * Or:  just e2e
 */

import { chromium } from "playwright";
import { createServer, type IncomingMessage, type ServerResponse } from "http";
import { readFileSync, existsSync } from "fs";
import { join, extname } from "path";
import { fileURLToPath } from "url";
import type { Server } from "http";
import type { Browser, Page } from "playwright";

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const PROJECT_ROOT = join(__dirname, "../..");
const PORT = 8799;

const MIME_TYPES: Record<string, string> = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".mjs": "application/javascript",
  ".wasm": "application/wasm",
  ".json": "application/json",
  ".xlsx":
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".css": "text/css",
  ".png": "image/png",
};

// Performance thresholds (ms)
const PERF = {
  load_p90: 500, // file load + first render
  render_p90: 30, // single render() call
  scroll_p90: 20, // scroll + render
};

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ScrollDebug {
  viewport_scroll_x: number;
  viewport_scroll_y: number;
  viewport_width: number;
  viewport_height: number;
  container_scroll_left: number;
  container_scroll_top: number;
  frozen_cols_width: number;
  frozen_rows_height: number;
  visible_start_row: number;
  visible_end_row: number;
  show_headers: boolean;
}

interface HeaderConfig {
  row_header_width: number;
  col_header_height: number;
  visible: boolean;
}

interface RenderMetrics {
  prep_ms: number;
  draw_ms: number;
  total_ms: number;
  visible_cells: number;
  skipped: boolean;
}

interface Pixel {
  r: number;
  g: number;
  b: number;
  a: number;
}

// ---------------------------------------------------------------------------
// Test harness
// ---------------------------------------------------------------------------

let passed = 0;
let failed = 0;
let skipped = 0;
const failures: string[] = [];

function ok(name: string, detail = ""): void {
  passed++;
  const d = detail ? ` (${detail})` : "";
  console.log(`  \x1b[32m✓\x1b[0m ${name}${d}`);
}

function fail(name: string, detail = ""): void {
  failed++;
  const d = detail ? ` (${detail})` : "";
  console.log(`  \x1b[31m✗\x1b[0m ${name}${d}`);
  failures.push(`${name}: ${detail}`);
}

function skip(name: string, reason = ""): void {
  skipped++;
  console.log(`  \x1b[33m-\x1b[0m ${name} [skipped: ${reason}]`);
}

function assert(cond: boolean, name: string, detail = ""): void {
  if (cond) ok(name, detail);
  else fail(name, detail);
}

function approxColor(
  px: Pixel,
  target: [number, number, number],
  tol = 20,
): boolean {
  return (
    Math.abs(px.r - target[0]) <= tol &&
    Math.abs(px.g - target[1]) <= tol &&
    Math.abs(px.b - target[2]) <= tol
  );
}

// ---------------------------------------------------------------------------
// Static file server
// ---------------------------------------------------------------------------

function startServer(): Promise<Server> {
  return new Promise((resolve) => {
    const server = createServer(
      (req: IncomingMessage, res: ServerResponse) => {
        const url = req.url ?? "/";
        const filePath = join(PROJECT_ROOT, url);

        if (!existsSync(filePath)) {
          res.writeHead(404);
          res.end("Not found: " + url);
          return;
        }

        try {
          const ext = extname(filePath);
          const mime = MIME_TYPES[ext] ?? "application/octet-stream";
          const content = readFileSync(filePath);
          res.writeHead(200, {
            "Content-Type": mime,
            "Cross-Origin-Opener-Policy": "same-origin",
            "Cross-Origin-Embedder-Policy": "require-corp",
          });
          res.end(content);
        } catch {
          res.writeHead(500);
          res.end("Server error");
        }
      },
    );
    server.listen(PORT, () => resolve(server));
  });
}

// ---------------------------------------------------------------------------
// Page helpers
// ---------------------------------------------------------------------------

const TEST_HTML = `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; }
  body { font-family: sans-serif; display: flex; flex-direction: column; }
  #viewer-container {
    background: #fff;
    overflow: hidden;
    flex: 1;
    min-height: 0;
    position: relative;
    height: 700px;
    width: 1200px;
  }
  #viewer-canvas, #viewer-overlay {
    position: absolute; top: 0; left: 0;
    display: block; width: 100%; height: 100%;
  }
  #viewer-overlay { pointer-events: none; }
</style>
</head>
<body>
<div id="viewer-container">
  <canvas id="viewer-canvas"></canvas>
  <canvas id="viewer-overlay"></canvas>
</div>
<script type="module">
import init, { XlView } from '/pkg/xlview.js';

let viewer = null;
let renderPending = false;

function requestRender() {
  if (renderPending) return;
  renderPending = true;
  requestAnimationFrame(() => {
    renderPending = false;
    if (viewer) viewer.render();
  });
}

window.initViewer = async () => {
  await init();
  const canvas = document.getElementById('viewer-canvas');
  const overlay = document.getElementById('viewer-overlay');
  const container = document.getElementById('viewer-container');
  const dpr = window.devicePixelRatio || 1;
  const target = document.querySelector('[data-xlview-scroll]') || canvas.parentElement || container;
  const rect = target.getBoundingClientRect();
  const w = target.clientWidth || rect.width;
  const h = target.clientHeight || rect.height;

  canvas.width = Math.max(1, Math.round(w * dpr));
  canvas.height = Math.max(1, Math.round(h * dpr));
  overlay.width = canvas.width;
  overlay.height = canvas.height;
  canvas.style.width = w + 'px';
  canvas.style.height = h + 'px';
  overlay.style.width = w + 'px';
  overlay.style.height = h + 'px';

  const edpr = w > 0 ? canvas.width / w : dpr;
  viewer = XlView.newWithOverlay
    ? XlView.newWithOverlay(canvas, overlay, edpr)
    : new XlView(canvas, edpr);
  viewer.set_render_callback(requestRender);
  await new Promise(requestAnimationFrame);

  // Re-measure after DOM reflow
  const parent = document.querySelector('[data-xlview-scroll]') || canvas.parentElement || container;
  const pr = parent.getBoundingClientRect();
  const pw = parent.clientWidth || pr.width;
  const ph = parent.clientHeight || pr.height;
  const nw = Math.max(1, Math.round(pw * dpr));
  const nh = Math.max(1, Math.round(ph * dpr));
  if (canvas.width !== nw || canvas.height !== nh) {
    canvas.width = nw; canvas.height = nh;
    overlay.width = nw; overlay.height = nh;
    canvas.style.width = pw + 'px'; canvas.style.height = ph + 'px';
    overlay.style.width = pw + 'px'; overlay.style.height = ph + 'px';
    viewer.resize(nw, nh, dpr);
  }
  window.viewer = viewer;
  return true;
};

window.loadFile = async (url) => {
  const response = await fetch(url);
  const buffer = await response.arrayBuffer();
  viewer.load(new Uint8Array(buffer));
  viewer.render();
  await new Promise(requestAnimationFrame);
  return true;
};

window.forceResize = () => {
  const canvas = document.getElementById('viewer-canvas');
  const overlay = document.getElementById('viewer-overlay');
  const container = document.getElementById('viewer-container');
  const parent = document.querySelector('[data-xlview-scroll]') || canvas.parentElement || container;
  const dpr = window.devicePixelRatio || 1;
  const rect = parent.getBoundingClientRect();
  const w = parent.clientWidth || rect.width;
  const h = parent.clientHeight || rect.height;
  const nw = Math.max(1, Math.round(w * dpr));
  const nh = Math.max(1, Math.round(h * dpr));
  canvas.width = nw; canvas.height = nh;
  overlay.width = nw; overlay.height = nh;
  canvas.style.width = w + 'px'; canvas.style.height = h + 'px';
  overlay.style.width = w + 'px'; overlay.style.height = h + 'px';
  viewer.resize(nw, nh, dpr);
};
</script>
</body>
</html>
`;

async function setupPage(page: Page): Promise<void> {
  // Intercept test.html to serve our harness
  await page.route("**/test.html", (route) => {
    route.fulfill({ status: 200, contentType: "text/html", body: TEST_HTML });
  });

  await page.goto(`http://localhost:${PORT}/test.html`);
  await page.evaluate(() =>
    (window as any).initViewer()
  );
  await page.waitForTimeout(100);
}

async function loadFile(page: Page, name: string): Promise<void> {
  await page.evaluate(
    (url: string) => (window as any).loadFile(url),
    `/test/${name}`,
  );
  await page.waitForTimeout(200);
}

async function getDebug(page: Page): Promise<ScrollDebug> {
  return page.evaluate(() => (window as any).viewer.get_scroll_debug());
}

async function getHeaders(page: Page): Promise<HeaderConfig> {
  return page.evaluate(() => (window as any).viewer.get_header_config());
}

async function renderWithMetrics(page: Page): Promise<RenderMetrics> {
  return page.evaluate(() => (window as any).viewer.render_with_metrics());
}

async function pixelAt(
  page: Page,
  canvasId: string,
  x: number,
  y: number,
): Promise<Pixel> {
  return page.evaluate(
    ({ id, px, py }: { id: string; px: number; py: number }) => {
      const c = document.getElementById(id) as HTMLCanvasElement;
      const ctx = c.getContext("2d")!;
      const d = ctx.getImageData(px, py, 1, 1).data;
      return { r: d[0]!, g: d[1]!, b: d[2]!, a: d[3]! };
    },
    { id: canvasId, px: x, py: y },
  );
}

async function scrollContainer(
  page: Page,
  dx: number,
  dy: number,
): Promise<void> {
  await page.evaluate(
    ({ dx, dy }: { dx: number; dy: number }) => {
      const canvas = document.getElementById("viewer-canvas");
      const divs = Array.from(document.querySelectorAll("div"));
      const container = divs.find((d) => {
        const s = getComputedStyle(d);
        return d.hasAttribute('data-xlview-scroll');
      });
      if (container) {
        container.scrollLeft += dx;
        container.scrollTop += dy;
        container.dispatchEvent(new Event("scroll", { bubbles: true }));
      }
      (window as any).viewer.render();
    },
    { dx, dy },
  );
  await page.waitForTimeout(50);
}

async function setSheet(page: Page, index: number): Promise<void> {
  await page.evaluate(
    (i: number) => {
      (window as any).viewer.set_active_sheet(i);
      (window as any).viewer.render();
    },
    index,
  );
  await page.waitForTimeout(100);
}

async function clickCell(page: Page, x: number, y: number): Promise<void> {
  await page.evaluate(
    ({ x, y }: { x: number; y: number }) => {
      const v = (window as any).viewer;
      v.on_mouse_down(x, y);
      v.on_mouse_up(x, y);
      v.render();
    },
    { x, y },
  );
  await page.waitForTimeout(50);
}

// ---------------------------------------------------------------------------
// Test suites
// ---------------------------------------------------------------------------

async function testMinimal(page: Page): Promise<void> {
  console.log("\n--- minimal.xlsx ---");
  await loadFile(page, "minimal.xlsx");

  const debug = await getDebug(page);
  const hdr = await getHeaders(page);

  assert(debug.show_headers, "headers enabled");
  assert(hdr.visible, "header config visible");
  assert(hdr.col_header_height > 0, "col header height > 0", `${hdr.col_header_height}`);
  assert(hdr.row_header_width > 0, "row header width > 0", `${hdr.row_header_width}`);
  assert(
    debug.visible_start_row < 2,
    "starts at row 0",
    `row=${debug.visible_start_row}`,
  );
  assert(
    debug.container_scroll_left === 0 && debug.container_scroll_top === 0,
    "container scroll at origin",
  );

  // Pixel: header region should be gray #F3F3F3
  const HEADER_GRAY: [number, number, number] = [243, 243, 243];
  const cornerPx = await pixelAt(page, "viewer-overlay", 5, 5);
  assert(
    approxColor(cornerPx, HEADER_GRAY),
    "corner pixel is header gray",
    `rgb(${cornerPx.r},${cornerPx.g},${cornerPx.b})`,
  );
}

async function testKitchenSink(page: Page): Promise<void> {
  console.log("\n--- kitchen_sink_v2.xlsx ---");
  await loadFile(page, "kitchen_sink_v2.xlsx");

  const debug = await getDebug(page);
  assert(debug.show_headers, "headers enabled");
  assert(debug.visible_start_row === 0, "starts at row 0");

  // Sheet tabs
  const sheetCount: number = await page.evaluate(() =>
    (window as any).viewer.sheet_count(),
  );
  assert(sheetCount >= 2, "has multiple sheets", `count=${sheetCount}`);

  // Switch to second sheet and back
  if (sheetCount >= 2) {
    await setSheet(page, 1);
    const sheet2 = await getDebug(page);
    assert(sheet2.show_headers, "sheet 2 headers enabled");
    assert(sheet2.visible_start_row === 0, "sheet 2 starts at row 0");

    await setSheet(page, 0);
    const back = await getDebug(page);
    assert(back.visible_start_row === 0, "back to sheet 1 at row 0");
  }
}

async function testLargeFile(page: Page): Promise<void> {
  console.log("\n--- large_5000x20.xlsx ---");
  await loadFile(page, "large_5000x20.xlsx");

  const debug = await getDebug(page);
  assert(debug.show_headers, "headers enabled");
  // With frozen rows, visible_start_row is the first scrollable row (row 1 when 1 frozen row)
  assert(debug.visible_start_row <= 1, "starts at row 0 or 1", `row=${debug.visible_start_row}`);
  assert(debug.frozen_rows_height > 0, "has frozen rows", `h=${debug.frozen_rows_height}`);

  // Scroll down 500px
  await scrollContainer(page, 0, 500);
  const after = await getDebug(page);
  assert(
    after.visible_start_row > 0 && after.visible_start_row < 50,
    "scroll down moves rows",
    `row=${after.visible_start_row}`,
  );
  assert(
    after.container_scroll_top > 0,
    "container scrollTop > 0",
    `${after.container_scroll_top}`,
  );

  // Frozen row height should still be the same
  assert(
    Math.abs(after.frozen_rows_height - debug.frozen_rows_height) < 1,
    "frozen rows height unchanged after scroll",
  );

  // Scroll right 500px
  await scrollContainer(page, 500, 0);
  const right = await getDebug(page);
  assert(
    right.container_scroll_left > 0,
    "container scrollLeft > 0 after horizontal scroll",
    `${right.container_scroll_left}`,
  );

  // Scroll back to origin
  await page.evaluate(() => {
    const canvas = document.getElementById("viewer-canvas");
    const divs = Array.from(document.querySelectorAll("div"));
    const container = divs.find((d) => {
      const s = getComputedStyle(d);
      return d.hasAttribute('data-xlview-scroll');
    });
    if (container) {
      container.scrollLeft = 0;
      container.scrollTop = 0;
      container.dispatchEvent(new Event("scroll", { bubbles: true }));
    }
    (window as any).viewer.render();
  });
  await page.waitForTimeout(100);
  const origin = await getDebug(page);
  assert(origin.visible_start_row <= 1, "back to top after reset", `row=${origin.visible_start_row}`);
}

async function testHeaderPixels(page: Page): Promise<void> {
  console.log("\n--- header pixel validation ---");
  await loadFile(page, "kitchen_sink_v2.xlsx");

  const hdr = await getHeaders(page);
  if (!hdr.visible) {
    skip("header pixels", "headers not visible");
    return;
  }

  const dpr: number = await page.evaluate(() => window.devicePixelRatio || 1);
  const HEADER_GRAY: [number, number, number] = [243, 243, 243];

  // Sample overlay canvas at header positions (physical pixels = CSS * dpr)
  // Corner area
  const corner = await pixelAt(page, "viewer-overlay", 5, 5);
  assert(
    approxColor(corner, HEADER_GRAY),
    "overlay corner = header gray",
    `rgb(${corner.r},${corner.g},${corner.b})`,
  );

  // Column header (right of row header)
  const colHeaderX = Math.round((hdr.row_header_width + 20) * dpr);
  const colHeader = await pixelAt(page, "viewer-overlay", colHeaderX, 5);
  assert(
    approxColor(colHeader, HEADER_GRAY),
    "col header = header gray",
    `x=${colHeaderX} rgb(${colHeader.r},${colHeader.g},${colHeader.b})`,
  );

  // Row header (below col header)
  const rowHeaderY = Math.round((hdr.col_header_height + 10) * dpr);
  const rowHeader = await pixelAt(page, "viewer-overlay", 5, rowHeaderY);
  assert(
    approxColor(rowHeader, HEADER_GRAY),
    "row header = header gray",
    `y=${rowHeaderY} rgb(${rowHeader.r},${rowHeader.g},${rowHeader.b})`,
  );
}

async function testSelection(page: Page): Promise<void> {
  console.log("\n--- selection ---");
  await loadFile(page, "kitchen_sink_v2.xlsx");

  const hdr = await getHeaders(page);
  // Click a cell in the data area (CSS coords, past headers)
  const cellX = hdr.row_header_width + 60;
  const cellY = hdr.col_header_height + 30;
  await clickCell(page, cellX, cellY);

  const sel: number[] | null = await page.evaluate(() =>
    (window as any).viewer.get_selection(),
  );
  assert(sel !== null && sel.length === 4, "selection set after click", `sel=${sel}`);
}

async function testComments(page: Page): Promise<void> {
  console.log("\n--- comments ---");
  await loadFile(page, "test_comments.xlsx");

  const debug = await getDebug(page);
  assert(debug.show_headers, "headers enabled");

  // Try to get comment at a known position (cell A1 area)
  const hdr = await getHeaders(page);
  const comment: string | null = await page.evaluate(
    ({ x, y }: { x: number; y: number }) =>
      (window as any).viewer.get_comment_at(x, y),
    { x: hdr.row_header_width + 5, y: hdr.col_header_height + 5 },
  );
  // Just check API doesn't crash - comment may or may not be at A1
  ok("get_comment_at does not crash", comment ? "found" : "none at pos");
}

async function testScrollPerformance(page: Page): Promise<void> {
  console.log("\n--- scroll performance (large_5000x20.xlsx) ---");
  await loadFile(page, "large_5000x20.xlsx");

  // Warm up
  await renderWithMetrics(page);

  // Measure 20 scroll + render cycles
  const metrics = await page.evaluate(() => {
    const v = (window as any).viewer;
    const canvas = document.getElementById("viewer-canvas");
    const divs = Array.from(document.querySelectorAll("div"));
    const container = divs.find((d: Element) => {
      const s = getComputedStyle(d);
      return d.hasAttribute('data-xlview-scroll');
    });
    if (!container) return null;

    const results: Array<{ total_ms: number; draw_ms: number; skipped: boolean }> = [];
    for (let i = 0; i < 20; i++) {
      (container as HTMLElement).scrollTop += 60;
      container.dispatchEvent(new Event("scroll", { bubbles: true }));
      const m = v.render_with_metrics();
      results.push({ total_ms: m.total_ms, draw_ms: m.draw_ms, skipped: m.skipped });
    }
    return results;
  });

  if (!metrics) {
    skip("scroll perf", "container not found");
    return;
  }

  const nonSkipped = metrics.filter((m) => !m.skipped);
  if (nonSkipped.length === 0) {
    skip("scroll perf", "all frames skipped");
    return;
  }

  const totals = nonSkipped.map((m) => m.total_ms).sort((a, b) => a - b);
  const draws = nonSkipped.map((m) => m.draw_ms).sort((a, b) => a - b);
  const p90Idx = Math.min(totals.length - 1, Math.floor(0.9 * totals.length));
  const p90Total = totals[p90Idx]!;
  const p90Draw = draws[p90Idx]!;
  const median = totals[Math.floor(totals.length / 2)]!;

  console.log(
    `    frames=${nonSkipped.length} median=${median.toFixed(1)}ms p90_total=${p90Total.toFixed(1)}ms p90_draw=${p90Draw.toFixed(1)}ms`,
  );

  assert(
    p90Total < PERF.scroll_p90,
    `scroll p90 < ${PERF.scroll_p90}ms`,
    `${p90Total.toFixed(1)}ms`,
  );
}

async function testRenderPerformance(page: Page): Promise<void> {
  console.log("\n--- render performance ---");
  await loadFile(page, "kitchen_sink_v2.xlsx");

  // Force invalidate + re-render 10 times
  const metrics = await page.evaluate(() => {
    const v = (window as any).viewer;
    const results: Array<{ total_ms: number; draw_ms: number }> = [];
    for (let i = 0; i < 10; i++) {
      v.invalidate();
      const m = v.render_with_metrics();
      results.push({ total_ms: m.total_ms, draw_ms: m.draw_ms });
    }
    return results;
  });

  const totals = metrics.map((m) => m.total_ms).sort((a, b) => a - b);
  const p90Idx = Math.min(totals.length - 1, Math.floor(0.9 * totals.length));
  const p90 = totals[p90Idx]!;
  const median = totals[Math.floor(totals.length / 2)]!;

  console.log(`    frames=${metrics.length} median=${median.toFixed(1)}ms p90=${p90.toFixed(1)}ms`);

  assert(
    p90 < PERF.render_p90,
    `render p90 < ${PERF.render_p90}ms`,
    `${p90.toFixed(1)}ms`,
  );
}

async function testColors(page: Page): Promise<void> {
  console.log("\n--- colors_test.xlsx ---");
  await loadFile(page, "colors_test.xlsx");

  const debug = await getDebug(page);
  assert(debug.show_headers, "headers enabled");
  assert(debug.visible_start_row === 0, "starts at row 0");

  // Just ensure non-blank canvas
  const hdr = await getHeaders(page);
  const dpr: number = await page.evaluate(() => window.devicePixelRatio || 1);
  const cellX = Math.round((hdr.row_header_width + 40) * dpr);
  const cellY = Math.round((hdr.col_header_height + 20) * dpr);
  const px = await pixelAt(page, "viewer-canvas", cellX, cellY);
  assert(px.a > 0, "canvas has content", `alpha=${px.a}`);
}

async function testStyledFile(page: Page): Promise<void> {
  console.log("\n--- styled.xlsx ---");
  await loadFile(page, "styled.xlsx");

  const debug = await getDebug(page);
  assert(debug.show_headers, "headers enabled");
  assert(debug.visible_start_row === 0, "starts at row 0");
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<number> {
  console.log("xlview E2E test suite\n");

  // Check WASM build exists
  const wasmPath = join(PROJECT_ROOT, "pkg/xlview_bg.wasm");
  if (!existsSync(wasmPath)) {
    console.error("ERROR: WASM not built. Run: wasm-pack build --target web --dev");
    return 1;
  }

  const server = await startServer();
  console.log(`Server: http://localhost:${PORT}`);

  const browser: Browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
  });
  const page: Page = await context.newPage();

  // Capture errors
  const errors: string[] = [];
  page.on("pageerror", (err) => errors.push(err.message));

  try {
    await setupPage(page);

    // Content correctness
    await testMinimal(page);
    await testKitchenSink(page);
    await testLargeFile(page);
    await testColors(page);
    await testStyledFile(page);

    // Features
    await testHeaderPixels(page);
    await testSelection(page);
    await testComments(page);

    // Performance
    await testScrollPerformance(page);
    await testRenderPerformance(page);

    // Page errors
    if (errors.length > 0) {
      console.log("\n--- page errors ---");
      for (const e of errors) {
        fail("page error", e);
      }
    }
  } catch (err) {
    fail("FATAL", (err as Error).message);
    console.error(err);
  } finally {
    await browser.close();
    server.close();
  }

  // Summary
  console.log("\n" + "=".repeat(50));
  console.log(
    `\x1b[${failed > 0 ? "31" : "32"}m${passed} passed, ${failed} failed, ${skipped} skipped\x1b[0m`,
  );
  if (failures.length > 0) {
    console.log("\nFailures:");
    for (const f of failures) {
      console.log(`  - ${f}`);
    }
  }
  console.log("=".repeat(50));

  return failed > 0 ? 1 : 0;
}

process.exit(await main());

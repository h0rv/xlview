#!/usr/bin/env node
/**
 * Headless browser test for scroll/header rendering issues.
 *
 * Run with: node tests/browser/run_scroll_test.js
 *
 * Requires: npm install playwright (or npx playwright install)
 */

import { chromium } from "playwright";
import { createServer, type IncomingMessage, type ServerResponse } from "http";
import { readFileSync, existsSync } from "fs";
import { join, extname } from "path";
import { fileURLToPath } from "url";
import type { Server } from "http";
import type { Browser, Page } from "playwright";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const PROJECT_ROOT = join(__dirname, "../..");

const MIME_TYPES: Record<string, string> = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".mjs": "application/javascript",
  ".wasm": "application/wasm",
  ".json": "application/json",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

// Simple static file server
function startServer(port: number): Promise<Server> {
  return new Promise((resolve) => {
    const server = createServer((req: IncomingMessage, res: ServerResponse) => {
      const filePath = join(
        PROJECT_ROOT,
        req.url === "/" ? "/tests/browser/scroll_test.html" : req.url!,
      );

      if (!existsSync(filePath)) {
        res.writeHead(404);
        res.end("Not found: " + req.url);
        return;
      }

      const ext = extname(filePath);
      const mime = MIME_TYPES[ext] ?? "application/octet-stream";

      try {
        const content = readFileSync(filePath);
        res.writeHead(200, {
          "Content-Type": mime,
          "Cross-Origin-Opener-Policy": "same-origin",
          "Cross-Origin-Embedder-Policy": "require-corp",
        });
        res.end(content);
      } catch (e) {
        res.writeHead(500);
        res.end("Error: " + (e as Error).message);
      }
    });

    server.listen(port, () => {
      resolve(server);
    });
  });
}

interface TestPage {
  name: string;
  url: string;
  needsFileClick?: boolean;
}

interface ScrollDebug {
  viewport_scroll_x: number;
  viewport_scroll_y: number;
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

async function runTests(): Promise<void> {
  console.log("Starting scroll/header rendering tests...\n");

  // Start server
  const PORT = 8765;
  const server = await startServer(PORT);
  console.log(`Server running on http://localhost:${PORT}`);

  // Launch browser
  const browser: Browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
  });
  const page: Page = await context.newPage();

  // Collect console logs
  const logs: string[] = [];
  page.on("console", (msg) => {
    logs.push(`[${msg.type()}] ${msg.text()}`);
  });

  page.on("pageerror", (err) => {
    logs.push(`[error] ${err.message}`);
  });

  let testsPassed = 0;
  let testsFailed = 0;

  function pass(name: string, details = ""): void {
    testsPassed++;
    console.log(`  PASS ${name}`);
    if (details) console.log(`     ${details}`);
  }

  function fail(name: string, details = ""): void {
    testsFailed++;
    console.log(`  FAIL ${name}`);
    if (details) console.log(`     ${details}`);
  }

  try {
    const testPages: TestPage[] = [
      {
        name: "bench-style",
        url: `http://localhost:${PORT}/tests/browser/bench_style_test.html`,
      },
      {
        name: "index-style",
        url: `http://localhost:${PORT}/tests/browser/index_style_test.html`,
      },
      {
        name: "index-demo",
        url: `http://localhost:${PORT}/index.html`,
        needsFileClick: true,
      },
    ];

    const runTestsForPage = async (testPage: TestPage): Promise<void> => {
      console.log(`\nLoading ${testPage.name} test page...`);
      await page.goto(testPage.url);

      // Wait for viewer to be initialized
      console.log("   Waiting for WASM initialization...");
      await page
        .waitForFunction(
          () => {
            const w = window as unknown as Record<string, unknown>;
            const v = w.viewer as Record<string, unknown> | undefined;
            return v && typeof v.get_scroll_debug === "function";
          },
          { timeout: 10000 },
        )
        .catch(async () => {
          // If not auto-initialized, try manual init
          console.log("   Auto-init failed, trying manual init...");
          await page.evaluate(async () => {
            const w = window as unknown as Record<string, unknown>;
            if (typeof w.initViewer === "function" && !w.viewer) {
              await (w.initViewer as () => Promise<unknown>)();
            }
          });
          await page.waitForTimeout(1000);
        });

      const viewerReady = await page.evaluate(() => {
        const w = window as unknown as Record<string, unknown>;
        const v = w.viewer as Record<string, unknown> | undefined;
        return v && typeof v.get_scroll_debug === "function";
      });

      if (!viewerReady && !testPage.needsFileClick) {
        console.log("   FAIL Viewer failed to initialize");
        fail(
          `${testPage.name}: viewer init`,
          "viewer.get_scroll_debug() not available",
        );
        return;
      }

      if (viewerReady) {
        console.log("   PASS Viewer initialized");
      } else {
        console.log("   PASS Viewer initialized (not exposed on window)");
      }

      if (testPage.needsFileClick) {
        console.log("   Loading kitchen_sink_v2.xlsx via index.html UI...");
        // Wait for manifest buttons to render and click the v2 file.
        await page.waitForFunction(
          () => {
            return Array.from(document.querySelectorAll("button")).some(
              (btn) =>
                btn.textContent &&
                btn.textContent.includes("kitchen_sink_v2.xlsx"),
            );
          },
          { timeout: 10000 },
        );

        await page.evaluate(() => {
          const btn = Array.from(document.querySelectorAll("button")).find(
            (b) =>
              b.textContent && b.textContent.includes("kitchen_sink_v2.xlsx"),
          );
          if (btn) {
            btn.click();
          }
        });

        // Wait for viewer to load and render
        await page.waitForTimeout(800);
      }

      // Ensure any post-setup resize runs before tests.
      await page.evaluate(() => {
        const w = window as unknown as Record<string, unknown>;
        if (typeof w.force_resize === "function") {
          (w.force_resize as () => void)();
        }
      });
      await page.waitForTimeout(200);

      // ========== TESTS ==========
      console.log("\nRunning tests...\n");

      if (!viewerReady && testPage.needsFileClick) {
        // Visual pixel validation via canvas sampling for index.html demo
        const pixel = await page.evaluate(() => {
          const canvas = document.getElementById(
            "viewer-canvas",
          ) as HTMLCanvasElement | null;
          if (!canvas) return null;
          const ctx = canvas.getContext("2d");
          if (!ctx) return null;
          const getPixel = (
            x: number,
            y: number,
          ): [number, number, number, number] => {
            const data = ctx.getImageData(x, y, 1, 1).data;
            return [data[0]!, data[1]!, data[2]!, data[3]!];
          };
          return {
            corner: getPixel(10, 10),
            rowHeader: getPixel(10, 30),
            colHeader: getPixel(60, 10),
            cellA1: getPixel(60, 30),
          };
        });

        if (!pixel) {
          fail(
            `${testPage.name}: canvas pixel sample`,
            "canvas or context unavailable",
          );
        } else {
          pass(`${testPage.name}: canvas pixel sample`);

          const approx = (
            px: [number, number, number, number],
            target: [number, number, number],
            tol = 18,
          ): boolean =>
            Math.abs(px[0] - target[0]) <= tol &&
            Math.abs(px[1] - target[1]) <= tol &&
            Math.abs(px[2] - target[2]) <= tol;

          const HEADER: [number, number, number] = [243, 243, 243]; // #F3F3F3
          const BLUE: [number, number, number] = [68, 114, 196]; // #4472C4 (A1 header fill)

          if (
            approx(pixel.corner, HEADER) &&
            approx(pixel.rowHeader, HEADER) &&
            approx(pixel.colHeader, HEADER)
          ) {
            pass(`${testPage.name}: header pixels aligned`);
          } else {
            fail(
              `${testPage.name}: header pixels misaligned`,
              `corner=${pixel.corner} rowHeader=${pixel.rowHeader} colHeader=${pixel.colHeader}`,
            );
          }

          if (approx(pixel.cellA1, BLUE)) {
            pass(`${testPage.name}: A1 header cell at expected position`);
          } else {
            fail(
              `${testPage.name}: A1 header cell not at expected position`,
              `cellA1=${pixel.cellA1} expected~=${BLUE}`,
            );
          }
        }

        console.log("\nTaking screenshot...");
        await page.screenshot({
          path: join(
            PROJECT_ROOT,
            `tests/browser/test-result-${testPage.name}.png`,
          ),
          fullPage: false,
        });
        console.log(
          `   Saved to tests/browser/test-result-${testPage.name}.png`,
        );
        return;
      }

      // Test 1: Get scroll debug info
      const debug = (await page.evaluate(() => {
        const w = window as unknown as Record<string, unknown>;
        const v = w.viewer as Record<string, Function> | undefined;
        if (!v || !v.get_scroll_debug) {
          return null;
        }
        return v.get_scroll_debug();
      })) as ScrollDebug | null;

      if (!debug) {
        fail(
          `${testPage.name}: get_scroll_debug`,
          "viewer.get_scroll_debug() not available",
        );
      } else {
        pass(`${testPage.name}: get_scroll_debug`);
        console.log("\n   Debug info:");
        console.log(
          `     viewport_scroll: (${debug.viewport_scroll_x}, ${debug.viewport_scroll_y})`,
        );
        console.log(
          `     container_scroll: (${debug.container_scroll_left}, ${debug.container_scroll_top})`,
        );
        console.log(
          `     frozen: (${debug.frozen_cols_width}, ${debug.frozen_rows_height})`,
        );
        console.log(
          `     visible_rows: ${debug.visible_start_row} - ${debug.visible_end_row}`,
        );
        console.log(`     show_headers: ${debug.show_headers}`);
        console.log("");

        // Test 2: Initial scroll position
        const frozenX = debug.frozen_cols_width || 0;
        const frozenY = debug.frozen_rows_height || 0;
        const scrollCorrectX = Math.abs(debug.viewport_scroll_x - frozenX) < 1;
        const scrollCorrectY = Math.abs(debug.viewport_scroll_y - frozenY) < 1;

        if (scrollCorrectX && scrollCorrectY) {
          pass(
            `${testPage.name}: initial scroll at frozen boundary`,
            `Expected: (${frozenX}, ${frozenY}), Got: (${debug.viewport_scroll_x}, ${debug.viewport_scroll_y})`,
          );
        } else {
          fail(
            `${testPage.name}: initial scroll at frozen boundary`,
            `Expected: (${frozenX}, ${frozenY}), Got: (${debug.viewport_scroll_x}, ${debug.viewport_scroll_y})`,
          );
        }

        // Test 3: Container scroll should be 0,0
        if (
          debug.container_scroll_left === 0 &&
          debug.container_scroll_top === 0
        ) {
          pass(`${testPage.name}: container scroll at origin`);
        } else {
          fail(
            `${testPage.name}: container scroll at origin`,
            `Got: (${debug.container_scroll_left}, ${debug.container_scroll_top})`,
          );
        }

        // Test 4: First visible row should NOT be ~30
        if (debug.visible_start_row < 5) {
          pass(
            `${testPage.name}: first visible row near 0`,
            `visible_start_row: ${debug.visible_start_row}`,
          );
        } else if (debug.visible_start_row >= 25) {
          fail(
            `${testPage.name}: ROW 30 BUG DETECTED`,
            `visible_start_row: ${debug.visible_start_row} (should be near 0)`,
          );
        } else {
          fail(
            `${testPage.name}: first visible row unexpected`,
            `visible_start_row: ${debug.visible_start_row}`,
          );
        }

        // Test 5: Headers should be visible
        if (debug.show_headers === true) {
          pass(`${testPage.name}: headers enabled`);
        } else {
          fail(
            `${testPage.name}: headers not enabled`,
            `show_headers: ${debug.show_headers}`,
          );
        }

        // Test 5b: Check header config dimensions
        const headerConfig = (await page.evaluate(() => {
          const w = window as unknown as Record<string, unknown>;
          const v = w.viewer as Record<string, Function> | undefined;
          return v?.get_header_config ? v.get_header_config() : null;
        })) as HeaderConfig | null;
        if (headerConfig) {
          console.log(`\n   Header config:`);
          console.log(
            `     row_header_width: ${headerConfig.row_header_width}`,
          );
          console.log(
            `     col_header_height: ${headerConfig.col_header_height}`,
          );
          console.log(`     visible: ${headerConfig.visible}`);

          if (
            headerConfig.col_header_height > 0 &&
            headerConfig.row_header_width > 0
          ) {
            pass(
              `${testPage.name}: header dimensions non-zero`,
              `col_header_height: ${headerConfig.col_header_height}, row_header_width: ${headerConfig.row_header_width}`,
            );
          } else {
            fail(
              `${testPage.name}: header dimensions are zero`,
              `col_header_height: ${headerConfig.col_header_height}, row_header_width: ${headerConfig.row_header_width}`,
            );
          }
        }

        // Test 5c: Canvas sizing should match scroll container client size
        const sizing = await page.evaluate(() => {
          const canvas = document.getElementById(
            "viewer-canvas",
          ) as HTMLCanvasElement | null;
          if (!canvas) return null;
          const candidates = Array.from(document.querySelectorAll("div"));
          const container = candidates.find((d) => {
            const style = getComputedStyle(d);
            return d.hasAttribute('data-xlview-scroll');
          });
          if (!container) return null;
          const rect = canvas.getBoundingClientRect();
          return {
            canvasCssWidth: rect.width,
            canvasCssHeight: rect.height,
            containerClientWidth: container.clientWidth,
            containerClientHeight: container.clientHeight,
          };
        });

        if (sizing) {
          const wDiff = Math.abs(
            sizing.canvasCssWidth - sizing.containerClientWidth,
          );
          const hDiff = Math.abs(
            sizing.canvasCssHeight - sizing.containerClientHeight,
          );
          if (wDiff <= 1 && hDiff <= 1) {
            pass(
              `${testPage.name}: canvas matches scroll container size`,
              `dw=${wDiff.toFixed(2)} dh=${hDiff.toFixed(2)}`,
            );
          } else {
            fail(
              `${testPage.name}: canvas size mismatch`,
              `canvas=(${sizing.canvasCssWidth}x${sizing.canvasCssHeight}) container=(${sizing.containerClientWidth}x${sizing.containerClientHeight})`,
            );
          }
        } else {
          fail(
            `${testPage.name}: canvas sizing check`,
            "canvas or scroll container missing",
          );
        }

        // Take screenshot BEFORE scrolling
        console.log("\nTaking INITIAL screenshot (before scroll)...");
        await page.screenshot({
          path: join(
            PROJECT_ROOT,
            `tests/browser/test-initial-${testPage.name}.png`,
          ),
          fullPage: false,
        });
        console.log(
          `   Saved to tests/browser/test-initial-${testPage.name}.png`,
        );

        // Test 6: Native scroll container scroll should not jump
        console.log("\n   Scrolling native container down 100px...");
        const nativeScroll = await page.evaluate(() => {
          const w = window as unknown as Record<string, unknown>;
          const v = w.viewer as Record<string, Function> | undefined;
          if (!v || !v.get_scroll_debug) {
            return { skipped: true as const, reason: "viewer missing" };
          }
          const canvas = document.getElementById("viewer-canvas");
          const candidates = Array.from(document.querySelectorAll("div"));
          const container = candidates.find((d) => {
            const style = getComputedStyle(d);
            return d.hasAttribute('data-xlview-scroll');
          });
          if (!container) {
            return {
              skipped: true as const,
              reason: "scroll container not found",
            };
          }
          container.scrollTop = 100;
          container.scrollLeft = 0;
          container.dispatchEvent(new Event("scroll", { bubbles: true }));
          v.render();
          return {
            skipped: false as const,
            containerTop: container.scrollTop,
            containerHeight: container.clientHeight,
            containerScrollHeight: container.scrollHeight,
            debug: v.get_scroll_debug() as ScrollDebug,
          };
        });

        if (nativeScroll && nativeScroll.skipped) {
          console.log(`   Skipping native scroll test: ${nativeScroll.reason}`);
        } else if (
          nativeScroll &&
          !nativeScroll.skipped &&
          nativeScroll.debug
        ) {
          const debugAfterScroll = nativeScroll.debug;
          console.log(
            `     Container scrollTop = ${nativeScroll.containerTop}`,
          );
          console.log(
            `     Container height = ${nativeScroll.containerHeight}, scrollHeight = ${nativeScroll.containerScrollHeight}`,
          );
          console.log(
            `     After scroll: viewport_scroll_y = ${debugAfterScroll.viewport_scroll_y}`,
          );
          console.log(
            `     After scroll: visible_start_row = ${debugAfterScroll.visible_start_row}`,
          );

          // After scrolling 100px, we should be around row 5 (100/20), not row 30+
          if (debugAfterScroll.visible_start_row < 10) {
            pass(
              `${testPage.name}: native scroll does not jump to row 30`,
              `After 100px scroll: row ${debugAfterScroll.visible_start_row}`,
            );
          } else {
            fail(
              `${testPage.name}: native scroll jumped unexpectedly`,
              `After 100px scroll: row ${debugAfterScroll.visible_start_row} (expected < 10)`,
            );
          }
        }
      }

      // Take screenshot
      console.log("\nTaking screenshot...");
      await page.screenshot({
        path: join(
          PROJECT_ROOT,
          `tests/browser/test-result-${testPage.name}.png`,
        ),
        fullPage: false,
      });
      console.log(`   Saved to tests/browser/test-result-${testPage.name}.png`);
    };

    for (const testPage of testPages) {
      await runTestsForPage(testPage);
    }
  } catch (error) {
    console.error("\nTest error:", (error as Error).message);
    testsFailed++;
  }

  // Print console logs if there were errors
  if (logs.some((l) => l.includes("[error]"))) {
    console.log("\nBrowser console logs:");
    logs.forEach((l) => console.log("   " + l));
  }

  // Cleanup
  await browser.close();
  server.close();

  // Summary
  console.log("\n" + "=".repeat(50));
  console.log(`Results: ${testsPassed} passed, ${testsFailed} failed`);
  console.log("=".repeat(50) + "\n");

  process.exit(testsFailed > 0 ? 1 : 0);
}

function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return 0;
  const idx = Math.min(
    sorted.length - 1,
    Math.max(0, Math.floor(p * (sorted.length - 1))),
  );
  return sorted[idx]!;
}

function median(values: number[]): number {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
}

interface TuneResult {
  windowMs: number;
  error?: string;
  medianDraw?: number;
  p90Draw?: number;
  medianTotal?: number;
  p90Total?: number;
}

async function runTune(): Promise<void> {
  console.log("Tuning fast scroll window...\n");

  const PORT = 8765;
  const server = await startServer(PORT);
  console.log(`Server running on http://localhost:${PORT}`);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
  });
  const page = await context.newPage();

  const windows = [40, 60, 80, 100, 120, 140, 160, 200];
  const results: TuneResult[] = [];

  for (const windowMs of windows) {
    await page.goto(
      `http://localhost:${PORT}/tests/browser/index_style_test.html`,
    );

    await page.waitForFunction(
      () => {
        const w = window as unknown as Record<string, unknown>;
        const v = w.viewer as Record<string, unknown> | undefined;
        return v && typeof v.get_scroll_debug === "function";
      },
      { timeout: 15000 },
    );

    await page.evaluate(() => {
      const w = window as unknown as Record<string, unknown>;
      if (typeof w.force_resize === "function") {
        (w.force_resize as () => void)();
      }
    });

    await page.evaluate((ms: number) => {
      const w = window as unknown as Record<string, unknown>;
      const v = w.viewer as Record<string, Function>;
      v.set_fast_scroll_auto(false);
      v.set_fast_scroll_window_ms(ms);
    }, windowMs);

    const run = await page.evaluate(() => {
      const canvas = document.getElementById(
        "viewer-canvas",
      ) as HTMLCanvasElement | null;
      const candidates = Array.from(document.querySelectorAll("div"));
      const container = candidates.find((d) => {
        const style = getComputedStyle(d);
        return d.hasAttribute('data-xlview-scroll');
      });
      if (!container) return { error: "scroll container not found" };

      container.scrollTop = 0;
      container.scrollLeft = 0;
      container.dispatchEvent(new Event("scroll", { bubbles: true }));

      const metrics: Array<Record<string, number>> = [];
      const steps = 48;
      const stepX = 120;
      const stepY = 80;

      const w = window as unknown as Record<string, unknown>;
      const v = w.viewer as Record<string, Function>;

      for (let i = 0; i < steps; i++) {
        const dir = i % 2 === 0 ? 1 : -1;
        const diag = i % 3 === 0 ? 1 : -1;
        container.scrollLeft += dir * stepX;
        container.scrollTop += diag * stepY;
        container.dispatchEvent(new Event("scroll", { bubbles: true }));
        if (v.render_with_metrics) {
          metrics.push(v.render_with_metrics() as Record<string, number>);
        } else {
          v.render();
        }
      }

      const settle = v.render_with_metrics
        ? (v.render_with_metrics() as Record<string, number>)
        : null;
      return { metrics, settle };
    });

    if ("error" in run && run.error) {
      results.push({ windowMs, error: run.error as string });
      continue;
    }

    const runData = run as {
      metrics: Array<Record<string, number>>;
      settle: Record<string, number> | null;
    };
    const draw = runData.metrics
      .map((m) => m.draw_ms)
      .filter((n): n is number => typeof n === "number");
    const total = runData.metrics
      .map((m) => m.total_ms)
      .filter((n): n is number => typeof n === "number");
    const sortedDraw = [...draw].sort((a, b) => a - b);
    const sortedTotal = [...total].sort((a, b) => a - b);

    results.push({
      windowMs,
      medianDraw: median(draw),
      p90Draw: percentile(sortedDraw, 0.9),
      medianTotal: median(total),
      p90Total: percentile(sortedTotal, 0.9),
    });
  }

  await browser.close();
  server.close();

  console.log("\nFast scroll sweep results (ms):");
  console.table(results);

  const valid = results.filter((r) => !r.error);
  if (valid.length) {
    const best = [...valid].sort(
      (a, b) => (a.p90Total ?? Infinity) - (b.p90Total ?? Infinity),
    )[0]!;
    console.log(
      `\nRecommended window: ${best.windowMs}ms (lowest p90 total = ${best.p90Total?.toFixed(2)}ms)`,
    );
  }
}

if (process.argv.includes("--tune")) {
  runTune().catch((err) => {
    console.error(err);
    process.exit(1);
  });
} else {
  runTests().catch((err) => {
    console.error("Fatal error:", err);
    process.exit(1);
  });
}

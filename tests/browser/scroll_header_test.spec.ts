import { test, expect } from "@playwright/test";

/**
 * Browser tests for scroll and header rendering issues.
 *
 * These tests verify:
 * 1. Headers render on first load
 * 2. Scroll starts at row 0, not row 30
 * 3. Scroll coordinates are correctly mapped from container to viewport
 */

const TEST_PAGE = `
<!DOCTYPE html>
<html>
<head>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: sans-serif; }
    #container {
      width: 1200px;
      height: 800px;
      position: relative;
      border: 1px solid #ccc;
    }
    #canvas {
      position: absolute;
      inset: 0;
    }
  </style>
</head>
<body>
  <div id="container">
    <canvas id="canvas"></canvas>
  </div>
  <script type="module">
    window.initViewer = async function() {
      const wasm = await import('/pkg/xlview.js');
      await wasm.default();

      const canvas = document.getElementById('canvas');
      const container = document.getElementById('container');
      const dpr = window.devicePixelRatio || 1;
      const rect = container.getBoundingClientRect();
      const width = container.clientWidth || rect.width;
      const height = container.clientHeight || rect.height;

      canvas.width = Math.floor(width * dpr);
      canvas.height = Math.floor(height * dpr);
      canvas.style.width = width + 'px';
      canvas.style.height = height + 'px';

      const viewer = new wasm.XlView(canvas, dpr);
      viewer.resize(canvas.width, canvas.height, dpr);

      window.viewer = viewer;
      window.wasm = wasm;
      return viewer;
    };

    window.loadFile = async function(url) {
      const response = await fetch(url);
      const buffer = await response.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      window.viewer.load(bytes);
      window.viewer.render();
    };
  </script>
</body>
</html>
`;

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

test.describe("Scroll and Header Tests", () => {
  test.beforeEach(async ({ page }) => {
    // Serve the test page
    await page.route("**/test.html", (route) => {
      route.fulfill({
        status: 200,
        contentType: "text/html",
        body: TEST_PAGE,
      });
    });

    // Allow WASM and pkg files to load from file system
    await page.route("**/pkg/**", async (route) => {
      // In real test, would serve from filesystem
      route.continue();
    });
  });

  test("scroll debug shows row 0 on initial load", async ({ page }) => {
    await page.goto("/test.html");

    // Initialize viewer
    await page.evaluate(() =>
      (
        window as unknown as Record<string, () => Promise<unknown>>
      ).initViewer(),
    );

    // Load a test file
    await page.evaluate(() =>
      (
        window as unknown as Record<string, (url: string) => Promise<unknown>>
      ).loadFile("/fixtures/kitchen_sink.xlsx"),
    );

    // Get scroll debug info
    const debug = (await page.evaluate(() => {
      return (
        window as unknown as Record<string, Record<string, Function>>
      ).viewer.get_scroll_debug();
    })) as ScrollDebug;

    console.log("Scroll debug:", debug);

    // Verify initial scroll position
    expect(debug.viewport_scroll_x).toBe(debug.frozen_cols_width || 0);
    expect(debug.viewport_scroll_y).toBe(debug.frozen_rows_height || 0);
    expect(debug.visible_start_row).toBe(0);
    expect(debug.container_scroll_left).toBe(0);
    expect(debug.container_scroll_top).toBe(0);
  });

  test("headers are visible on first render", async ({ page }) => {
    await page.goto("/test.html");
    await page.evaluate(() =>
      (
        window as unknown as Record<string, () => Promise<unknown>>
      ).initViewer(),
    );
    await page.evaluate(() =>
      (
        window as unknown as Record<string, (url: string) => Promise<unknown>>
      ).loadFile("/fixtures/kitchen_sink.xlsx"),
    );

    const debug = (await page.evaluate(() => {
      return (
        window as unknown as Record<string, Record<string, Function>>
      ).viewer.get_scroll_debug();
    })) as ScrollDebug;

    // Headers should be enabled
    expect(debug.show_headers).toBe(true);

    // Take a screenshot for visual verification
    await page.screenshot({ path: "test-results/initial-render.png" });
  });

  test("scroll does not jump to row 30", async ({ page }) => {
    await page.goto("/test.html");
    await page.evaluate(() =>
      (
        window as unknown as Record<string, () => Promise<unknown>>
      ).initViewer(),
    );
    await page.evaluate(() =>
      (
        window as unknown as Record<string, (url: string) => Promise<unknown>>
      ).loadFile("/fixtures/kitchen_sink.xlsx"),
    );

    // Small scroll
    await page.evaluate(() => {
      const v = (window as unknown as Record<string, Record<string, Function>>)
        .viewer;
      v.scroll(0, 100);
      v.render();
    });

    const debug = (await page.evaluate(() => {
      return (
        window as unknown as Record<string, Record<string, Function>>
      ).viewer.get_scroll_debug();
    })) as ScrollDebug;

    console.log("After scroll debug:", debug);

    // Scroll should be approximately 100px from start, NOT jumping to row 30 (~600px)
    expect(debug.viewport_scroll_y).toBeLessThan(200);
    expect(debug.visible_start_row).toBeLessThan(10);
  });
});

import type {
  BenchProvider,
  ProviderCreateOpts,
} from "../../src/demo/types.js";

export async function createXlviewProvider({
  canvas,
  domRoot,
  dpr,
  width,
  height,
}: ProviderCreateOpts): Promise<BenchProvider> {
  const wasm = await import("../../pkg/xlview.js");
  await wasm.default();

  const parseMetrics = wasm.parse_xlsx_metrics ?? null;

  if (domRoot) {
    domRoot.style.display = "none";
    domRoot.innerHTML = "";
  }
  if (canvas) {
    canvas.style.display = "block";
  }

  const viewer = new wasm.XlView(canvas, dpr);
  viewer.resize(width, height, dpr);
  // After native scroll setup, the canvas is reparented; re-measure to avoid
  // sizing mismatches from the tab bar reducing available height.
  await new Promise(requestAnimationFrame);
  const parent = canvas.parentElement;
  if (parent) {
    const rect = parent.getBoundingClientRect();
    const cssWidth = parent.clientWidth || rect.width;
    const cssHeight = parent.clientHeight || rect.height;
    const nextWidth = Math.max(1, Math.round(cssWidth * dpr));
    const nextHeight = Math.max(1, Math.round(cssHeight * dpr));
    if (canvas.width !== nextWidth || canvas.height !== nextHeight) {
      canvas.width = nextWidth;
      canvas.height = nextHeight;
      canvas.style.width = `${cssWidth}px`;
      canvas.style.height = `${cssHeight}px`;
      viewer.resize(canvas.width, canvas.height, dpr);
    }
  }
  const loadMetrics = viewer.load_with_metrics
    ? viewer.load_with_metrics.bind(viewer)
    : null;
  const renderMetrics = viewer.render_with_metrics
    ? viewer.render_with_metrics.bind(viewer)
    : null;
  const scrollMetrics = viewer.scroll_with_metrics
    ? viewer.scroll_with_metrics.bind(viewer)
    : null;

  return {
    id: "xlview",
    label: "xlview",
    async parse(bytes: Uint8Array) {
      if (parseMetrics) {
        return parseMetrics(bytes);
      }
      return wasm.parse_xlsx_to_js(bytes);
    },
    async load(bytes: Uint8Array) {
      if (loadMetrics) {
        return loadMetrics(bytes);
      }
      return viewer.load(bytes);
    },
    async render() {
      viewer.invalidate();
      if (renderMetrics) {
        return renderMetrics();
      }
      return viewer.render();
    },
    async scroll({ dx, dy }) {
      if (scrollMetrics) {
        return scrollMetrics(dx ?? 0, dy ?? 0);
      }
      viewer.scroll(dx ?? 0, dy ?? 0);
      return null;
    },
    async resetScroll() {
      viewer.set_scroll(0, 0);
    },
    async destroy() {
      // Allow GC to reclaim the viewer; no explicit teardown needed.
    },
  };
}

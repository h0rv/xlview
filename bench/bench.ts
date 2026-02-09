import { providers } from "./providers/index.js";
import type {
  BenchProvider,
  BenchProviderEntry,
  ManifestEntry,
} from "../src/demo/types.js";

const providerSelect = document.getElementById(
  "provider-select",
) as HTMLSelectElement;
const datasetSelect = document.getElementById(
  "dataset-select",
) as HTMLSelectElement;
const iterationsInput = document.getElementById(
  "iterations",
) as HTMLInputElement;
const warmupInput = document.getElementById("warmup") as HTMLInputElement;
const runBtn = document.getElementById("run-btn") as HTMLButtonElement;
const output = document.getElementById("output") as HTMLPreElement;
const statusEl = document.getElementById("status") as HTMLSpanElement;
const envEl = document.getElementById("env") as HTMLDivElement;
const canvas = document.getElementById("viewer-canvas") as HTMLCanvasElement;
const viewerContainer = document.getElementById(
  "viewer-container",
) as HTMLDivElement;
const domRoot = document.getElementById("viewer-dom") as HTMLDivElement;

const manifest: ManifestEntry[] = await fetch("./manifest.json").then((res) =>
  res.json(),
);

interface Stats {
  min: number;
  max: number;
  mean: number;
  median: number;
  stdev: number;
  samples: number;
}

function setStatus(message: string): void {
  statusEl.textContent = message;
}

function stats(values: number[]): Stats | null {
  if (!values.length) {
    return null;
  }
  const sorted = [...values].sort((a, b) => a - b);
  const sum = values.reduce((acc, v) => acc + v, 0);
  const mean = sum / values.length;
  const median = sorted[Math.floor(sorted.length / 2)]!;
  const min = sorted[0]!;
  const max = sorted[sorted.length - 1]!;
  const variance =
    values.reduce((acc, v) => acc + (v - mean) ** 2, 0) / values.length;
  const stdev = Math.sqrt(variance);
  return {
    min,
    max,
    mean,
    median,
    stdev,
    samples: values.length,
  };
}

function resizeCanvas(): { width: number; height: number; dpr: number } {
  const rect = viewerContainer.getBoundingClientRect();
  const width = viewerContainer.clientWidth || rect.width;
  const height = viewerContainer.clientHeight || rect.height;
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.floor(width * dpr));
  canvas.height = Math.max(1, Math.floor(height * dpr));
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  return { width: canvas.width, height: canvas.height, dpr };
}

async function timeAsync<T>(
  fn: () => Promise<T>,
): Promise<{ ms: number; result: T }> {
  const start = performance.now();
  const result = await fn();
  const end = performance.now();
  return { ms: end - start, result };
}

interface EnvInfo {
  userAgent: string;
  platform: string;
  deviceMemory: number | null;
  hardwareConcurrency: number | null;
  dpr: number;
  viewport: { width: number; height: number };
}

function getEnv(): EnvInfo {
  return {
    userAgent: navigator.userAgent,
    platform: navigator.platform,
    deviceMemory:
      ((navigator as unknown as Record<string, unknown>).deviceMemory as
        | number
        | null) ?? null,
    hardwareConcurrency: navigator.hardwareConcurrency ?? null,
    dpr: window.devicePixelRatio || 1,
    viewport: {
      width: window.innerWidth,
      height: window.innerHeight,
    },
  };
}

async function measureMemory(): Promise<unknown> {
  const perf = performance as unknown as Record<string, unknown>;
  if (typeof perf.measureUserAgentSpecificMemory === "function") {
    try {
      const memory = await (
        perf.measureUserAgentSpecificMemory as () => Promise<unknown>
      )();
      return memory;
    } catch {
      return null;
    }
  }
  const mem = perf.memory as
    | { jsHeapSizeLimit: number; usedJSHeapSize: number }
    | undefined;
  if (mem) {
    return {
      jsHeapSizeLimit: mem.jsHeapSizeLimit,
      usedJSHeapSize: mem.usedJSHeapSize,
    };
  }
  return null;
}

interface BenchOptions {
  providerId?: string;
  iterations?: number;
  warmup?: number;
  datasetId?: string;
  scrollSteps?: number;
  scrollStepPx?: number;
}

async function runBench(options: BenchOptions = {}): Promise<unknown> {
  const providerId = options.providerId ?? providerSelect.value;
  const iterations = Number(options.iterations ?? iterationsInput.value);
  const warmup = Number(options.warmup ?? warmupInput.value);
  const datasetFilter = options.datasetId ?? datasetSelect.value;
  const scrollSteps = Number(options.scrollSteps ?? 8);
  const scrollStepPx = Number(options.scrollStepPx ?? 160);

  const providerEntry = (providers as Record<string, BenchProviderEntry>)[
    providerId
  ];
  if (!providerEntry) {
    throw new Error(`Unknown provider: ${providerId}`);
  }

  const env = getEnv();
  envEl.textContent = `UA: ${env.userAgent}`;

  setStatus("Preparing...");

  const datasets =
    datasetFilter === "all"
      ? manifest
      : manifest.filter((entry) => entry.id === datasetFilter);

  const results: unknown[] = [];

  for (const entry of datasets) {
    const canvasMetrics = resizeCanvas();
    const provider: BenchProvider = await providerEntry.create({
      canvas,
      container: viewerContainer,
      domRoot,
      dpr: canvasMetrics.dpr,
      width: canvasMetrics.width,
      height: canvasMetrics.height,
    });

    try {
      setStatus(`Loading ${entry.id}...`);
      const buffer = await fetch(entry.file).then((res) => res.arrayBuffer());
      const bytes = new Uint8Array(buffer);

      // Warmup
      for (let i = 0; i < warmup; i += 1) {
        await provider.parse(bytes);
        await provider.load(bytes);
        await provider.render();
      }

      const parseTimes: number[] = [];
      const loadTimes: number[] = [];
      const renderTimes: number[] = [];
      const e2eTimes: number[] = [];
      const scrollTimes: number[] = [];
      const scrollStallRatios: number[] = [];
      const scrollStallYs: number[] = [];
      const scrollMaxYs: number[] = [];
      const scrollEndYs: number[] = [];
      const scrollZeroSteps: number[] = [];
      const parseInternalTimes: number[] = [];
      const parseDetailTimes: Record<string, number[]> = {
        relationships_ms: [],
        theme_ms: [],
        shared_strings_ms: [],
        styles_ms: [],
        workbook_info_ms: [],
        sheets_ms: [],
        charts_resolve_ms: [],
        images_ms: [],
        dxf_ms: [],
        style_resolve_ms: [],
        format_number_ms: [],
        format_number_date_ms: [],
        format_number_number_ms: [],
        value_parse_ms: [],
        text_unescape_ms: [],
      };
      const parseDetailCounts: Record<string, number[]> = {
        total_cells: [],
        total_rows: [],
        total_cols: [],
        total_merges: [],
        total_styles: [],
        total_default_styles: [],
        total_style_cache_hits: [],
        total_style_cache_misses: [],
        total_string_cells: [],
        total_number_cells: [],
        total_bool_cells: [],
        total_error_cells: [],
        total_date_cells: [],
        total_shared_string_cells: [],
        total_inline_string_cells: [],
        total_numfmt_builtin: [],
        total_numfmt_custom: [],
        total_numfmt_general: [],
        total_format_number_calls: [],
        total_format_number_date_calls: [],
        total_format_number_number_calls: [],
        total_value_parse_calls: [],
        total_text_unescape_calls: [],
        total_comments: [],
        total_hyperlinks: [],
        total_data_validations: [],
        total_conditional_formats: [],
        total_drawings: [],
        total_charts: [],
        shared_strings_count: [],
        shared_strings_chars: [],
        sheets_count: [],
        styles_fonts: [],
        styles_fills: [],
        styles_borders: [],
        styles_cell_xfs: [],
        styles_cell_style_xfs: [],
        styles_num_fmts: [],
        styles_named_styles: [],
        styles_dxf: [],
        styles_indexed_colors: [],
      };
      const loadParseTimes: number[] = [];
      const loadLayoutTimes: number[] = [];
      const loadTotalTimes: number[] = [];
      const renderPrepTimes: number[] = [];
      const renderDrawTimes: number[] = [];
      const renderTotalTimes: number[] = [];
      const renderVisibleCells: number[] = [];

      setStatus(`Benchmarking ${entry.id}...`);

      for (let i = 0; i < iterations; i += 1) {
        const parse = await timeAsync(() => provider.parse(bytes));
        parseTimes.push(parse.ms);
        const pr = parse.result as Record<string, unknown> | null;
        if (pr && typeof pr.parse_ms === "number") {
          parseInternalTimes.push(pr.parse_ms);
          for (const key of Object.keys(parseDetailTimes)) {
            if (typeof pr[key] === "number") {
              parseDetailTimes[key]!.push(pr[key]);
            }
          }
          for (const key of Object.keys(parseDetailCounts)) {
            if (typeof pr[key] === "number") {
              parseDetailCounts[key]!.push(pr[key]);
            }
          }
        }
      }

      for (let i = 0; i < iterations; i += 1) {
        const load = await timeAsync(() => provider.load(bytes));
        loadTimes.push(load.ms);
        const lr = load.result as Record<string, unknown> | null;
        if (lr && typeof lr.parse_ms === "number") {
          loadParseTimes.push(lr.parse_ms);
        }
        if (lr && typeof lr.layout_ms === "number") {
          loadLayoutTimes.push(lr.layout_ms);
        }
        if (lr && typeof lr.total_ms === "number") {
          loadTotalTimes.push(lr.total_ms);
        }

        const render = await timeAsync(() => provider.render());
        renderTimes.push(render.ms);
        const rr = render.result as Record<string, unknown> | null;
        if (rr) {
          if (typeof rr.prep_ms === "number") {
            renderPrepTimes.push(rr.prep_ms);
          }
          if (typeof rr.draw_ms === "number") {
            renderDrawTimes.push(rr.draw_ms);
          }
          if (typeof rr.total_ms === "number") {
            renderTotalTimes.push(rr.total_ms);
          }
          if (typeof rr.visible_cells === "number") {
            renderVisibleCells.push(rr.visible_cells);
          }
        }
      }

      for (let i = 0; i < iterations; i += 1) {
        const e2e = await timeAsync(async () => {
          await provider.load(bytes);
          await provider.render();
        });
        e2eTimes.push(e2e.ms);
      }

      if (provider.scroll) {
        await provider.load(bytes);
        await provider.render();
        for (let i = 0; i < iterations; i += 1) {
          if (provider.resetScroll) {
            await provider.resetScroll();
          }
          const scroll = await timeAsync(async () => {
            let stallRatio: number | null = null;
            let stallY: number | null = null;
            let maxY: number | null = null;
            let endY: number | null = null;
            let zeroSteps = 0;

            for (let step = 0; step < scrollSteps; step += 1) {
              const result = (await provider.scroll!({
                dx: 0,
                dy: scrollStepPx,
                step,
                totalSteps: scrollSteps,
              })) as Record<string, unknown> | null;
              if (result && typeof result.applied_dy === "number") {
                if (result.applied_dy === 0) {
                  zeroSteps += 1;
                  if (
                    stallRatio === null &&
                    typeof result.max_y === "number" &&
                    result.max_y > 0
                  ) {
                    stallRatio = (result.scroll_y as number) / result.max_y;
                    stallY = result.scroll_y as number;
                    maxY = result.max_y;
                  }
                }
                if (typeof result.scroll_y === "number") {
                  endY = result.scroll_y;
                }
                if (typeof result.max_y === "number") {
                  maxY = result.max_y;
                }
              }
              if (provider.render) {
                await provider.render();
              }
            }

            return { stallRatio, stallY, maxY, endY, zeroSteps };
          });
          scrollTimes.push(scroll.ms);
          if (scroll.result) {
            if (typeof scroll.result.stallRatio === "number") {
              scrollStallRatios.push(scroll.result.stallRatio);
            }
            if (typeof scroll.result.stallY === "number") {
              scrollStallYs.push(scroll.result.stallY);
            }
            if (typeof scroll.result.maxY === "number") {
              scrollMaxYs.push(scroll.result.maxY);
            }
            if (typeof scroll.result.endY === "number") {
              scrollEndYs.push(scroll.result.endY);
            }
            if (typeof scroll.result.zeroSteps === "number") {
              scrollZeroSteps.push(scroll.result.zeroSteps);
            }
          }
        }
      }

      results.push({
        dataset: entry,
        metrics: {
          parse_ms: parseTimes,
          parse_internal_ms: parseInternalTimes,
          parse_details_ms: parseDetailTimes,
          parse_details_counts: parseDetailCounts,
          load_ms: loadTimes,
          load_parse_ms: loadParseTimes,
          load_layout_ms: loadLayoutTimes,
          load_total_ms: loadTotalTimes,
          render_ms: renderTimes,
          render_prep_ms: renderPrepTimes,
          render_draw_ms: renderDrawTimes,
          render_total_ms: renderTotalTimes,
          render_visible_cells: renderVisibleCells,
          e2e_ms: e2eTimes,
          scroll_ms: scrollTimes,
          scroll_stall_ratio: scrollStallRatios,
          scroll_stall_y: scrollStallYs,
          scroll_max_y: scrollMaxYs,
          scroll_end_y: scrollEndYs,
          scroll_zero_steps: scrollZeroSteps,
        },
        stats: {
          parse: stats(parseTimes),
          parse_internal: stats(parseInternalTimes),
          parse_details: Object.fromEntries(
            Object.entries(parseDetailTimes).map(([key, values]) => [
              key,
              stats(values),
            ]),
          ),
          parse_details_counts: Object.fromEntries(
            Object.entries(parseDetailCounts).map(([key, values]) => [
              key,
              stats(values),
            ]),
          ),
          load: stats(loadTimes),
          load_parse: stats(loadParseTimes),
          load_layout: stats(loadLayoutTimes),
          load_total: stats(loadTotalTimes),
          render: stats(renderTimes),
          render_prep: stats(renderPrepTimes),
          render_draw: stats(renderDrawTimes),
          render_total: stats(renderTotalTimes),
          render_visible_cells: stats(renderVisibleCells),
          e2e: stats(e2eTimes),
          scroll: stats(scrollTimes),
          scroll_stall_ratio: stats(scrollStallRatios),
          scroll_stall_y: stats(scrollStallYs),
          scroll_max_y: stats(scrollMaxYs),
          scroll_end_y: stats(scrollEndYs),
          scroll_zero_steps: stats(scrollZeroSteps),
        },
      });
    } finally {
      if (provider.destroy) {
        await provider.destroy();
      }
    }
  }

  const memory = await measureMemory();

  const payload = {
    meta: {
      provider: providerId,
      timestamp: new Date().toISOString(),
      env,
      memory,
      iterations,
      warmup,
      scroll: {
        steps: scrollSteps,
        stepPx: scrollStepPx,
      },
    },
    results,
  };

  setStatus("Done");
  output.textContent = JSON.stringify(payload, null, 2);
  return payload;
}

function populateOptions(): void {
  Object.values(providers).forEach((provider) => {
    const option = document.createElement("option");
    option.value = provider.id;
    option.textContent = provider.label;
    providerSelect.appendChild(option);
  });

  manifest.forEach((entry) => {
    const option = document.createElement("option");
    option.value = entry.id;
    option.textContent = `${entry.label} (${entry.size})`;
    datasetSelect.appendChild(option);
  });
}

populateOptions();

window.addEventListener("resize", () => {
  resizeCanvas();
});

runBtn.addEventListener("click", async () => {
  runBtn.disabled = true;
  try {
    const payload = await runBench();
    if (new URLSearchParams(window.location.search).has("debug")) {
      console.log("Bench results", payload);
    }
  } catch (error) {
    output.textContent = `Error: ${(error as Error).message}`;
  } finally {
    runBtn.disabled = false;
  }
});

(window as unknown as Record<string, unknown>).runBench = runBench;

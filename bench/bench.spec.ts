import { test } from "@playwright/test";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

test.setTimeout(10 * 60 * 1000);

const iterations = Number(process.env.BENCH_ITERATIONS ?? "7");
const warmup = Number(process.env.BENCH_WARMUP ?? "2");
const datasetId = process.env.BENCH_DATASET ?? "all";
const providerId = process.env.BENCH_PROVIDER ?? "xlview";
const scrollSteps = Number(process.env.BENCH_SCROLL_STEPS ?? "8");
const scrollStepPx = Number(process.env.BENCH_SCROLL_STEP_PX ?? "160");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const resultsDir = path.resolve(__dirname, "results");

test("xlview browser bench", async ({ page }) => {
  const consoleMessages: Array<{
    type: string;
    text: string;
    location?: string | null;
  }> = [];
  page.on("console", (msg) => {
    const location = msg.location();
    const locationText = location.url
      ? `${location.url}:${location.lineNumber}`
      : null;
    consoleMessages.push({
      type: msg.type(),
      text: msg.text(),
      location: locationText,
    });
  });
  page.on("pageerror", (error) => {
    consoleMessages.push({
      type: "pageerror",
      text: error.message,
      location: error.stack ?? null,
    });
  });

  await page.goto("/bench/bench.html");
  await page.waitForFunction(() => typeof window.runBench === "function");

  const payload = await page.evaluate(
    async (opts) => {
      return await window.runBench(opts);
    },
    { providerId, iterations, warmup, datasetId, scrollSteps, scrollStepPx },
  );

  payload.meta = payload.meta ?? {};
  payload.meta.browserLogs = consoleMessages;

  await fs.mkdir(resultsDir, { recursive: true });
  const filename = `bench-${providerId}-${Date.now()}.json`;
  const outputPath = path.join(resultsDir, filename);
  await fs.writeFile(outputPath, JSON.stringify(payload, null, 2));

  console.log(`Saved benchmark results to ${outputPath}`);
});

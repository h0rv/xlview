import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const resultsDir = path.join(__dirname, 'results');

const args = process.argv.slice(2);
const providerFilter = getArgValue('--provider');
const datasetFilter = getArgValue('--dataset');
const lastCount = Number(getArgValue('--last') ?? '5');
const compare = args.includes('--compare') || args.includes('--diff');
const csv = args.includes('--csv');
const csvOut = getArgValue('--csv-out');

function getArgValue(flag) {
  const idx = args.indexOf(flag);
  if (idx === -1) return null;
  return args[idx + 1] ?? null;
}

function formatMs(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return '-';
  return value.toFixed(2);
}

function formatPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return '-';
  return `${value.toFixed(1)}%`;
}

function pad(text, width) {
  const str = String(text);
  if (str.length >= width) return str;
  return str + ' '.repeat(width - str.length);
}

function buildRunSummary(run) {
  const meta = run.meta || {};
  const resultCount = Array.isArray(run.results) ? run.results.length : 0;
  return {
    provider: meta.provider ?? 'unknown',
    timestamp: meta.timestamp ?? 'unknown',
    iterations: meta.iterations ?? '-',
    warmup: meta.warmup ?? '-',
    datasetCount: resultCount,
    userAgent: meta.env?.userAgent ?? 'unknown',
  };
}

function buildDatasetStats(run) {
  const byId = new Map();
  for (const entry of run.results ?? []) {
    const id = entry.dataset?.id ?? 'unknown';
    if (datasetFilter && id !== datasetFilter) continue;
    const stats = entry.stats ?? {};
    byId.set(id, {
      label: entry.dataset?.label ?? id,
      parse: stats.parse?.mean ?? null,
      parse_internal: stats.parse_internal?.mean ?? null,
      parse_details: stats.parse_details ?? null,
      load: stats.load?.mean ?? null,
      load_parse: stats.load_parse?.mean ?? null,
      load_layout: stats.load_layout?.mean ?? null,
      load_total: stats.load_total?.mean ?? null,
      render: stats.render?.mean ?? null,
      render_prep: stats.render_prep?.mean ?? null,
      render_draw: stats.render_draw?.mean ?? null,
      render_total: stats.render_total?.mean ?? null,
      render_visible_cells: stats.render_visible_cells?.mean ?? null,
      e2e: stats.e2e?.mean ?? null,
      scroll: stats.scroll?.mean ?? null,
      size: entry.dataset?.size ?? null,
    });
  }
  return byId;
}

async function loadRuns() {
  let files = [];
  try {
    files = await fs.readdir(resultsDir);
  } catch {
    return [];
  }

  const runs = [];
  for (const file of files) {
    if (!file.endsWith('.json')) continue;
    const filePath = path.join(resultsDir, file);
    const content = await fs.readFile(filePath, 'utf8');
    try {
      const data = JSON.parse(content);
      if (providerFilter && data.meta?.provider !== providerFilter) {
        continue;
      }
      runs.push({ file, data });
    } catch {
      // skip invalid
    }
  }

  runs.sort((a, b) => {
    const ta = Date.parse(a.data.meta?.timestamp ?? '') || 0;
    const tb = Date.parse(b.data.meta?.timestamp ?? '') || 0;
    return tb - ta;
  });
  return runs;
}

function printRuns(runs) {
  if (!runs.length) {
    console.log('No benchmark results found.');
    return;
  }

  console.log('Runs (newest first)');
  const header = [
    pad('file', 32),
    pad('provider', 10),
    pad('timestamp', 24),
    pad('datasets', 9),
    pad('iters', 7),
    pad('warmup', 7),
  ].join(' ');
  console.log(header);
  console.log('-'.repeat(header.length));

  for (const run of runs.slice(0, lastCount)) {
    const summary = buildRunSummary(run.data);
    console.log([
      pad(run.file, 32),
      pad(summary.provider, 10),
      pad(summary.timestamp, 24),
      pad(summary.datasetCount, 9),
      pad(summary.iterations, 7),
      pad(summary.warmup, 7),
    ].join(' '));
  }
}

function printCompare(current, previous) {
  if (!current || !previous) {
    console.log('\nNot enough runs to compare.');
    return;
  }

  const currentStats = buildDatasetStats(current.data);
  const prevStats = buildDatasetStats(previous.data);
  const datasetIds = Array.from(new Set([...currentStats.keys(), ...prevStats.keys()]));

  if (!datasetIds.length) {
    console.log('\nNo datasets to compare.');
    return;
  }

  console.log('\nComparison (mean ms)');
  console.log(`Current: ${current.file}`);
  console.log(`Previous: ${previous.file}`);

  const header = [
    pad('dataset', 20),
    pad('parse', 8),
    pad('load', 8),
    pad('render', 8),
    pad('e2e', 8),
    pad('scroll', 8),
    pad('e2eΔ', 8),
    pad('e2eΔ%', 8),
  ].join(' ');
  console.log(header);
  console.log('-'.repeat(header.length));

  for (const id of datasetIds) {
    const cur = currentStats.get(id);
    const prev = prevStats.get(id);

    const curE2e = cur?.e2e ?? null;
    const prevE2e = prev?.e2e ?? null;
    const delta = (curE2e !== null && prevE2e !== null) ? curE2e - prevE2e : null;
    const deltaPct = (curE2e !== null && prevE2e) ? (delta / prevE2e) * 100 : null;

    console.log([
      pad(id, 20),
      pad(formatMs(cur?.parse), 8),
      pad(formatMs(cur?.load), 8),
      pad(formatMs(cur?.render), 8),
      pad(formatMs(cur?.e2e), 8),
      pad(formatMs(cur?.scroll), 8),
      pad(delta !== null ? formatMs(delta) : '-', 8),
      pad(deltaPct !== null ? formatPct(deltaPct) : '-', 8),
    ].join(' '));
  }
}

async function writeCsv(runs) {
  if (!runs.length) return;
  const rows = [];
  rows.push([
    'file',
    'provider',
    'timestamp',
    'dataset',
    'dataset_label',
    'dataset_size',
    'parse_mean_ms',
    'parse_internal_mean_ms',
    'parse_relationships_mean_ms',
    'parse_theme_mean_ms',
    'parse_shared_strings_mean_ms',
    'parse_styles_mean_ms',
    'parse_workbook_info_mean_ms',
    'parse_sheets_mean_ms',
    'parse_charts_resolve_mean_ms',
    'parse_images_mean_ms',
    'parse_dxf_mean_ms',
    'parse_value_parse_mean_ms',
    'parse_text_unescape_mean_ms',
    'load_mean_ms',
    'load_parse_mean_ms',
    'load_layout_mean_ms',
    'load_total_mean_ms',
    'render_mean_ms',
    'render_prep_mean_ms',
    'render_draw_mean_ms',
    'render_total_mean_ms',
    'render_visible_cells_mean',
    'e2e_mean_ms',
    'scroll_mean_ms',
    'iterations',
    'warmup',
  ].join(','));

  for (const run of runs) {
    const summary = buildRunSummary(run.data);
    const statsById = buildDatasetStats(run.data);
    for (const [id, stats] of statsById.entries()) {
      rows.push([
        run.file,
        summary.provider,
        summary.timestamp,
        id,
        stats.label ?? '',
        stats.size ?? '',
        formatMs(stats.parse),
        formatMs(stats.parse_internal),
        formatMs(stats.parse_details?.relationships_ms?.mean ?? null),
        formatMs(stats.parse_details?.theme_ms?.mean ?? null),
        formatMs(stats.parse_details?.shared_strings_ms?.mean ?? null),
        formatMs(stats.parse_details?.styles_ms?.mean ?? null),
        formatMs(stats.parse_details?.workbook_info_ms?.mean ?? null),
        formatMs(stats.parse_details?.sheets_ms?.mean ?? null),
        formatMs(stats.parse_details?.charts_resolve_ms?.mean ?? null),
        formatMs(stats.parse_details?.images_ms?.mean ?? null),
        formatMs(stats.parse_details?.dxf_ms?.mean ?? null),
        formatMs(stats.parse_details?.value_parse_ms?.mean ?? null),
        formatMs(stats.parse_details?.text_unescape_ms?.mean ?? null),
        formatMs(stats.load),
        formatMs(stats.load_parse),
        formatMs(stats.load_layout),
        formatMs(stats.load_total),
        formatMs(stats.render),
        formatMs(stats.render_prep),
        formatMs(stats.render_draw),
        formatMs(stats.render_total),
        formatMs(stats.render_visible_cells),
        formatMs(stats.e2e),
        formatMs(stats.scroll),
        summary.iterations,
        summary.warmup,
      ].join(','));
    }
  }

  const csvPath = csvOut ? path.resolve(process.cwd(), csvOut) : path.join(resultsDir, 'summary.csv');
  await fs.writeFile(csvPath, rows.join('\n'), 'utf8');
  console.log(`\nWrote CSV summary to ${csvPath}`);
}

const runs = await loadRuns();
printRuns(runs);

if (compare) {
  printCompare(runs[0], runs[1]);
}

if (csv) {
  await writeCsv(runs);
}

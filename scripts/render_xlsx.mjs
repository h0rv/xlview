#!/usr/bin/env node
/**
 * CLI tool to render XLSX files to standalone HTML for testing.
 *
 * Usage:
 *   node scripts/render_xlsx.mjs test/kitchen_sink.xlsx
 *   node scripts/render_xlsx.mjs test/kitchen_sink.xlsx -o output.html
 *   node scripts/render_xlsx.mjs test/kitchen_sink.xlsx --json  # Output raw JSON
 *   node scripts/render_xlsx.mjs test/kitchen_sink.xlsx --open  # Open in browser
 */

import { readFile, writeFile } from 'fs/promises';
import { fileURLToPath } from 'url';
import { dirname, join, basename } from 'path';
import { execSync } from 'child_process';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');

// Parse args
const args = process.argv.slice(2);
const inputFile = args.find(a => !a.startsWith('-'));
const outputFile = args.includes('-o') ? args[args.indexOf('-o') + 1] : null;
const jsonOnly = args.includes('--json');
const openBrowser = args.includes('--open');
const quiet = args.includes('-q') || args.includes('--quiet');

if (!inputFile) {
  console.error('Usage: node scripts/render_xlsx.mjs <input.xlsx> [-o output.html] [--json] [--open]');
  process.exit(1);
}

// Dynamic import of WASM
async function loadWasm() {
  const wasmPath = join(ROOT, 'pkg', 'xlview.js');
  try {
    const wasm = await import(wasmPath);
    await wasm.default();
    return wasm;
  } catch (e) {
    console.error('Failed to load WASM. Run: cargo build --target wasm32-unknown-unknown && wasm-bindgen ...');
    console.error(e.message);
    process.exit(1);
  }
}

// Import renderer
async function loadRenderer() {
  const rendererPath = join(ROOT, 'renderer.js');
  try {
    return await import(rendererPath);
  } catch (e) {
    console.error('Failed to load renderer.js:', e.message);
    process.exit(1);
  }
}

// Generate standalone HTML
function generateStandaloneHtml(workbook, renderer) {
  const styles = renderer.defaultStyles || '';

  // Create a mock DOM environment for the renderer
  const sheets = workbook.sheets || [];

  // Generate HTML for each sheet
  let sheetsHtml = '';
  let tabsHtml = '';

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const isActive = i === 0;
    const tabColor = sheet.tabColor ? `border-top: 3px solid ${sheet.tabColor};` : '';

    tabsHtml += `<button class="xlsx-tab${isActive ? ' active' : ''}" data-sheet="${i}" style="${tabColor}">${escapeHtml(sheet.name)}</button>`;
    sheetsHtml += generateSheetHtml(sheet, isActive);
  }

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>XLSX Render: ${basename(inputFile)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f5f5f5; }
    .xlsx-viewer { background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden; }
    ${styles}
    .xlsx-table td { border: 1px solid #d0d0d0; }
    .info { padding: 10px; background: #e8f4fc; border-bottom: 1px solid #ccc; font-size: 12px; }
  </style>
</head>
<body>
  <div class="info">
    <strong>File:</strong> ${escapeHtml(inputFile)} |
    <strong>Sheets:</strong> ${sheets.length} |
    <strong>Rendered:</strong> ${new Date().toISOString()}
  </div>
  <div class="xlsx-viewer">
    <div class="xlsx-tabs">${tabsHtml}</div>
    <div class="xlsx-sheets">${sheetsHtml}</div>
  </div>
  <script>
    document.querySelectorAll('.xlsx-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.xlsx-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.xlsx-sheet-container').forEach(s => s.style.display = 'none');
        tab.classList.add('active');
        document.querySelector('.xlsx-sheet-container[data-sheet="' + tab.dataset.sheet + '"]').style.display = 'block';
      });
    });
  </script>
</body>
</html>`;
}

function generateSheetHtml(sheet, isActive) {
  const display = isActive ? 'block' : 'none';

  // Build cell map
  const cellMap = new Map();
  for (const cell of (sheet.cells || [])) {
    cellMap.set(`${cell.row},${cell.col}`, cell);
  }

  // Build merge map
  const mergeMap = new Map();
  const mergeSkip = new Set();
  for (const merge of (sheet.merges || [])) {
    mergeMap.set(`${merge.startRow},${merge.startCol}`, merge);
    for (let r = merge.startRow; r <= merge.endRow; r++) {
      for (let c = merge.startCol; c <= merge.endCol; c++) {
        if (r !== merge.startRow || c !== merge.startCol) {
          mergeSkip.add(`${r},${c}`);
        }
      }
    }
  }

  // Column widths
  const colWidths = new Map();
  for (const col of (sheet.columns || [])) {
    colWidths.set(col.index, col.width);
  }

  // Row heights
  const rowHeights = new Map();
  for (const row of (sheet.rows || [])) {
    rowHeights.set(row.index, row.height);
  }

  // Generate table
  let html = `<div class="xlsx-sheet-container" data-sheet="${sheet.name}" style="display: ${display}; overflow: auto; max-height: 80vh;">`;
  html += '<div class="xlsx-sheet-wrapper"><table class="xlsx-table">';

  // Colgroup
  html += '<colgroup>';
  for (let c = 0; c < sheet.maxCol; c++) {
    const width = colWidths.get(c) ?? sheet.defaultColWidth ?? 64;
    html += `<col style="width: ${width}px;">`;
  }
  html += '</colgroup>';

  // Rows
  html += '<tbody>';
  for (let r = 0; r < sheet.maxRow; r++) {
    const height = rowHeights.get(r) ?? sheet.defaultRowHeight ?? 20;
    html += `<tr style="height: ${height}px;">`;

    for (let c = 0; c < sheet.maxCol; c++) {
      const key = `${r},${c}`;

      if (mergeSkip.has(key)) continue;

      const cell = cellMap.get(key);
      const merge = mergeMap.get(key);

      let attrs = '';
      if (merge) {
        const rowSpan = merge.endRow - merge.startRow + 1;
        const colSpan = merge.endCol - merge.startCol + 1;
        if (rowSpan > 1) attrs += ` rowspan="${rowSpan}"`;
        if (colSpan > 1) attrs += ` colspan="${colSpan}"`;
      }

      const style = cell?.style ? generateCellStyle(cell.style) : '';
      const value = cell?.formatted ?? cell?.value ?? '';

      html += `<td${attrs}${style ? ` style="${style}"` : ''}>${formatCellValue(cell, value)}</td>`;
    }

    html += '</tr>';
  }
  html += '</tbody></table></div></div>';

  return html;
}

function generateCellStyle(style) {
  const parts = [];

  if (style.fontName) parts.push(`font-family: "${style.fontName}"`);
  if (style.fontSize) parts.push(`font-size: ${style.fontSize}pt`);
  if (style.fontBold) parts.push('font-weight: bold');
  if (style.fontItalic) parts.push('font-style: italic');
  if (style.fontColor) parts.push(`color: ${style.fontColor}`);
  if (style.bgColor) parts.push(`background-color: ${style.bgColor}`);

  if (style.alignH) {
    const map = { left: 'left', center: 'center', right: 'right', justify: 'justify' };
    if (map[style.alignH]) parts.push(`text-align: ${map[style.alignH]}`);
  }

  if (style.alignV) {
    const map = { top: 'top', center: 'middle', middle: 'middle', bottom: 'bottom' };
    if (map[style.alignV]) parts.push(`vertical-align: ${map[style.alignV]}`);
  }

  if (style.borderTop?.style && style.borderTop.style !== 'none') {
    parts.push(`border-top: ${getBorderCss(style.borderTop)}`);
  }
  if (style.borderRight?.style && style.borderRight.style !== 'none') {
    parts.push(`border-right: ${getBorderCss(style.borderRight)}`);
  }
  if (style.borderBottom?.style && style.borderBottom.style !== 'none') {
    parts.push(`border-bottom: ${getBorderCss(style.borderBottom)}`);
  }
  if (style.borderLeft?.style && style.borderLeft.style !== 'none') {
    parts.push(`border-left: ${getBorderCss(style.borderLeft)}`);
  }

  if (style.underline) parts.push('text-decoration: underline');
  if (style.strikethrough) parts.push('text-decoration: line-through');
  if (style.wrapText) parts.push('white-space: normal; word-wrap: break-word');
  if (style.textRotation) parts.push(`transform: rotate(-${style.textRotation}deg)`);

  return parts.join('; ');
}

function getBorderCss(border) {
  const widthMap = { thin: '1px', medium: '2px', thick: '3px', double: '3px', dotted: '1px', dashed: '1px' };
  const styleMap = { thin: 'solid', medium: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed' };

  const width = widthMap[border.style] || '1px';
  const style = styleMap[border.style] || 'solid';
  const color = border.color || '#000000';

  return `${width} ${style} ${color}`;
}

function formatCellValue(cell, value) {
  if (!cell) return '';

  // Handle hyperlinks
  if (cell.hyperlink) {
    return `<a href="${escapeHtml(cell.hyperlink)}" target="_blank">${escapeHtml(String(value))}</a>`;
  }

  // Handle rich text
  if (cell.richText && Array.isArray(cell.richText)) {
    return cell.richText.map(run => {
      let text = escapeHtml(run.text || '');
      if (run.style) {
        const styles = [];
        if (run.style.bold) styles.push('font-weight: bold');
        if (run.style.italic) styles.push('font-style: italic');
        if (run.style.color) styles.push(`color: ${run.style.color}`);
        if (styles.length) {
          text = `<span style="${styles.join('; ')}">${text}</span>`;
        }
      }
      return text;
    }).join('');
  }

  return escapeHtml(String(value));
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Main
async function main() {
  if (!quiet) console.error(`Loading ${inputFile}...`);

  // Read input file
  const xlsxData = await readFile(inputFile);

  // Load WASM and parse
  const wasm = await loadWasm();
  const workbook = wasm.parse_xlsx_to_js(new Uint8Array(xlsxData));

  if (jsonOnly) {
    console.log(JSON.stringify(workbook, null, 2));
    return;
  }

  // Load renderer for styles
  const renderer = await loadRenderer();

  // Generate HTML
  const html = generateStandaloneHtml(workbook, renderer);

  // Output
  const outPath = outputFile || inputFile.replace(/\.xlsx$/i, '.html');
  await writeFile(outPath, html);
  if (!quiet) console.error(`Written: ${outPath}`);

  // Optionally open in browser
  if (openBrowser) {
    const cmd = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start' : 'xdg-open';
    execSync(`${cmd} "${outPath}"`);
  }

  // Print summary
  if (!quiet) {
    console.error(`\nSummary:`);
    console.error(`  Sheets: ${workbook.sheets?.length || 0}`);
    for (const sheet of (workbook.sheets || [])) {
      const styledCells = (sheet.cells || []).filter(c => c.style && Object.keys(c.style).length > 0).length;
      console.error(`    - ${sheet.name}: ${sheet.cells?.length || 0} cells, ${styledCells} styled, ${sheet.maxRow}x${sheet.maxCol}`);
    }
  }
}

main().catch(e => {
  console.error('Error:', e.message);
  process.exit(1);
});

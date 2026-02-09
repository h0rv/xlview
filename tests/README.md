# Visual Regression Tests

This directory contains Playwright-based visual regression tests for the xlview Excel viewer.

## Setup

1. Install dependencies:
```bash
npm install
```

2. Install Playwright browsers:
```bash
npx playwright install chromium
```

## Running Tests

### Run all tests
```bash
npm test
```

### Run tests in UI mode (interactive)
```bash
npm run test:ui
```

### Run tests in debug mode
```bash
npm run test:debug
```

### Update baseline screenshots
When you intentionally change the rendering and need to update the baseline images:
```bash
npm run test:update
```

## How It Works

The tests:
1. Start a local Python HTTP server automatically (configured in `playwright.config.ts`)
2. Navigate to `http://localhost:8080`
3. Wait for WASM to initialize
4. Load test Excel files using the built-in buttons
5. Take screenshots of the rendered output
6. Compare screenshots against baseline images stored in `tests/snapshots/`

## Test Files

- `visual.spec.ts` - Main visual regression test suite
  - Tests `minimal.xlsx` rendering
  - Tests `styled.xlsx` rendering
  - Tests multi-sheet tab functionality
  - Tests file upload functionality

## Baseline Screenshots

Baseline screenshots are stored in `tests/snapshots/` and are committed to the repository. When tests run, Playwright compares the current rendering against these baselines and fails if there are significant differences.

## Troubleshooting

### WASM initialization failures
If tests fail because WASM isn't loading, check:
- The `pkg/` directory contains the compiled WASM files
- The Python server is running on port 8080
- No other process is using port 8080

### Screenshot differences
Minor pixel differences are expected across platforms. The tests allow up to 100 pixels of difference. If tests fail:
1. Review the diff images in `tests/test-results/`
2. If changes are intentional, update baselines with `npm run test:update`
3. If unintentional, investigate rendering changes

### Browser differences
Tests currently only run in Chromium for simplicity. You can add Firefox/WebKit in `playwright.config.ts` if needed.

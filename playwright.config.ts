import { defineConfig, devices } from '@playwright/test';

/**
 * Minimal Playwright configuration for xlview visual regression testing
 */
export default defineConfig({
  testDir: './tests',

  // Folder for test artifacts such as screenshots, videos, traces, etc.
  outputDir: './tests/test-results',

  // Folder for baseline screenshots
  snapshotDir: './tests/snapshots',

  // Fail fast in CI
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,

  // Reporter to use
  reporter: [
    ['html', { outputFolder: './tests/playwright-report' }],
    ['list']
  ],

  use: {
    // Base URL for tests
    baseURL: 'http://localhost:8080',

    // Collect trace when retrying the failed test
    trace: 'on-first-retry',

    // Screenshot on failure
    screenshot: 'only-on-failure',
  },

  // Configure projects for major browsers
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],

  // Run local dev server before tests
  webServer: {
    command: 'python3 -m http.server 8080',
    url: 'http://localhost:8080',
    reuseExistingServer: !process.env.CI,
    timeout: 10000,
  },
});

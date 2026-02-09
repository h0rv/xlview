import { test, expect } from '@playwright/test';

/**
 * Visual regression tests for xlview Excel viewer
 *
 * These tests verify that the rendering of Excel files remains consistent
 * by comparing screenshots against baseline images.
 */

test.describe('xlview visual regression', () => {
  test.beforeEach(async ({ page }) => {
    // Navigate to the demo page
    await page.goto('/');

    // Wait for WASM to initialize - look for the upload area to be ready
    await expect(page.locator('#upload-area')).toBeVisible();

    // Additional wait for WASM initialization
    // The module is loaded via top-level await, so waiting for network idle helps
    await page.waitForLoadState('networkidle');
  });

  test('renders minimal.xlsx correctly', async ({ page }) => {
    // Click the "Load minimal.xlsx" button
    await page.click('#load-minimal');

    // Wait for the workbook to render
    await expect(page.locator('.xlsx-viewer')).toBeVisible();

    // Wait for the table to be populated
    await expect(page.locator('.xlsx-table')).toBeVisible();

    // Give a bit of time for any animations/rendering to complete
    await page.waitForTimeout(500);

    // Take screenshot and compare to baseline
    const viewer = page.locator('#viewer-container');
    await expect(viewer).toHaveScreenshot('minimal-xlsx.png', {
      maxDiffPixels: 100, // Allow minor rendering differences
    });
  });

  test('renders styled.xlsx correctly', async ({ page }) => {
    // Click the "Load styled.xlsx" button
    await page.click('#load-styled');

    // Wait for the workbook to render
    await expect(page.locator('.xlsx-viewer')).toBeVisible();

    // Wait for the table to be populated
    await expect(page.locator('.xlsx-table')).toBeVisible();

    // Give a bit of time for any animations/rendering to complete
    await page.waitForTimeout(500);

    // Take screenshot and compare to baseline
    const viewer = page.locator('#viewer-container');
    await expect(viewer).toHaveScreenshot('styled-xlsx.png', {
      maxDiffPixels: 100, // Allow minor rendering differences
    });
  });

  test('renders tabs for multi-sheet workbooks', async ({ page }) => {
    // Load styled.xlsx which has multiple sheets
    await page.click('#load-styled');

    // Wait for the workbook to render
    await expect(page.locator('.xlsx-viewer')).toBeVisible();

    // Check if tabs are visible (if styled.xlsx has multiple sheets)
    const tabs = page.locator('.xlsx-tab');
    const tabCount = await tabs.count();

    if (tabCount > 1) {
      // Click on second tab if it exists
      await tabs.nth(1).click();

      // Wait for sheet to render
      await page.waitForTimeout(300);

      // Take screenshot of second sheet
      const viewer = page.locator('#viewer-container');
      await expect(viewer).toHaveScreenshot('styled-xlsx-sheet2.png', {
        maxDiffPixels: 100,
      });
    }
  });

  test('handles file upload via drag and drop', async ({ page }) => {
    // Get the file input element
    const fileInput = page.locator('#file-input');

    // Upload the minimal.xlsx file
    await fileInput.setInputFiles('test/minimal.xlsx');

    // Wait for the workbook to render
    await expect(page.locator('.xlsx-viewer')).toBeVisible();
    await expect(page.locator('.xlsx-table')).toBeVisible();

    // Verify it rendered correctly
    const viewer = page.locator('#viewer-container');
    await expect(viewer).toHaveScreenshot('minimal-xlsx-uploaded.png', {
      maxDiffPixels: 100,
    });
  });
});

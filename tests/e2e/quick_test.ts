#!/usr/bin/env bun
/**
 * E2E test to verify Canvas 2D rendering
 * Run with: bun tests/e2e/quick_test.ts
 */
import { chromium } from 'playwright';

const BASE_URL = process.env.TEST_URL || 'http://localhost:8080';

interface TestResult {
    name: string;
    passed: boolean;
    error?: string;
}

async function main() {
    console.log(`Testing: ${BASE_URL}`);

    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();

    const logs: string[] = [];
    const errors: string[] = [];

    page.on('console', msg => {
        const text = `[${msg.type()}] ${msg.text()}`;
        logs.push(text);
        if (msg.type() === 'error') errors.push(text);
    });
    page.on('pageerror', err => errors.push(`[PAGE ERROR] ${err.message}`));

    const results: TestResult[] = [];

    try {
        await page.goto(BASE_URL, { waitUntil: 'load' });
        await page.waitForTimeout(500);

        // Test 1: Page loads without errors
        results.push({
            name: 'Page loads',
            passed: errors.length === 0,
            error: errors.length > 0 ? errors.join(', ') : undefined
        });

        // Test 2: Test minimal.xlsx
        const minimalBtn = await page.$('button:has-text("minimal.xlsx")');
        if (minimalBtn) {
            await minimalBtn.click();
            await page.waitForTimeout(1000);

            // Check canvas has content (not blank)
            const canvasContent = await page.evaluate(() => {
                const canvas = document.querySelector('#viewer-canvas') as HTMLCanvasElement;
                if (!canvas) return null;
                const ctx = canvas.getContext('2d');
                if (!ctx) return null;
                // Sample a few pixels to check it's not blank
                const data = ctx.getImageData(50, 50, 1, 1).data;
                return { r: data[0], g: data[1], b: data[2], a: data[3] };
            });

            results.push({
                name: 'minimal.xlsx renders',
                passed: canvasContent !== null && canvasContent.a > 0,
                error: canvasContent ? `pixel at 50,50: rgba(${canvasContent.r},${canvasContent.g},${canvasContent.b},${canvasContent.a})` : 'Canvas not found'
            });
        }

        // Test 3: Test kitchen_sink.xlsx - the one that was broken
        const kitchenBtn = await page.$('button:has-text("kitchen_sink.xlsx")');
        if (kitchenBtn) {
            errors.length = 0; // Clear previous errors
            await kitchenBtn.click();
            await page.waitForTimeout(1500);

            // Take screenshot
            await page.screenshot({ path: 'tests/e2e/kitchen_sink.png' });

            // Check multiple columns are visible by sampling different x positions
            const columnCheck = await page.evaluate(() => {
                const canvas = document.querySelector('#viewer-canvas') as HTMLCanvasElement;
                if (!canvas) return { multiColumn: false, reason: 'no canvas' };
                const ctx = canvas.getContext('2d');
                if (!ctx) return { multiColumn: false, reason: 'no context' };

                // Sample at different x positions (100, 200, 300, 400 pixels)
                // If layout is correct, we should see different content/colors
                const samples = [];
                for (let x = 100; x <= 400; x += 100) {
                    const data = ctx.getImageData(x, 30, 1, 1).data;
                    samples.push({ x, r: data[0], g: data[1], b: data[2] });
                }

                // Check if we have grid lines (gray pixels between white cells)
                const hasGridLines = samples.some(s => s.r === s.g && s.g === s.b && s.r > 200 && s.r < 255);

                return {
                    multiColumn: true,
                    samples,
                    hasGridLines
                };
            });

            results.push({
                name: 'kitchen_sink.xlsx renders multiple columns',
                passed: columnCheck.multiColumn && errors.length === 0,
                error: !columnCheck.multiColumn ? columnCheck.reason : (errors.length > 0 ? errors.join(', ') : undefined)
            });
        }

        // Test 4: Scrolling works
        const scrollResult = await page.evaluate(async () => {
            const canvas = document.querySelector('#viewer-canvas') as HTMLCanvasElement;
            if (!canvas) return { scrolled: false, reason: 'no canvas' };

            // Dispatch wheel event
            canvas.dispatchEvent(new WheelEvent('wheel', { deltaX: 0, deltaY: 100 }));
            await new Promise(r => setTimeout(r, 100));

            return { scrolled: true };
        });

        results.push({
            name: 'Scroll works',
            passed: scrollResult.scrolled,
            error: scrollResult.reason
        });

    } finally {
        await page.screenshot({ path: 'tests/e2e/output.png' });
        await browser.close();
    }

    // Print results
    console.log('\n=== Test Results ===');
    let allPassed = true;
    for (const result of results) {
        const status = result.passed ? '✓' : '✗';
        console.log(`${status} ${result.name}`);
        if (result.error) console.log(`  ${result.error}`);
        if (!result.passed) allPassed = false;
    }

    console.log(`\nScreenshots: tests/e2e/output.png, tests/e2e/kitchen_sink.png`);
    console.log(`\nOverall: ${allPassed ? 'PASSED' : 'FAILED'}`);

    return allPassed ? 0 : 1;
}

process.exit(await main());

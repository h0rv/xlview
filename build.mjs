import * as esbuild from 'esbuild';

const watch = process.argv.includes('--watch');

/** @type {esbuild.BuildOptions} */
const shared = {
  bundle: false,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  sourcemap: true,
};

const entryPoints = [
  // Demo viewer + compare
  { in: 'src/demo/viewer.ts', out: 'src/demo/viewer' },
  { in: 'src/demo/compare.ts', out: 'src/demo/compare' },
  // Bench
  { in: 'bench/bench.ts', out: 'bench/bench' },
  { in: 'bench/providers/index.ts', out: 'bench/providers/index' },
  { in: 'bench/providers/xlview.ts', out: 'bench/providers/xlview' },
  { in: 'bench/providers/sheetjs.ts', out: 'bench/providers/sheetjs' },
  // Standalone scripts
  { in: 'verify_comments.ts', out: 'verify_comments' },
  // Browser tests
  { in: 'tests/browser/run_scroll_test.ts', out: 'tests/browser/run_scroll_test' },
];

async function build() {
  for (const entry of entryPoints) {
    const opts = {
      ...shared,
      entryPoints: [entry.in],
      outfile: entry.out + '.js',
    };

    // Node scripts need node platform
    if (
      entry.in === 'verify_comments.ts' ||
      entry.in === 'tests/browser/run_scroll_test.ts'
    ) {
      opts.platform = 'node';
    }

    if (watch) {
      const ctx = await esbuild.context(opts);
      await ctx.watch();
      console.log(`Watching ${entry.in}...`);
    } else {
      await esbuild.build(opts);
      console.log(`Built ${entry.out}.js`);
    }
  }

  if (!watch) {
    console.log('\nAll TypeScript files compiled successfully.');
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});

# xlview - XLSX Viewer
# Rendering: Canvas 2D via web-sys (works in all browsers)

set dotenv-load

# Default recipe - show help
default:
    @just --list

# === Development ===

# Build WASM module for development
build-wasm:
    wasm-pack build --target web --dev

# Build WASM module for release
build-release:
    wasm-pack build --target web --release

# Build TypeScript (compile .ts -> .js via esbuild)
build-ts:
    bun build.mjs

# Type check TypeScript
typecheck:
    bun run tsc --noEmit

# Format TypeScript
ts-fmt:
    bun run prettier --write 'src/demo/**/*.ts' 'bench/**/*.ts' 'tests/browser/**/*.ts' 'verify_comments.ts'

# Lint TypeScript (type check)
ts-lint:
    bun run tsc --noEmit

# Format + lint + build TypeScript
ts-all: ts-fmt ts-lint build-ts

# Build everything (WASM + TypeScript)
build: build-wasm build-ts

# Watch and rebuild on changes
watch:
    cargo watch -s "wasm-pack build --target web --dev"

# === Quality ===

# Format all Rust code
fmt:
    cargo fmt

# Check formatting without changes
fmt-check:
    cargo fmt -- --check

# Run clippy with strict lints on lib code
lint:
    cargo clippy --lib -- -D warnings

# Run all tests
test:
    cargo test --lib

# Run all quality checks (fmt + lint + test + typecheck)
check: fmt-check lint test typecheck

# Run everything: format, build, lint, test
all: fmt build lint test

# === Demo ===

# Serve the demo on port 8080
serve:
    @echo "Starting Canvas 2D viewer at http://localhost:8080"
    python3 -m http.server 8080

# Build and serve demo
demo: build serve

# === E2E Testing ===

# Run full E2E test suite (headless, self-contained, fast)
e2e:
    bun tests/e2e/test_suite.ts

# Build WASM + run E2E suite
e2e-full: build-wasm e2e

# Run quick smoke test (legacy)
e2e-quick:
    bun tests/e2e/quick_test.ts

# Run scroll/header browser tests (legacy)
e2e-scroll:
    bun tests/browser/run_scroll_test.ts

# === Maintenance ===

# Clean build artifacts
clean:
    rm -rf pkg target tests/e2e/output.png

# Update dependencies
update:
    cargo update

# Show outdated dependencies
outdated:
    cargo outdated

# Check WASM bundle size
size: build-release
    @echo "WASM bundle size:"
    @ls -lh pkg/xlview_bg.wasm
    @gzip -c pkg/xlview_bg.wasm | wc -c | awk '{printf "Gzipped: %.1f KB\n", $1/1024}'

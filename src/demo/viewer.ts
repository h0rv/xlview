import init, { XlView } from "../../pkg/xlview.js";
import type { ManifestEntry } from "./types.js";

// Initialize WASM on page load
await init();

const uploadArea = document.getElementById("upload-area") as HTMLLabelElement;
const fileInput = document.getElementById("file-input") as HTMLInputElement;
const viewerContainer = document.getElementById(
  "viewer-container",
) as HTMLDivElement;
let canvas = document.getElementById("viewer-canvas") as HTMLCanvasElement;
let overlayCanvas = document.getElementById(
  "viewer-overlay",
) as HTMLCanvasElement;
const rendererSelect = document.getElementById(
  "renderer-select",
) as HTMLSelectElement;

let viewer: XlView | null = null;
let currentFile: File | null = null;
let renderPending = false;
let activeBackend: "canvas2d" | "webgpu" = "canvas2d";

// Detect WebGPU support and disable option if unavailable
if (!("gpu" in navigator)) {
  const webgpuOption = rendererSelect?.querySelector(
    'option[value="webgpu"]',
  ) as HTMLOptionElement | null;
  if (webgpuOption) {
    webgpuOption.disabled = true;
    webgpuOption.textContent = "WebGPU (not supported)";
  }
}

// Resize canvases to match container.
// The main canvas size is controlled by Rust (oversized for buffer pre-rendering).
// JS only sizes the overlay canvas and reports viewport dimensions.
function resizeCanvas(): number {
  // After setup_native_scroll, the canvas lives inside a clipping spacer.
  // Use the scroll container (marked with data-xlview-scroll) for viewport sizing.
  const target =
    (document.querySelector("[data-xlview-scroll]") as HTMLElement) ??
    canvas.parentElement ??
    viewerContainer;
  const rect = target.getBoundingClientRect();
  const width = target.clientWidth || rect.width;
  const height = target.clientHeight || rect.height;
  const dpr = window.devicePixelRatio || 1;

  const physW = Math.max(1, Math.round(width * dpr));
  const physH = Math.max(1, Math.round(height * dpr));

  // Overlay canvas: always viewport-sized (follows scroll via CSS transform)
  if (activeBackend === "canvas2d") {
    overlayCanvas.width = physW;
    overlayCanvas.height = physH;
    overlayCanvas.style.width = width + "px";
    overlayCanvas.style.height = height + "px";
    overlayCanvas.style.display = "";
  } else {
    // Hide overlay when using wgpu (single-pass renderer)
    overlayCanvas.style.display = "none";
  }

  // Before viewer exists, set main canvas to viewport size as a fallback.
  // Once viewer.resize() is called, Rust overrides to buffer dimensions.
  if (!viewer) {
    canvas.width = physW;
    canvas.height = physH;
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
  }

  const effectiveDpr = width > 0 ? physW / width : dpr;
  if (viewer) {
    // Rust resize() will set the main canvas to oversized buffer dimensions.
    viewer.resize(physW, physH, effectiveDpr);
  }
  return effectiveDpr;
}

(window as unknown as Record<string, unknown>).force_resize = resizeCanvas;

// Request render on next animation frame
function requestRender(): void {
  if (renderPending) return;
  renderPending = true;
  requestAnimationFrame(() => {
    renderPending = false;
    if (viewer) {
      try {
        viewer.render();
      } catch (e) {
        console.error("Render error:", e);
      }
    }
  });
}

// Destroy existing viewer and reset DOM state.
// Creates fresh canvas elements because once a canvas has a context type
// (e.g. "2d"), it cannot be reused with a different context ("webgpu").
function destroyViewer(): void {
  // Remove old elements from DOM *before* freeing the viewer so that
  // no stale events can fire on elements whose Rust closures are dropped.
  viewerContainer.innerHTML = "";

  if (viewer) {
    viewer.free();
    viewer = null;
  }

  // Create fresh canvas elements
  const newCanvas = document.createElement("canvas");
  newCanvas.id = "viewer-canvas";
  const newOverlay = document.createElement("canvas");
  newOverlay.id = "viewer-overlay";

  viewerContainer.appendChild(newCanvas);
  viewerContainer.appendChild(newOverlay);

  canvas = newCanvas;
  overlayCanvas = newOverlay;
}

// Initialize viewer with the selected backend
async function initViewer(): Promise<boolean> {
  try {
    const dpr = resizeCanvas() || window.devicePixelRatio || 1;
    const useWgpu = rendererSelect?.value === "webgpu";

    if (
      useWgpu &&
      "gpu" in navigator &&
      typeof (XlView as unknown as Record<string, unknown>).newWithWgpu ===
        "function"
    ) {
      activeBackend = "webgpu";
      overlayCanvas.style.display = "none";
      viewer = await (
        XlView as unknown as {
          newWithWgpu: (
            c: HTMLCanvasElement,
            d: number,
          ) => Promise<XlView>;
        }
      ).newWithWgpu(canvas, dpr);
    } else {
      activeBackend = "canvas2d";
      overlayCanvas.style.display = "";
      viewer = XlView.newWithOverlay
        ? XlView.newWithOverlay(canvas, overlayCanvas, dpr)
        : new XlView(canvas, dpr);
    }
    viewer.set_render_callback(requestRender);
    // Defer resize/render to next frame to ensure DOM layout is complete
    // after setup_native_scroll moves canvases to new container
    requestAnimationFrame(() => {
      resizeCanvas();
      requestRender();
    });
    return true;
  } catch (e) {
    showError(`Failed to initialize viewer: ${e}`);
    return false;
  }
}

// Renderer toggle: re-create viewer with the selected backend
rendererSelect?.addEventListener("change", async () => {
  const savedFile = currentFile;
  destroyViewer();
  const success = await initViewer();
  if (success && savedFile) {
    try {
      const arrayBuffer = await savedFile.arrayBuffer();
      viewer!.load(new Uint8Array(arrayBuffer));
      resizeCanvas();
    } catch (err) {
      showError(`Failed to reload file: ${(err as Error).message}`);
    }
  }
});

// Drag and drop on upload button
uploadArea.addEventListener("dragover", (e: DragEvent) => {
  e.preventDefault();
  uploadArea.classList.add("dragover");
});

uploadArea.addEventListener("dragleave", () => {
  uploadArea.classList.remove("dragover");
});

uploadArea.addEventListener("drop", (e: DragEvent) => {
  e.preventDefault();
  uploadArea.classList.remove("dragover");

  const file = e.dataTransfer?.files[0];
  if (file) handleFile(file);
});

// File input change
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
});

// Window resize
window.addEventListener("resize", resizeCanvas);

// Load test file buttons dynamically from manifest
async function loadTestFileButtons(): Promise<void> {
  const container = document.getElementById("test-files");
  const uploadBtn = document.getElementById("upload-area");

  try {
    const response = await fetch("test/manifest.json");
    if (!response.ok) {
      // Keep just the upload button
      return;
    }
    const manifest: ManifestEntry[] = await response.json();

    // Insert buttons before the upload button
    for (const file of manifest) {
      const btn = document.createElement("button");
      btn.style.cssText =
        "padding: 8px 16px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; background: #fff;";
      btn.textContent = `${file.name} (${file.size})`;
      btn.addEventListener("click", () => loadTestFile(`test/${file.name}`));
      container?.insertBefore(btn, uploadBtn);
    }
  } catch {
    // Keep just the upload button on error
  }
}
loadTestFileButtons();

async function loadTestFile(path: string): Promise<void> {
  try {
    const response = await fetch(path);
    if (!response.ok)
      throw new Error(`Failed to fetch ${path}: ${response.status}`);
    const blob = await response.blob();
    const file = new File([blob], path.split("/").pop() ?? "file.xlsx", {
      type: blob.type,
    });
    currentFile = file;
    await handleFile(file);
  } catch (err) {
    showError(`Failed to load file: ${(err as Error).message}`);
  }
}

async function handleFile(file: File): Promise<void> {
  if (!file.name.match(/\.(xlsx|xlsm|csv|tsv)$/i)) {
    showError("Supported formats: .xlsx, .xlsm, .csv, .tsv");
    return;
  }

  currentFile = file;

  // Initialize viewer if not already done
  if (!viewer) {
    const success = await initViewer();
    if (!success) return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    viewer!.load(new Uint8Array(arrayBuffer));
    // Force layout recalculation after load - DOM changes from load()
    // may not be reflected in dimensions immediately
    resizeCanvas();
  } catch (err) {
    showError(`Failed to load file: ${(err as Error).message}`);
  }
}

function showError(message: string): void {
  viewerContainer.innerHTML = `<div class="error">${message}</div>`;
  viewer = null;
}

// Suppress unused variable warnings
void currentFile;

// Initialize on page load
initViewer();

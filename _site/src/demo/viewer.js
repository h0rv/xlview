import init, { XlView } from "../../pkg/xlview.js";
await init();
const uploadArea = document.getElementById("upload-area");
const fileInput = document.getElementById("file-input");
const viewerContainer = document.getElementById(
  "viewer-container"
);
const canvas = document.getElementById("viewer-canvas");
const overlayCanvas = document.getElementById(
  "viewer-overlay"
);
let viewer = null;
let currentFile = null;
let renderPending = false;
function resizeCanvas() {
  const target = document.querySelector("[data-xlview-scroll]") ?? canvas.parentElement ?? viewerContainer;
  const rect = target.getBoundingClientRect();
  const width = target.clientWidth || rect.width;
  const height = target.clientHeight || rect.height;
  const dpr = window.devicePixelRatio || 1;
  const physW = Math.max(1, Math.round(width * dpr));
  const physH = Math.max(1, Math.round(height * dpr));
  overlayCanvas.width = physW;
  overlayCanvas.height = physH;
  overlayCanvas.style.width = width + "px";
  overlayCanvas.style.height = height + "px";
  if (!viewer) {
    canvas.width = physW;
    canvas.height = physH;
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
  }
  const effectiveDpr = width > 0 ? physW / width : dpr;
  if (viewer) {
    viewer.resize(physW, physH, effectiveDpr);
  }
  return effectiveDpr;
}
window.force_resize = resizeCanvas;
function requestRender() {
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
async function initViewer() {
  try {
    const dpr = resizeCanvas() || window.devicePixelRatio || 1;
    viewer = XlView.newWithOverlay ? XlView.newWithOverlay(canvas, overlayCanvas, dpr) : new XlView(canvas, dpr);
    viewer.set_render_callback(requestRender);
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
uploadArea.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadArea.classList.add("dragover");
});
uploadArea.addEventListener("dragleave", () => {
  uploadArea.classList.remove("dragover");
});
uploadArea.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadArea.classList.remove("dragover");
  const file = e.dataTransfer?.files[0];
  if (file) handleFile(file);
});
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
});
window.addEventListener("resize", resizeCanvas);
async function loadTestFileButtons() {
  const container = document.getElementById("test-files");
  const uploadBtn = document.getElementById("upload-area");
  try {
    const response = await fetch("test/manifest.json");
    if (!response.ok) {
      return;
    }
    const manifest = await response.json();
    for (const file of manifest) {
      const btn = document.createElement("button");
      btn.style.cssText = "padding: 8px 16px; cursor: pointer; border: 1px solid #ccc; border-radius: 4px; background: #fff;";
      btn.textContent = `${file.name} (${file.size})`;
      btn.addEventListener("click", () => loadTestFile(`test/${file.name}`));
      container?.insertBefore(btn, uploadBtn);
    }
  } catch {
  }
}
loadTestFileButtons();
async function loadTestFile(path) {
  try {
    const response = await fetch(path);
    if (!response.ok)
      throw new Error(`Failed to fetch ${path}: ${response.status}`);
    const blob = await response.blob();
    const file = new File([blob], path.split("/").pop() ?? "file.xlsx", {
      type: blob.type
    });
    currentFile = file;
    await handleFile(file);
  } catch (err) {
    showError(`Failed to load file: ${err.message}`);
  }
}
async function handleFile(file) {
  if (!file.name.match(/\.xlsx?$/i)) {
    showError("Please select an Excel file (.xlsx)");
    return;
  }
  currentFile = file;
  if (!viewer) {
    const success = await initViewer();
    if (!success) return;
  }
  try {
    const arrayBuffer = await file.arrayBuffer();
    viewer.load(new Uint8Array(arrayBuffer));
    resizeCanvas();
  } catch (err) {
    showError(`Failed to load file: ${err.message}`);
  }
}
function showError(message) {
  viewerContainer.innerHTML = `<div class="error">${message}</div>`;
  viewer = null;
}
void currentFile;
initViewer();
//# sourceMappingURL=viewer.js.map

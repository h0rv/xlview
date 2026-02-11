import init, { XlEdit } from "../../pkg/xlview.js";
import type { ManifestEntry } from "./types.js";

// Initialize WASM on page load
await init();

const uploadArea = document.getElementById("upload-area") as HTMLLabelElement;
const fileInput = document.getElementById("file-input") as HTMLInputElement;
const editorContainer = document.getElementById(
  "editor-container",
) as HTMLDivElement;
const canvas = document.getElementById("editor-canvas") as HTMLCanvasElement;
const overlayCanvas = document.getElementById(
  "editor-overlay",
) as HTMLCanvasElement;
const saveBtn = document.getElementById("save-btn") as HTMLButtonElement;
const dirtyBadge = document.getElementById("dirty-badge") as HTMLSpanElement;

let editor: XlEdit | null = null;
let currentFileName = "edited.xlsx";
let renderPending = false;

// Resize canvases to match container
function resizeCanvas(): number {
  const target =
    (document.querySelector("[data-xlview-scroll]") as HTMLElement) ??
    canvas.parentElement ??
    editorContainer;
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

  if (!editor) {
    canvas.width = physW;
    canvas.height = physH;
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
  }

  const effectiveDpr = width > 0 ? physW / width : dpr;
  if (editor) {
    editor.resize(physW, physH, effectiveDpr);
  }
  return effectiveDpr;
}

// Request render on next animation frame
function requestRender(): void {
  if (renderPending) return;
  renderPending = true;
  requestAnimationFrame(() => {
    renderPending = false;
    if (editor) {
      try {
        editor.render();
      } catch (e) {
        console.error("Render error:", e);
      }
    }
  });
}

// Update dirty indicator
function updateDirtyState(): void {
  if (editor?.is_dirty()) {
    dirtyBadge.style.display = "inline-block";
    saveBtn.disabled = false;
    saveBtn.style.opacity = "1";
  } else {
    dirtyBadge.style.display = "none";
    saveBtn.disabled = true;
    saveBtn.style.opacity = "0.5";
  }
}

// Initialize editor
async function initEditor(): Promise<boolean> {
  try {
    const dpr = resizeCanvas() || window.devicePixelRatio || 1;
    editor = new XlEdit(canvas, overlayCanvas, dpr);
    editor.set_render_callback(requestRender);
    requestAnimationFrame(() => {
      resizeCanvas();
      requestRender();
    });
    return true;
  } catch (e) {
    showError(`Failed to initialize editor: ${e}`);
    return false;
  }
}

// Wire up input element events (blur/keydown) on the <input> that Rust creates
let inputEl: HTMLInputElement | null = null;

function commitFromInput() {
  if (!editor || !editor.is_editing()) return;
  const value = editor.input_value();
  if (value != null) {
    editor.commit_edit(value);
    updateDirtyState();
  }
}

function attachInputHandlers(input: HTMLInputElement) {
  if (inputEl === input) return;
  inputEl = input;

  input.addEventListener("blur", () => {
    setTimeout(() => {
      if (editor?.is_editing()) commitFromInput();
    }, 0);
  });

  input.addEventListener("keydown", (event: KeyboardEvent) => {
    if (event.key === "Escape") {
      editor?.cancel_edit();
      event.preventDefault();
    } else if (event.key === "Enter" || event.key === "Tab") {
      commitFromInput();
      event.preventDefault();
    }
  });
}

// Watch for Rust-created <input> appearing
const inputObserver = new MutationObserver((mutations) => {
  for (const m of mutations) {
    for (let i = 0; i < m.addedNodes.length; i++) {
      const node = m.addedNodes[i];
      if (node instanceof HTMLInputElement) attachInputHandlers(node);
    }
  }
});
inputObserver.observe(editorContainer, { childList: true, subtree: true });

// Double-click to begin editing
editorContainer.addEventListener("dblclick", (event: MouseEvent) => {
  if (!editor) return;

  const target =
    (document.querySelector("[data-xlview-scroll]") as HTMLElement) ??
    editorContainer;
  const rect = target.getBoundingClientRect();
  const x = event.clientX - rect.left;
  const y = event.clientY - rect.top;

  const cell = editor.cell_at_point(x, y);
  if (cell && cell.length >= 2) {
    editor.begin_edit(cell[0], cell[1]);
    const inp = editorContainer.querySelector("input");
    if (inp) attachInputHandlers(inp);
  }
});

// Save button
saveBtn.addEventListener("click", () => {
  if (!editor || !editor.is_dirty()) return;
  try {
    const bytes = editor.save();
    const blob = new Blob([bytes as unknown as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = currentFileName.replace(/\.\w+$/, "") + "_edited.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    showError(`Save failed: ${(err as Error).message}`);
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

// Load test file buttons from manifest
async function loadTestFileButtons(): Promise<void> {
  const container = document.getElementById("test-files");
  const uploadBtn = document.getElementById("upload-area");
  try {
    const response = await fetch("test/manifest.json");
    if (!response.ok) return;
    const manifest: ManifestEntry[] = await response.json();
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

  currentFileName = file.name;

  if (!editor) {
    const success = await initEditor();
    if (!success) return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    editor!.load(new Uint8Array(arrayBuffer));
    resizeCanvas();
    updateDirtyState();
  } catch (err) {
    showError(`Failed to load file: ${(err as Error).message}`);
  }
}

function showError(message: string): void {
  editorContainer.innerHTML = `<div class="error">${message}</div>`;
  editor = null;
}

// Initialize on page load
initEditor();

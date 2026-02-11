import _init, { XlEdit } from "../pkg/xlview.js";
import type { InitInput } from "../pkg/xlview.js";

export { XlEdit };

// Reuse init from xl-view (safe to call multiple times)
let _initPromise: Promise<void> | null = null;
export function init(input?: InitInput): Promise<void> {
  if (!_initPromise) _initPromise = _init(input).then(() => {});
  return _initPromise;
}

// Helper: read viewport dimensions from scroll container (or fallback)
function viewportSize(container: HTMLElement) {
  const el =
    (container.querySelector("[data-xlview-scroll]") as HTMLElement) ??
    container;
  const w = el.clientWidth || el.getBoundingClientRect().width;
  const h = el.clientHeight || el.getBoundingClientRect().height;
  const dpr = window.devicePixelRatio || 1;
  return {
    w,
    h,
    dpr,
    pw: Math.max(1, Math.round(w * dpr)),
    ph: Math.max(1, Math.round(h * dpr)),
  };
}

export interface MountedEditor {
  /** Load an XLSX file from bytes */
  load(data: Uint8Array): void;
  /** Save modified XLSX — returns the bytes */
  save(): Uint8Array;
  /** Download modified XLSX as a file */
  download(filename?: string): void;
  /** Destroy the editor and release resources */
  destroy(): void;
  /** Access the underlying XlEdit instance */
  readonly editor: XlEdit;
}

export async function mountEditor(
  container: HTMLElement,
): Promise<MountedEditor> {
  await init();

  const dpr = window.devicePixelRatio || 1;
  const pw = Math.max(1, Math.round((container.clientWidth || 300) * dpr));
  const ph = Math.max(1, Math.round((container.clientHeight || 150) * dpr));

  const base = document.createElement("canvas");
  const overlay = document.createElement("canvas");
  base.width = pw;
  base.height = ph;
  overlay.width = pw;
  overlay.height = ph;

  // Ensure positioning context
  const pos = getComputedStyle(container).position;
  if (!pos || pos === "static") container.style.position = "relative";

  container.appendChild(base);
  container.appendChild(overlay);

  const editor = new XlEdit(base, overlay, dpr);

  // RAF-batched render callback
  let pending = false;
  editor.set_render_callback(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      try {
        editor.render();
      } catch {
        /* freed */
      }
    });
  });

  // Resize handler
  const doResize = () => {
    const { w, h, dpr: d, pw: rpw, ph: rph } = viewportSize(container);
    overlay.width = rpw;
    overlay.height = rph;
    overlay.style.width = w + "px";
    overlay.style.height = h + "px";
    editor.resize(rpw, rph, d);
  };
  await new Promise<void>((r) =>
    requestAnimationFrame(() => {
      doResize();
      r();
    }),
  );

  // ResizeObserver
  let destroyed = false;
  const ro = new ResizeObserver(() => {
    if (!destroyed) doResize();
  });
  ro.observe(container);

  // Wire up input element events (blur/keydown) via MutationObserver
  // so they attach to the <input> that Rust creates dynamically.
  let inputEl: HTMLInputElement | null = null;

  function commitFromInput() {
    if (!editor.is_editing()) return;
    const value = editor.input_value();
    if (value != null) {
      editor.commit_edit(value);
    }
  }

  function attachInputHandlers(input: HTMLInputElement) {
    if (inputEl === input) return;
    inputEl = input;

    input.addEventListener("blur", () => {
      // Small delay so click-based commits (Enter/Tab) fire first
      setTimeout(() => {
        if (editor.is_editing()) commitFromInput();
      }, 0);
    });

    input.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        editor.cancel_edit();
        event.preventDefault();
      } else if (event.key === "Enter" || event.key === "Tab") {
        commitFromInput();
        event.preventDefault();
      }
    });
  }

  // Watch for Rust-created <input> appearing in the container
  const observer = new MutationObserver((mutations) => {
    for (const m of mutations) {
      for (let i = 0; i < m.addedNodes.length; i++) {
        const node = m.addedNodes[i];
        if (node instanceof HTMLInputElement) attachInputHandlers(node);
      }
    }
  });
  observer.observe(container, { childList: true, subtree: true });

  // Also check for any input already present
  const existing = container.querySelector("input");
  if (existing) attachInputHandlers(existing);

  // Double-click to begin editing
  container.addEventListener("dblclick", (event: MouseEvent) => {
    const rect = container.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const cell = editor.cell_at_point(x, y);
    if (cell && cell.length >= 2) {
      editor.begin_edit(cell[0], cell[1]);
      // Re-check for input after begin_edit creates it
      const inp = container.querySelector("input");
      if (inp) attachInputHandlers(inp);
    }
  });

  return {
    load(data: Uint8Array) {
      editor.load(data);
      requestAnimationFrame(doResize);
    },
    save(): Uint8Array {
      return editor.save();
    },
    download(filename = "edited.xlsx") {
      const bytes = editor.save();
      const blob = new Blob([bytes as unknown as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    },
    destroy() {
      destroyed = true;
      ro.disconnect();
      editor.free();
    },
    get editor() {
      return editor;
    },
  };
}

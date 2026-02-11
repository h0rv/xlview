export { XlEdit } from "../pkg/xlview.js";
export type { InitInput, InitOutput } from "../pkg/xlview.js";

/** Initialize the WASM module. Safe to call multiple times. */
export function init(
  input?: import("../pkg/xlview.js").InitInput,
): Promise<void>;

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
  readonly editor: import("../pkg/xlview.js").XlEdit;
}

/**
 * Mount an xl-edit instance into a container element.
 * Creates canvases, sets up resize handling + editing, and returns a controller.
 * Double-click a cell to edit. Enter commits, Escape cancels.
 */
export function mountEditor(container: HTMLElement): Promise<MountedEditor>;

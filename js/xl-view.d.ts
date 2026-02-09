export {
  XlView,
  parse_xlsx,
  parse_xlsx_to_js,
  parse_xlsx_metrics,
  version,
} from "../pkg/xlview.js";
export type { InitInput, InitOutput } from "../pkg/xlview.js";

/** Initialize the WASM module. Safe to call multiple times. */
export function init(input?: import("../pkg/xlview.js").InitInput): Promise<void>;

export interface MountedViewer {
  /** Load an XLSX file from bytes */
  load(data: Uint8Array): void;
  /** Destroy the viewer and release resources */
  destroy(): void;
  /** Access the underlying XlView instance */
  readonly viewer: import("../pkg/xlview.js").XlView;
}

/**
 * Mount an xlview instance into a container element.
 * Creates canvases, sets up resize handling, and returns a controller.
 */
export function mount(container: HTMLElement): Promise<MountedViewer>;

/** `<xl-view>` custom element for drop-in usage */
export declare class XlViewElement extends HTMLElement {
  static readonly observedAttributes: string[];
  /** Load XLSX from bytes */
  load(data: Uint8Array): void;
  /** Access underlying XlView instance (null before mount completes) */
  readonly xlview: import("../pkg/xlview.js").XlView | null;
}

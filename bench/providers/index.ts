import { createXlviewProvider } from "./xlview.js";
import { createSheetJsProvider } from "./sheetjs.js";
import type { BenchProviderEntry } from "../../src/demo/types.js";

export const providers: Record<string, BenchProviderEntry> = {
  xlview: {
    id: "xlview",
    label: "xlview (WASM)",
    create: createXlviewProvider,
  },
  sheetjs: {
    id: "sheetjs",
    label: "SheetJS (DOM)",
    create: createSheetJsProvider,
  },
};

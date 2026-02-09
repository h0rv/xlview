import type {
  BenchProvider,
  ProviderCreateOpts,
} from "../../src/demo/types.js";

// SheetJS has no @types package; use `any` for the module.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let XLSX: any = null;

async function loadXlsx(): Promise<unknown> {
  if (XLSX) return XLSX;
  XLSX = await import("xlsx");
  return XLSX;
}

function ensureDomRoot(
  domRoot: HTMLElement | null,
  canvas: HTMLCanvasElement | null,
): void {
  if (canvas) canvas.style.display = "none";
  if (domRoot) {
    domRoot.style.display = "block";
    domRoot.innerHTML = "";
  }
}

export async function createSheetJsProvider({
  domRoot,
  canvas,
}: ProviderCreateOpts): Promise<BenchProvider> {
  const xlsx = await loadXlsx();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let workbook: any = null;
  let activeSheetName: string | null = null;

  ensureDomRoot(domRoot, canvas);

  function renderSheet(): void {
    if (!workbook || !domRoot) return;
    const sheetName = activeSheetName ?? workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const html: string = (xlsx as any).utils.sheet_to_html(sheet, {
      id: "sheetjs-table",
      editable: false,
    });
    domRoot.innerHTML = html;
    const table = domRoot.querySelector("table");
    if (table) {
      table.style.borderCollapse = "collapse";
      table.style.width = "100%";
      table.style.fontSize = "12px";
    }
    const cells = domRoot.querySelectorAll("td, th");
    cells.forEach((cell) => {
      (cell as HTMLElement).style.border = "1px solid #d1d5db";
      (cell as HTMLElement).style.padding = "2px 4px";
      (cell as HTMLElement).style.whiteSpace = "nowrap";
    });
  }

  return {
    id: "sheetjs",
    label: "SheetJS (DOM)",
    async parse(bytes: Uint8Array) {
      const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (xlsx as any).read(data, {
        type: "array",
        cellStyles: true,
        dense: true,
      });
    },
    async load(bytes: Uint8Array) {
      const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbook = (xlsx as any).read(data, {
        type: "array",
        cellStyles: true,
        dense: true,
      });
      activeSheetName = workbook.SheetNames[0] ?? null;
      return workbook;
    },
    async render() {
      renderSheet();
    },
    async scroll({ dx, dy }) {
      if (!domRoot) return;
      domRoot.scrollLeft += dx ?? 0;
      domRoot.scrollTop += dy ?? 0;
    },
    async resetScroll() {
      if (!domRoot) return;
      domRoot.scrollLeft = 0;
      domRoot.scrollTop = 0;
    },
  };
}

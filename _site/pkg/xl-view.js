import _init, {
  XlView,
  parse_xlsx,
  parse_xlsx_to_js,
  parse_xlsx_metrics,
  version
} from "../pkg/xlview.js";
let _initPromise = null;
function init(input) {
  if (!_initPromise) _initPromise = _init(input).then(() => {
  });
  return _initPromise;
}
function viewportSize(container) {
  const el = container.querySelector("[data-xlview-scroll]") ?? container;
  const w = el.clientWidth || el.getBoundingClientRect().width;
  const h = el.clientHeight || el.getBoundingClientRect().height;
  const dpr = window.devicePixelRatio || 1;
  return {
    w,
    h,
    dpr,
    pw: Math.max(1, Math.round(w * dpr)),
    ph: Math.max(1, Math.round(h * dpr))
  };
}
async function mount(container) {
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
  const pos = getComputedStyle(container).position;
  if (!pos || pos === "static") container.style.position = "relative";
  container.appendChild(base);
  container.appendChild(overlay);
  const viewer = XlView.newWithOverlay(base, overlay, dpr);
  let pending = false;
  viewer.set_render_callback(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      try {
        viewer.render();
      } catch {
      }
    });
  });
  const doResize = () => {
    const { w, h, dpr: d, pw: rpw, ph: rph } = viewportSize(container);
    overlay.width = rpw;
    overlay.height = rph;
    overlay.style.width = w + "px";
    overlay.style.height = h + "px";
    viewer.resize(rpw, rph, d);
  };
  await new Promise(
    (r) => requestAnimationFrame(() => {
      doResize();
      r();
    })
  );
  let destroyed = false;
  const ro = new ResizeObserver(() => {
    if (!destroyed) doResize();
  });
  ro.observe(container);
  return {
    load(data) {
      viewer.load(data);
      requestAnimationFrame(doResize);
    },
    destroy() {
      destroyed = true;
      ro.disconnect();
      viewer.free();
    },
    get viewer() {
      return viewer;
    }
  };
}
class XlViewElement extends HTMLElement {
  _mounted = null;
  _shadow;
  _container;
  static get observedAttributes() {
    return ["src"];
  }
  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: "open" });
    this._container = document.createElement("div");
    this._container.style.cssText = "width:100%;height:100%;position:relative;overflow:hidden";
    this._shadow.appendChild(this._container);
  }
  async connectedCallback() {
    this._mounted = await mount(this._container);
    const src = this.getAttribute("src");
    if (src) this._loadUrl(src);
    this.dispatchEvent(new Event("ready"));
  }
  disconnectedCallback() {
    this._mounted?.destroy();
    this._mounted = null;
  }
  attributeChangedCallback(name, old, val) {
    if (name === "src" && val && val !== old && this._mounted)
      this._loadUrl(val);
  }
  async _loadUrl(url) {
    const res = await fetch(url);
    if (!res.ok) {
      this.dispatchEvent(
        new CustomEvent("error", { detail: `HTTP ${res.status}` })
      );
      return;
    }
    this._mounted?.load(new Uint8Array(await res.arrayBuffer()));
  }
  /** Load XLSX from bytes */
  load(data) {
    this._mounted?.load(data);
  }
  /** Access underlying XlView */
  get xlview() {
    return this._mounted?.viewer ?? null;
  }
}
customElements.define("xl-view", XlViewElement);
export {
  XlView,
  XlViewElement,
  init,
  mount,
  parse_xlsx,
  parse_xlsx_metrics,
  parse_xlsx_to_js,
  version
};
//# sourceMappingURL=xl-view.js.map

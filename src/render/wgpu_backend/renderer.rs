//! wgpu renderer backend for xlview.
//!
//! Implements `RenderBackend` using WebGPU via the `wgpu` crate. Text is
//! rendered on a transparent Canvas 2D overlay positioned over the WebGPU
//! canvas, reusing the browser's native font pipeline for perfect quality.

use super::buffers::{orthographic_projection, Globals, LineInstance, RectInstance};
use super::pipelines;
use crate::cell_ref::{parse_cell_range, parse_sqref};
use crate::error::Result;
use crate::render::backend::{CellRenderData, CellStyleData, RenderBackend, RenderParams};
use crate::render::colors::{self, parse_color_f32, Rgb};
use crate::render::selection;
use crate::types::{CFRuleType, ValidationType};

use wasm_bindgen::JsCast;
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement};

/// Maximum instances per buffer. 128k instances should cover any reasonable
/// spreadsheet viewport.
const MAX_RECTS: usize = 131_072;
const MAX_LINES: usize = 262_144;

/// Default cell text color (black).
const DEFAULT_TEXT_COLOR: [f32; 4] = [0.0, 0.0, 0.0, 1.0];
/// Default background color (white).
const DEFAULT_BG_COLOR: [f32; 4] = [1.0, 1.0, 1.0, 1.0];
/// Grid line color (light gray).
const GRID_COLOR: [f32; 4] = [0.878, 0.878, 0.878, 1.0]; // #E0E0E0
/// Selection fill color (translucent blue).
const SELECTION_FILL: [f32; 4] = [0.102, 0.463, 0.824, 0.15];
/// Selection border color (solid blue).
const SELECTION_BORDER: [f32; 4] = [0.102, 0.463, 0.824, 1.0];
/// Header background color.
const HEADER_BG: [f32; 4] = [0.969, 0.969, 0.969, 1.0]; // #F8F8F8
/// Header background color (selected).
const HEADER_BG_SELECTED: [f32; 4] = [0.855, 0.918, 0.965, 1.0]; // #DAE9F6
/// Header border color.
const HEADER_BORDER: [f32; 4] = [0.839, 0.839, 0.839, 1.0]; // #D6D6D6
/// Header text color.
const HEADER_TEXT_COLOR: [f32; 4] = [0.333, 0.333, 0.333, 1.0]; // #555
/// Frozen pane divider color.
const FROZEN_DIVIDER_COLOR: [f32; 4] = [0.6, 0.6, 0.6, 1.0]; // #999
/// Comment indicator color (red).
const COMMENT_INDICATOR_COLOR: [f32; 4] = [1.0, 0.0, 0.0, 1.0]; // #FF0000
/// Validation dropdown arrow color (gray).
const VALIDATION_ARROW_COLOR: [f32; 4] = [0.4, 0.4, 0.4, 1.0]; // #666

/// Convert [f32; 4] RGBA to CSS rgba() string.
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
fn color_to_css(c: [f32; 4]) -> String {
    let r = (c[0] * 255.0).round().clamp(0.0, 255.0) as u8;
    let g = (c[1] * 255.0).round().clamp(0.0, 255.0) as u8;
    let b = (c[2] * 255.0).round().clamp(0.0, 255.0) as u8;
    format!("rgba({},{},{},{})", r, g, b, c[3])
}

/// Build CSS font string: "italic bold 14px Arial"
fn build_font_string(size: f32, family: &str, bold: bool, italic: bool) -> String {
    let style = if italic { "italic " } else { "" };
    let weight = if bold { "bold " } else { "" };
    format!("{style}{weight}{size}px {family}")
}

/// Fingerprint of the frame state for dirty-skip optimization.
#[derive(Clone, PartialEq)]
struct FrameFingerprint {
    scroll_x: u32,
    scroll_y: u32,
    scale: u32,
    sel_bounds: Option<(u32, u32, u32, u32)>,
    width: u32,
    height: u32,
    cell_count: usize,
}

impl FrameFingerprint {
    fn from_params(params: &RenderParams, width: u32, height: u32) -> Self {
        Self {
            scroll_x: params.viewport.scroll_x.to_bits(),
            scroll_y: params.viewport.scroll_y.to_bits(),
            scale: params.viewport.scale.to_bits(),
            sel_bounds: params.selection,
            width,
            height,
            cell_count: params.cells.len(),
        }
    }
}

/// The wgpu-based spreadsheet renderer.
///
/// WebGPU handles geometry (backgrounds, grid lines, borders, selection).
/// A transparent Canvas 2D overlay handles all text rendering.
pub struct WgpuRenderer {
    canvas: HtmlCanvasElement,
    device: wgpu::Device,
    queue: wgpu::Queue,
    surface: wgpu::Surface<'static>,
    surface_config: wgpu::SurfaceConfiguration,
    rect_pipeline: wgpu::RenderPipeline,
    line_pipeline: wgpu::RenderPipeline,
    globals_buffer: wgpu::Buffer,
    globals_bind_group: wgpu::BindGroup,
    rect_buffer: wgpu::Buffer,
    line_buffer: wgpu::Buffer,
    // Text overlay (Canvas 2D positioned over the WebGPU canvas)
    text_overlay: HtmlCanvasElement,
    text_ctx: CanvasRenderingContext2d,
    width: u32,
    height: u32,
    dpr: f32,
    // Retained instance buffers (reused across frames)
    retained_rects: Vec<RectInstance>,
    retained_lines: Vec<LineInstance>,
    // Dirty frame skip
    last_fingerprint: Option<FrameFingerprint>,
    // Geometry cache: counts for base geometry (cells+headers+grid)
    cached_base_rects: usize,
    cached_base_lines: usize,
}

impl WgpuRenderer {
    /// Create a new wgpu renderer. This is async because WebGPU adapter/device
    /// creation is asynchronous.
    #[allow(clippy::cast_possible_truncation)]
    pub async fn new(
        canvas: HtmlCanvasElement,
        dpr: f32,
    ) -> std::result::Result<Self, String> {
        let width = canvas.width().max(1);
        let height = canvas.height().max(1);

        // Get wgpu instance with WebGPU backend
        let instance = wgpu::Instance::new(&wgpu::InstanceDescriptor {
            backends: wgpu::Backends::BROWSER_WEBGPU,
            ..Default::default()
        });

        // Create surface from canvas
        let surface_target = wgpu::SurfaceTarget::Canvas(canvas.clone());
        let surface = instance
            .create_surface(surface_target)
            .map_err(|e| format!("Failed to create surface: {e}"))?;

        let adapter = instance
            .request_adapter(&wgpu::RequestAdapterOptions {
                power_preference: wgpu::PowerPreference::HighPerformance,
                compatible_surface: Some(&surface),
                force_fallback_adapter: false,
            })
            .await
            .map_err(|e| format!("No suitable GPU adapter found: {e}"))?;

        let (device, queue) = adapter
            .request_device(&wgpu::DeviceDescriptor {
                label: Some("xlview device"),
                required_features: wgpu::Features::empty(),
                required_limits: wgpu::Limits::downlevel_webgl2_defaults(),
                experimental_features: wgpu::ExperimentalFeatures::default(),
                memory_hints: wgpu::MemoryHints::MemoryUsage,
                trace: wgpu::Trace::Off,
            })
            .await
            .map_err(|e| format!("Failed to create device: {e}"))?;

        let surface_caps = surface.get_capabilities(&adapter);
        let format = surface_caps
            .formats
            .first()
            .copied()
            .ok_or_else(|| "No surface formats available".to_string())?;

        let surface_config = wgpu::SurfaceConfiguration {
            usage: wgpu::TextureUsages::RENDER_ATTACHMENT,
            format,
            width,
            height,
            present_mode: wgpu::PresentMode::AutoVsync,
            alpha_mode: surface_caps
                .alpha_modes
                .first()
                .copied()
                .unwrap_or(wgpu::CompositeAlphaMode::Auto),
            view_formats: vec![],
            desired_maximum_frame_latency: 2,
        };
        surface.configure(&device, &surface_config);

        // Globals uniform buffer + bind group layout
        let globals_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("globals uniform"),
            size: std::mem::size_of::<Globals>() as u64,
            usage: wgpu::BufferUsages::UNIFORM | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });

        let globals_layout = device.create_bind_group_layout(&wgpu::BindGroupLayoutDescriptor {
            label: Some("globals layout"),
            entries: &[wgpu::BindGroupLayoutEntry {
                binding: 0,
                visibility: wgpu::ShaderStages::VERTEX,
                ty: wgpu::BindingType::Buffer {
                    ty: wgpu::BufferBindingType::Uniform,
                    has_dynamic_offset: false,
                    min_binding_size: None,
                },
                count: None,
            }],
        });

        let globals_bind_group = device.create_bind_group(&wgpu::BindGroupDescriptor {
            label: Some("globals bind group"),
            layout: &globals_layout,
            entries: &[wgpu::BindGroupEntry {
                binding: 0,
                resource: globals_buffer.as_entire_binding(),
            }],
        });

        // Pipelines (geometry only — text is rendered on Canvas 2D overlay)
        let rect_pipeline = pipelines::create_rect_pipeline(&device, format, &globals_layout);
        let line_pipeline = pipelines::create_line_pipeline(&device, format, &globals_layout);

        // Instance buffers
        let rect_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("rect instances"),
            size: (MAX_RECTS * std::mem::size_of::<RectInstance>()) as u64,
            usage: wgpu::BufferUsages::VERTEX | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });
        let line_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("line instances"),
            size: (MAX_LINES * std::mem::size_of::<LineInstance>()) as u64,
            usage: wgpu::BufferUsages::VERTEX | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });

        // Create text overlay canvas (transparent, positioned over WebGPU canvas)
        let document = web_sys::window()
            .ok_or("No window")?
            .document()
            .ok_or("No document")?;
        let text_overlay: HtmlCanvasElement = document
            .create_element("canvas")
            .map_err(|_| "Failed to create overlay canvas")?
            .dyn_into()
            .map_err(|_| "Element is not a canvas")?;
        text_overlay.set_width(width);
        text_overlay.set_height(height);
        let overlay_style = text_overlay.style();
        let _ = overlay_style.set_property("position", "absolute");
        let _ = overlay_style.set_property("top", "0");
        let _ = overlay_style.set_property("left", "0");
        let _ = overlay_style.set_property("pointer-events", "none");
        let _ = overlay_style.set_property("z-index", "0");
        // Match the CSS dimensions of the main canvas
        let css_w = f64::from(width) / f64::from(dpr);
        let css_h = f64::from(height) / f64::from(dpr);
        let _ = overlay_style.set_property("width", &format!("{css_w}px"));
        let _ = overlay_style.set_property("height", &format!("{css_h}px"));

        let text_ctx: CanvasRenderingContext2d = text_overlay
            .get_context("2d")
            .map_err(|_| "Failed to get 2d context")?
            .ok_or("No 2d context")?
            .dyn_into()
            .map_err(|_| "Context is not CanvasRenderingContext2d")?;

        Ok(Self {
            canvas,
            device,
            queue,
            surface,
            surface_config,
            rect_pipeline,
            line_pipeline,
            globals_buffer,
            globals_bind_group,
            rect_buffer,
            line_buffer,
            text_overlay,
            text_ctx,
            width,
            height,
            dpr,
            retained_rects: Vec::with_capacity(4096),
            retained_lines: Vec::with_capacity(8192),
            last_fingerprint: None,
            cached_base_rects: 0,
            cached_base_lines: 0,
        })
    }

    /// Ensure the text overlay canvas is inserted into the DOM.
    /// Called at the start of each render. Idempotent — only inserts if
    /// the overlay has no parent (e.g. after `setup_native_scroll` restructures the DOM).
    fn ensure_overlay_in_dom(&self) {
        if self.text_overlay.parent_node().is_some() {
            return;
        }
        // Insert overlay as next sibling of the main canvas
        if let Some(parent) = self.canvas.parent_node() {
            let _ = parent.insert_before(
                &self.text_overlay,
                self.canvas.next_sibling().as_ref(),
            );
        }
    }

    /// Set CSS dimensions on the canvas element (logical pixels).
    pub fn set_canvas_css_size(&self, css_w: f32, css_h: f32) {
        let style = self.canvas.style();
        let _ = style.set_property("width", &format!("{css_w}px"));
        let _ = style.set_property("height", &format!("{css_h}px"));
        // Match overlay CSS size
        let overlay_style = self.text_overlay.style();
        let _ = overlay_style.set_property("width", &format!("{css_w}px"));
        let _ = overlay_style.set_property("height", &format!("{css_h}px"));
    }

    /// Reset the CSS transform on the canvas (clear scroll compensation offset).
    pub fn reset_canvas_transform(&self) {
        let _ = self
            .canvas
            .style()
            .set_property("transform", "translate(0px, 0px)");
    }

    /// Collect geometry from render params and issue draw calls.
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        clippy::too_many_lines
    )]
    fn render_frame(&mut self, params: &RenderParams) -> Result<()> {
        // --- Dirty frame skip ---
        let fingerprint = FrameFingerprint::from_params(params, self.width, self.height);
        if self.last_fingerprint.as_ref() == Some(&fingerprint) {
            return Ok(());
        }
        self.last_fingerprint = Some(fingerprint);

        // Ensure text overlay is in the DOM
        self.ensure_overlay_in_dom();

        let layout = params.layout;
        let viewport = params.viewport;

        let phys_w = self.width as f32;
        let phys_h = self.height as f32;

        // Header offsets in physical pixels
        let header_offset_x = if params.show_headers {
            params.header_config.row_header_width * self.dpr
        } else {
            0.0
        };
        let header_offset_y = if params.show_headers {
            params.header_config.col_header_height * self.dpr
        } else {
            0.0
        };

        // Reuse retained buffers
        let rects = &mut self.retained_rects;
        let lines = &mut self.retained_lines;
        rects.clear();
        lines.clear();

        // Clear text overlay
        let text_ctx = &self.text_ctx;
        text_ctx.clear_rect(0.0, 0.0, f64::from(phys_w), f64::from(phys_h));
        text_ctx.set_text_baseline("middle");

        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();
        let scale = viewport.scale * self.dpr;
        let dpr = self.dpr;

        // --- Headers ---
        if params.show_headers {
            let header_w = params.header_config.row_header_width * dpr;
            let header_h = params.header_config.col_header_height * dpr;

            // Top-left corner
            rects.push(RectInstance {
                pos: [0.0, 0.0],
                size: [header_w, header_h],
                color: HEADER_BG,
            });

            // Column headers
            let header_font = build_font_string(10.0 * dpr, "Arial", false, false);
            let header_css_color = color_to_css(HEADER_TEXT_COLOR);
            text_ctx.set_font(&header_font);
            text_ctx.set_fill_style_str(&header_css_color);

            let (start_col, end_col) = viewport.visible_cols(layout);
            for col_idx in start_col..end_col {
                let col_x = layout.col_positions.get(col_idx as usize).copied().unwrap_or(0.0);
                let col_w = layout.col_width(col_idx);
                let is_frozen = col_idx < layout.frozen_cols;
                let screen_x = if is_frozen {
                    col_x * scale
                } else {
                    frozen_width * scale + (col_x - viewport.scroll_x) * scale
                };
                let px = screen_x + header_w;
                let pw = col_w * scale;

                // Header selection highlighting
                let bg = if is_col_in_selection(col_idx, params.header_selection) {
                    HEADER_BG_SELECTED
                } else {
                    HEADER_BG
                };

                rects.push(RectInstance {
                    pos: [px, 0.0],
                    size: [pw, header_h],
                    color: bg,
                });
                // Column header border
                lines.push(LineInstance {
                    start: [px + pw, 0.0],
                    end: [px + pw, header_h],
                    width_pad: [1.0, 0.0],
                    color: HEADER_BORDER,
                });
                // Column label text (Canvas 2D)
                let label = col_label(col_idx);
                if let Ok(metrics) = text_ctx.measure_text(&label) {
                    let tw = metrics.width() as f32;
                    let text_x = f64::from(px + (pw - tw) / 2.0);
                    let text_y = f64::from(header_h / 2.0);
                    let _ = text_ctx.fill_text(&label, text_x, text_y);
                }
            }

            // Row headers
            let (start_row, end_row) = viewport.visible_rows(layout);
            // Re-set font/color in case something changed (defensive)
            text_ctx.set_font(&header_font);
            text_ctx.set_fill_style_str(&header_css_color);

            for row_idx in start_row..end_row {
                let row_y = layout.row_positions.get(row_idx as usize).copied().unwrap_or(0.0);
                let row_h = layout.row_height(row_idx);
                let is_frozen = row_idx < layout.frozen_rows;
                let screen_y = if is_frozen {
                    row_y * scale
                } else {
                    frozen_height * scale + (row_y - viewport.scroll_y) * scale
                };
                let py = screen_y + header_h;
                let ph = row_h * scale;

                // Header selection highlighting
                let bg = if is_row_in_selection(row_idx, params.header_selection) {
                    HEADER_BG_SELECTED
                } else {
                    HEADER_BG
                };

                rects.push(RectInstance {
                    pos: [0.0, py],
                    size: [header_w, ph],
                    color: bg,
                });
                // Row header border
                lines.push(LineInstance {
                    start: [0.0, py + ph],
                    end: [header_w, py + ph],
                    width_pad: [1.0, 0.0],
                    color: HEADER_BORDER,
                });
                // Row label text (Canvas 2D)
                let label = format!("{}", row_idx + 1);
                if let Ok(metrics) = text_ctx.measure_text(&label) {
                    let tw = metrics.width() as f32;
                    let text_x = f64::from((header_w - tw) / 2.0);
                    let text_y = f64::from(py + ph / 2.0);
                    let _ = text_ctx.fill_text(&label, text_x, text_y);
                }
            }

            // Header border lines
            lines.push(LineInstance {
                start: [header_w, 0.0],
                end: [header_w, phys_h],
                width_pad: [1.0, 0.0],
                color: HEADER_BORDER,
            });
            lines.push(LineInstance {
                start: [0.0, header_h],
                end: [phys_w, header_h],
                width_pad: [1.0, 0.0],
                color: HEADER_BORDER,
            });
        }

        // --- Collect CF color overrides ---
        let cf_overrides = collect_cf_color_overrides(
            params.cells,
            params.conditional_formatting,
            params.dxf_styles,
            params.conditional_formatting_cache,
        );

        // --- Cells: backgrounds, grid lines, text, indicators ---
        let hox = header_offset_x;
        let hoy = header_offset_y;

        for cell in params.cells {
            let col_x = layout.col_positions.get(cell.col as usize).copied().unwrap_or(0.0);
            let col_w = layout.col_width(cell.col);
            let row_y = layout.row_positions.get(cell.row as usize).copied().unwrap_or(0.0);
            let row_h = layout.row_height(cell.row);

            let is_col_frozen = cell.col < layout.frozen_cols;
            let is_row_frozen = cell.row < layout.frozen_rows;

            let screen_x = if is_col_frozen {
                col_x * scale
            } else {
                frozen_width * scale + (col_x - viewport.scroll_x) * scale
            };
            let screen_y = if is_row_frozen {
                row_y * scale
            } else {
                frozen_height * scale + (row_y - viewport.scroll_y) * scale
            };

            let px = screen_x + hox;
            let py = screen_y + hoy;
            let pw = col_w * scale;
            let ph = row_h * scale;

            // Resolve style
            let style = cell
                .style_override
                .as_ref()
                .or_else(|| {
                    cell.style_idx
                        .and_then(|i| params.style_cache.get(i))
                        .and_then(|s| s.as_ref())
                })
                .or(params.default_style.as_ref());

            // Cell background — check CF override, then pattern fill, then bg_color
            let bg = cf_overrides
                .get(&(cell.row, cell.col))
                .and_then(|ov| ov.bg_color.as_ref())
                .and_then(|c| parse_color_f32(c))
                .or_else(|| resolve_pattern_fill(style))
                .or_else(|| {
                    style
                        .and_then(|s| s.bg_color.as_deref())
                        .and_then(parse_color_f32)
                })
                .unwrap_or(DEFAULT_BG_COLOR);

            rects.push(RectInstance {
                pos: [px, py],
                size: [pw, ph],
                color: bg,
            });

            // Grid lines (right and bottom edges)
            lines.push(LineInstance {
                start: [px + pw, py],
                end: [px + pw, py + ph],
                width_pad: [1.0, 0.0],
                color: GRID_COLOR,
            });
            lines.push(LineInstance {
                start: [px, py + ph],
                end: [px + pw, py + ph],
                width_pad: [1.0, 0.0],
                color: GRID_COLOR,
            });

            // Borders (override grid lines) + diagonal borders
            if let Some(s) = style {
                emit_borders(s, px, py, pw, ph, dpr, lines);
            }

            // Comment indicator (red triangle in top-right corner)
            if cell.has_comment == Some(true) {
                let indicator_size = 6.0 * dpr;
                // Small red rect at top-right corner
                rects.push(RectInstance {
                    pos: [px + pw - indicator_size, py],
                    size: [indicator_size, indicator_size],
                    color: COMMENT_INDICATOR_COLOR,
                });
            }

            // Cell text (Canvas 2D overlay)
            let cf_font_override = cf_overrides.get(&(cell.row, cell.col));
            if let Some(value) = cell.value.as_deref() {
                if !value.is_empty() {
                    let font_size = style.and_then(|s| s.font_size).unwrap_or(11.0) * dpr;
                    let font_family = style
                        .and_then(|s| s.font_family.as_deref())
                        .or(params.minor_font)
                        .unwrap_or("Arial");
                    let bold = style.and_then(|s| s.bold).unwrap_or(false);
                    let italic = style.and_then(|s| s.italic).unwrap_or(false);
                    let text_color = cf_font_override
                        .and_then(|ov| ov.font_color.as_ref())
                        .and_then(|c| parse_color_f32(c))
                        .or_else(|| {
                            style
                                .and_then(|s| s.font_color.as_deref())
                                .and_then(parse_color_f32)
                        })
                        .unwrap_or(DEFAULT_TEXT_COLOR);

                    let font_str = build_font_string(font_size, font_family, bold, italic);
                    text_ctx.set_font(&font_str);
                    text_ctx.set_fill_style_str(&color_to_css(text_color));

                    if let Ok(metrics) = text_ctx.measure_text(value) {
                        let tw = metrics.width() as f32;

                        // Horizontal alignment
                        let align = style.and_then(|s| s.align_h.as_deref());
                        let padding = 3.0 * dpr;
                        let text_x = match align {
                            Some("center") | Some("centerContinuous") => {
                                px + (pw - tw) / 2.0
                            }
                            Some("right") => px + pw - tw - padding,
                            Some("left") | Some("fill") | Some("justify") | Some("distributed") => px + padding,
                            _ => {
                                // Auto: right-align numbers, left-align text
                                if cell.numeric_value.is_some() {
                                    px + pw - tw - padding
                                } else {
                                    px + padding
                                }
                            }
                        };
                        // Vertical center (text_baseline is "middle")
                        let text_y = py + ph / 2.0;

                        let _ = text_ctx.fill_text(
                            value,
                            f64::from(text_x),
                            f64::from(text_y),
                        );
                    }
                }
            }
        }

        // Record base geometry counts for potential future cache use
        self.cached_base_rects = rects.len();
        self.cached_base_lines = lines.len();

        // --- Validation dropdown arrows ---
        emit_validation_arrows(params, layout, viewport, scale, dpr, hox, hoy, rects);

        // --- Selection overlay ---
        if let Some(sel) = params.selection {
            let sel_rects = selection::selection_rects(sel, layout, viewport);
            for sr in &sel_rects {
                let sx = sr.x as f32 * dpr + hox;
                let sy = sr.y as f32 * dpr + hoy;
                let sw = sr.w as f32 * dpr;
                let sh = sr.h as f32 * dpr;

                // Selection fill
                rects.push(RectInstance {
                    pos: [sx, sy],
                    size: [sw, sh],
                    color: SELECTION_FILL,
                });

                // Selection border (only on edges that are visible)
                let border_w = 2.0;
                if sr.draw_top {
                    lines.push(LineInstance {
                        start: [sx, sy],
                        end: [sx + sw, sy],
                        width_pad: [border_w, 0.0],
                        color: SELECTION_BORDER,
                    });
                }
                if sr.draw_bottom {
                    lines.push(LineInstance {
                        start: [sx, sy + sh],
                        end: [sx + sw, sy + sh],
                        width_pad: [border_w, 0.0],
                        color: SELECTION_BORDER,
                    });
                }
                if sr.draw_left {
                    lines.push(LineInstance {
                        start: [sx, sy],
                        end: [sx, sy + sh],
                        width_pad: [border_w, 0.0],
                        color: SELECTION_BORDER,
                    });
                }
                if sr.draw_right {
                    lines.push(LineInstance {
                        start: [sx + sw, sy],
                        end: [sx + sw, sy + sh],
                        width_pad: [border_w, 0.0],
                        color: SELECTION_BORDER,
                    });
                }
            }
        }

        // --- Frozen pane dividers ---
        if layout.frozen_cols > 0 {
            let div_x = frozen_width * scale + hox;
            lines.push(LineInstance {
                start: [div_x, hoy],
                end: [div_x, phys_h],
                width_pad: [2.0, 0.0],
                color: FROZEN_DIVIDER_COLOR,
            });
        }
        if layout.frozen_rows > 0 {
            let div_y = frozen_height * scale + hoy;
            lines.push(LineInstance {
                start: [hox, div_y],
                end: [phys_w, div_y],
                width_pad: [2.0, 0.0],
                color: FROZEN_DIVIDER_COLOR,
            });
        }

        // --- Upload projection ---
        let proj = orthographic_projection(phys_w, phys_h);
        self.queue.write_buffer(
            &self.globals_buffer,
            0,
            bytemuck::bytes_of(&Globals { projection: proj }),
        );

        // --- Upload instance data ---
        let rect_count = rects.len().min(MAX_RECTS);
        let line_count = lines.len().min(MAX_LINES);

        if rect_count > 0 {
            self.queue.write_buffer(
                &self.rect_buffer,
                0,
                bytemuck::cast_slice(rects.get(..rect_count).unwrap_or(&[])),
            );
        }
        if line_count > 0 {
            self.queue.write_buffer(
                &self.line_buffer,
                0,
                bytemuck::cast_slice(lines.get(..line_count).unwrap_or(&[])),
            );
        }

        // --- Render pass ---
        let output = self
            .surface
            .get_current_texture()
            .map_err(|e| format!("Failed to get surface texture: {e}"))?;
        let view = output
            .texture
            .create_view(&wgpu::TextureViewDescriptor::default());

        let mut encoder = self
            .device
            .create_command_encoder(&wgpu::CommandEncoderDescriptor {
                label: Some("xlview encoder"),
            });

        {
            let mut pass = encoder.begin_render_pass(&wgpu::RenderPassDescriptor {
                label: Some("xlview render pass"),
                color_attachments: &[Some(wgpu::RenderPassColorAttachment {
                    view: &view,
                    depth_slice: None,
                    resolve_target: None,
                    ops: wgpu::Operations {
                        load: wgpu::LoadOp::Clear(wgpu::Color::WHITE),
                        store: wgpu::StoreOp::Store,
                    },
                })],
                depth_stencil_attachment: None,
                timestamp_writes: None,
                occlusion_query_set: None,
                multiview_mask: None,
            });

            // Draw rects
            if rect_count > 0 {
                pass.set_pipeline(&self.rect_pipeline);
                pass.set_bind_group(0, &self.globals_bind_group, &[]);
                pass.set_vertex_buffer(0, self.rect_buffer.slice(..));
                #[allow(clippy::cast_possible_truncation)]
                pass.draw(0..6, 0..rect_count as u32);
            }

            // Draw lines
            if line_count > 0 {
                pass.set_pipeline(&self.line_pipeline);
                pass.set_bind_group(0, &self.globals_bind_group, &[]);
                pass.set_vertex_buffer(0, self.line_buffer.slice(..));
                #[allow(clippy::cast_possible_truncation)]
                pass.draw(0..6, 0..line_count as u32);
            }
        }

        self.queue.submit(std::iter::once(encoder.finish()));
        output.present();

        Ok(())
    }
}

/// Emit border line instances from a cell style (free function to avoid borrow conflicts).
fn emit_borders(
    style: &CellStyleData,
    px: f32,
    py: f32,
    pw: f32,
    ph: f32,
    dpr: f32,
    lines: &mut Vec<LineInstance>,
) {
    if let Some(ref border) = style.border_top {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px, py],
                end: [px + pw, py],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
    if let Some(ref border) = style.border_bottom {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px, py + ph],
                end: [px + pw, py + ph],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
    if let Some(ref border) = style.border_left {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px, py],
                end: [px, py + ph],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
    if let Some(ref border) = style.border_right {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px + pw, py],
                end: [px + pw, py + ph],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
    // Diagonal borders
    if let Some(ref border) = style.border_diagonal_down {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px, py],
                end: [px + pw, py + ph],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
    if let Some(ref border) = style.border_diagonal_up {
        if let Some(color) = border_color_f32(border) {
            #[allow(clippy::cast_possible_truncation)]
            let w = border.width() as f32 * dpr;
            lines.push(LineInstance {
                start: [px, py + ph],
                end: [px + pw, py],
                width_pad: [w, 0.0],
                color,
            });
        }
    }
}

/// Convert a border style's color to f32 RGBA.
fn border_color_f32(border: &crate::render::backend::BorderStyleData) -> Option<[f32; 4]> {
    let color_str = border.color.as_deref()?;
    // First try parse_color to normalize, then to f32
    let normalized = colors::parse_color(color_str)?;
    parse_color_f32(&normalized)
}

/// Generate column label: 0 → "A", 25 → "Z", 26 → "AA", etc.
fn col_label(col: u32) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        let ch = b'A' + (n % 26) as u8;
        result.insert(0, ch as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

/// Check if a column index falls within the selection range (for header highlighting).
fn is_col_in_selection(col: u32, sel: Option<&crate::types::Selection>) -> bool {
    if let Some(s) = sel {
        let min_col = s.start_col.min(s.end_col);
        let max_col = s.start_col.max(s.end_col);
        col >= min_col && col <= max_col
    } else {
        false
    }
}

/// Check if a row index falls within the selection range (for header highlighting).
fn is_row_in_selection(row: u32, sel: Option<&crate::types::Selection>) -> bool {
    if let Some(s) = sel {
        let min_row = s.start_row.min(s.end_row);
        let max_row = s.start_row.max(s.end_row);
        row >= min_row && row <= max_row
    } else {
        false
    }
}

/// Resolve pattern fill from cell style. Uses pattern_fg_color as a solid fill
/// when pattern_type is present and is not "none".
fn resolve_pattern_fill(style: Option<&CellStyleData>) -> Option<[f32; 4]> {
    let s = style?;
    let pattern_type = s.pattern_type.as_deref()?;
    if pattern_type == "none" {
        return None;
    }
    let fg_color = s.pattern_fg_color.as_deref()?;
    parse_color_f32(fg_color)
        .or_else(|| colors::parse_color(fg_color).as_deref().and_then(parse_color_f32))
}

/// CF color override for a cell.
struct CfOverride {
    bg_color: Option<String>,
    font_color: Option<String>,
}

/// Collect conditional formatting color overrides (color scales + cellIs DXF) for visible cells.
fn collect_cf_color_overrides(
    cells: &[CellRenderData],
    cf_rules: &[crate::types::ConditionalFormatting],
    dxf_styles: &[crate::types::DxfStyle],
    cf_cache: &[crate::types::ConditionalFormattingCache],
) -> std::collections::HashMap<(u32, u32), CfOverride> {
    let mut overrides = std::collections::HashMap::new();

    for (idx, cf) in cf_rules.iter().enumerate() {
        let ranges_storage;
        let ranges = if let Some(cache) = cf_cache.get(idx) {
            &cache.ranges
        } else {
            ranges_storage = parse_sqref(&cf.sqref);
            &ranges_storage
        };

        let sorted_rules: Vec<&crate::types::CFRule> = if let Some(cache) = cf_cache.get(idx) {
            cache
                .sorted_rule_indices
                .iter()
                .filter_map(|&i| cf.rules.get(i))
                .collect()
        } else {
            let mut rules: Vec<&crate::types::CFRule> = cf.rules.iter().collect();
            rules.sort_by_key(|r| r.priority);
            rules
        };

        // Collect numeric values for range-based rules
        let range_values = collect_range_values(cells, ranges);
        let (min_val, max_val) = get_min_max(&range_values);
        let range_span = max_val - min_val;

        for rule in &sorted_rules {
            // Color scale: interpolate bg color
            if let Some(ref color_scale) = rule.color_scale {
                for cell in cells {
                    if !cell_in_ranges(cell, ranges) {
                        continue;
                    }
                    let Some(value) = cell.numeric_value else {
                        continue;
                    };
                    let position = normalize_value(value, min_val, range_span);
                    let color_hex = interpolate_color(&color_scale.colors, position);
                    let entry = overrides
                        .entry((cell.row, cell.col))
                        .or_insert_with(|| CfOverride {
                            bg_color: None,
                            font_color: None,
                        });
                    entry.bg_color = Some(color_hex);
                }
            }
            // cellIs DXF: fill + font color overrides
            else if rule.rule_type == CFRuleType::CellIs {
                let dxf = match rule.dxf_id {
                    Some(id) => dxf_styles.get(id as usize),
                    None => continue,
                };
                let Some(dxf) = dxf else { continue };
                if dxf.fill_color.is_none() && dxf.font_color.is_none() {
                    continue;
                }

                let compare_value: Option<f64> =
                    rule.formula.as_ref().and_then(|f| f.parse().ok());
                let compare_str = rule.formula.as_deref();
                let operator = rule.operator.as_deref().unwrap_or("equal");

                for cell in cells {
                    if !cell_in_ranges(cell, ranges) {
                        continue;
                    }
                    if !evaluate_cell_is(cell, operator, compare_value, compare_str) {
                        continue;
                    }
                    let entry = overrides
                        .entry((cell.row, cell.col))
                        .or_insert_with(|| CfOverride {
                            bg_color: None,
                            font_color: None,
                        });
                    if let Some(ref fill) = dxf.fill_color {
                        entry.bg_color = Some(fill.clone());
                    }
                    if let Some(ref font) = dxf.font_color {
                        entry.font_color = Some(font.clone());
                    }
                }
            }
        }
    }

    overrides
}

/// Check if a cell is within any of the given ranges.
fn cell_in_ranges(cell: &CellRenderData, ranges: &[(u32, u32, u32, u32)]) -> bool {
    ranges
        .iter()
        .any(|(sr, sc, er, ec)| cell.row >= *sr && cell.row <= *er && cell.col >= *sc && cell.col <= *ec)
}

/// Collect numeric values from cells within ranges.
fn collect_range_values(
    cells: &[CellRenderData],
    ranges: &[(u32, u32, u32, u32)],
) -> Vec<f64> {
    let mut values = Vec::new();
    for cell in cells {
        if cell_in_ranges(cell, ranges) {
            if let Some(num) = cell.numeric_value {
                values.push(num);
            }
        }
    }
    values
}

/// Get min and max from a values collection.
fn get_min_max(values: &[f64]) -> (f64, f64) {
    let min = values.iter().copied().fold(f64::INFINITY, f64::min);
    let max = values.iter().copied().fold(f64::NEG_INFINITY, f64::max);
    (min, max)
}

/// Normalize a value into 0.0..1.0 range.
fn normalize_value(value: f64, min: f64, range_span: f64) -> f64 {
    if range_span <= 0.0 {
        0.5
    } else {
        ((value - min) / range_span).clamp(0.0, 1.0)
    }
}

/// Interpolate between colors based on position (0.0 to 1.0).
#[allow(
    clippy::indexing_slicing,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss
)]
fn interpolate_color(color_strings: &[String], position: f64) -> String {
    if color_strings.is_empty() {
        return "#FFFFFF".to_string();
    }
    if color_strings.len() == 1 {
        return color_strings[0].clone();
    }

    let position = position.clamp(0.0, 1.0);
    let num_segments = color_strings.len() - 1;
    let segment_size = 1.0 / num_segments as f64;

    let segment_index = (position / segment_size).floor() as usize;
    let segment_index = segment_index.min(num_segments - 1);

    let segment_pos = (position - (segment_index as f64 * segment_size)) / segment_size;
    let segment_pos = segment_pos.clamp(0.0, 1.0);

    let color1 = Rgb::from_hex(&color_strings[segment_index]).unwrap_or(Rgb::new(255, 255, 255));
    let color2 =
        Rgb::from_hex(&color_strings[segment_index + 1]).unwrap_or(Rgb::new(255, 255, 255));

    let r = lerp_u8(color1.r, color2.r, segment_pos);
    let g = lerp_u8(color1.g, color2.g, segment_pos);
    let b = lerp_u8(color1.b, color2.b, segment_pos);

    format!("#{r:02X}{g:02X}{b:02X}")
}

/// Linear interpolation for u8 values.
#[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
fn lerp_u8(a: u8, b: u8, t: f64) -> u8 {
    let a = f64::from(a);
    let b = f64::from(b);
    (a + (b - a) * t).round().clamp(0.0, 255.0) as u8
}

/// Evaluate a cellIs conditional formatting operator against a cell.
fn evaluate_cell_is(
    cell: &CellRenderData,
    operator: &str,
    compare_value: Option<f64>,
    compare_str: Option<&str>,
) -> bool {
    let cell_value = cell.numeric_value;
    let cell_str = cell.value.as_deref();

    match operator {
        "equal" => {
            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                (cv - cmp).abs() < f64::EPSILON
            } else {
                cell_str == compare_str
            }
        }
        "notEqual" => {
            if let (Some(cv), Some(cmp)) = (cell_value, compare_value) {
                (cv - cmp).abs() >= f64::EPSILON
            } else {
                cell_str != compare_str
            }
        }
        "greaterThan" => {
            matches!((cell_value, compare_value), (Some(cv), Some(cmp)) if cv > cmp)
        }
        "greaterThanOrEqual" => {
            matches!((cell_value, compare_value), (Some(cv), Some(cmp)) if cv >= cmp)
        }
        "lessThan" => {
            matches!((cell_value, compare_value), (Some(cv), Some(cmp)) if cv < cmp)
        }
        "lessThanOrEqual" => {
            matches!((cell_value, compare_value), (Some(cv), Some(cmp)) if cv <= cmp)
        }
        "containsText" => {
            matches!((cell_str, compare_str), (Some(cv), Some(cmp)) if cv.contains(cmp))
        }
        "notContainsText" | "notContains" => {
            matches!((cell_str, compare_str), (Some(cv), Some(cmp)) if !cv.contains(cmp))
        }
        "beginsWith" => {
            matches!((cell_str, compare_str), (Some(cv), Some(cmp)) if cv.starts_with(cmp))
        }
        "endsWith" => {
            matches!((cell_str, compare_str), (Some(cv), Some(cmp)) if cv.ends_with(cmp))
        }
        _ => false,
    }
}

/// Emit validation dropdown arrow rects for list-type validations with show_dropdown=true.
#[allow(clippy::too_many_arguments, clippy::cast_possible_truncation)]
fn emit_validation_arrows(
    params: &RenderParams,
    layout: &crate::layout::SheetLayout,
    viewport: &crate::layout::Viewport,
    scale: f32,
    dpr: f32,
    hox: f32,
    hoy: f32,
    rects: &mut Vec<RectInstance>,
) {
    let frozen_width = layout.frozen_cols_width();
    let frozen_height = layout.frozen_rows_height();

    for validation_range in params.data_validations {
        if !matches!(
            validation_range.validation.validation_type,
            ValidationType::List
        ) {
            continue;
        }
        if !validation_range.validation.show_dropdown {
            continue;
        }

        for range_str in validation_range.sqref.split_whitespace() {
            let Some((min_row, min_col, max_row, max_col)) = parse_cell_range(range_str) else {
                continue;
            };

            let (start_row, end_row) = viewport.visible_rows(layout);
            let (start_col, end_col) = viewport.visible_cols(layout);
            let row_start = min_row.max(start_row);
            let row_end = max_row.min(end_row.saturating_sub(1));
            let col_start = min_col.max(start_col);
            let col_end = max_col.min(end_col.saturating_sub(1));

            if row_start > row_end || col_start > col_end {
                continue;
            }

            for row in row_start..=row_end {
                for col in col_start..=col_end {
                    let col_x =
                        layout.col_positions.get(col as usize).copied().unwrap_or(0.0);
                    let col_w = layout.col_width(col);
                    let row_y =
                        layout.row_positions.get(row as usize).copied().unwrap_or(0.0);
                    let row_h = layout.row_height(row);

                    let is_col_frozen = col < layout.frozen_cols;
                    let is_row_frozen = row < layout.frozen_rows;
                    let sx = if is_col_frozen {
                        col_x * scale
                    } else {
                        frozen_width * scale + (col_x - viewport.scroll_x) * scale
                    };
                    let sy = if is_row_frozen {
                        row_y * scale
                    } else {
                        frozen_height * scale + (row_y - viewport.scroll_y) * scale
                    };

                    let px = sx + hox;
                    let py = sy + hoy;
                    let pw = col_w * scale;
                    let ph = row_h * scale;

                    // Small gray rect at right edge of cell as dropdown indicator
                    let arrow_w = 8.0 * dpr;
                    let arrow_h = 6.0 * dpr;
                    let arrow_x = px + pw - arrow_w - 4.0 * dpr;
                    let arrow_y = py + (ph - arrow_h) / 2.0;

                    rects.push(RectInstance {
                        pos: [arrow_x, arrow_y],
                        size: [arrow_w, arrow_h],
                        color: VALIDATION_ARROW_COLOR,
                    });
                }
            }
        }
    }
}

impl RenderBackend for WgpuRenderer {
    fn init(&mut self) -> Result<()> {
        Ok(())
    }

    fn resize(&mut self, width: u32, height: u32, dpr: f32) {
        let width = width.max(1);
        let height = height.max(1);
        self.width = width;
        self.height = height;
        self.dpr = dpr;

        self.surface_config.width = width;
        self.surface_config.height = height;
        self.surface.configure(&self.device, &self.surface_config);

        // Set canvas buffer size
        self.canvas.set_width(width);
        self.canvas.set_height(height);

        // Resize text overlay to match
        self.text_overlay.set_width(width);
        self.text_overlay.set_height(height);

        // Invalidate frame fingerprint on resize
        self.last_fingerprint = None;
    }

    fn render(&mut self, params: &RenderParams) -> Result<()> {
        self.render_frame(params)
    }

    fn width(&self) -> u32 {
        self.width
    }

    fn height(&self) -> u32 {
        self.height
    }
}

#[cfg(test)]
#[allow(clippy::unwrap_used, clippy::panic)]
mod tests {
    use super::*;

    #[test]
    fn col_label_a() {
        assert_eq!(col_label(0), "A");
    }

    #[test]
    fn col_label_z() {
        assert_eq!(col_label(25), "Z");
    }

    #[test]
    fn col_label_aa() {
        assert_eq!(col_label(26), "AA");
    }

    #[test]
    fn col_label_az() {
        assert_eq!(col_label(51), "AZ");
    }

    #[test]
    fn test_interpolate_color_single() {
        let colors = vec!["#FF0000".to_string()];
        assert_eq!(interpolate_color(&colors, 0.5), "#FF0000");
    }

    #[test]
    fn test_interpolate_color_two() {
        let colors = vec!["#000000".to_string(), "#FFFFFF".to_string()];
        let result = interpolate_color(&colors, 0.5);
        // Should be approximately #808080
        assert_eq!(result, "#808080");
    }

    #[test]
    fn test_interpolate_color_edges() {
        let colors = vec!["#FF0000".to_string(), "#0000FF".to_string()];
        assert_eq!(interpolate_color(&colors, 0.0), "#FF0000");
        assert_eq!(interpolate_color(&colors, 1.0), "#0000FF");
    }

    #[test]
    fn test_normalize_value() {
        assert!((normalize_value(50.0, 0.0, 100.0) - 0.5).abs() < f64::EPSILON);
        assert!((normalize_value(0.0, 0.0, 100.0) - 0.0).abs() < f64::EPSILON);
        assert!((normalize_value(100.0, 0.0, 100.0) - 1.0).abs() < f64::EPSILON);
        // Zero range -> 0.5
        assert!((normalize_value(5.0, 5.0, 0.0) - 0.5).abs() < f64::EPSILON);
    }

    #[test]
    fn test_evaluate_cell_is_equal() {
        let cell = CellRenderData {
            row: 0,
            col: 0,
            value: Some("10".to_string()),
            numeric_value: Some(10.0),
            style_idx: None,
            style_override: None,
            has_hyperlink: None,
            has_comment: None,
            rich_text: None,
        };
        assert!(evaluate_cell_is(&cell, "equal", Some(10.0), Some("10")));
        assert!(!evaluate_cell_is(&cell, "equal", Some(5.0), Some("5")));
        assert!(evaluate_cell_is(
            &cell,
            "greaterThan",
            Some(5.0),
            Some("5")
        ));
        assert!(!evaluate_cell_is(
            &cell,
            "lessThan",
            Some(5.0),
            Some("5")
        ));
    }

    #[test]
    fn test_is_col_in_selection() {
        let sel = crate::types::Selection::cell_range(1, 2, 3, 5);
        assert!(is_col_in_selection(2, Some(&sel)));
        assert!(is_col_in_selection(5, Some(&sel)));
        assert!(!is_col_in_selection(1, Some(&sel)));
        assert!(!is_col_in_selection(6, Some(&sel)));
        assert!(!is_col_in_selection(3, None));
    }

    #[test]
    fn test_is_row_in_selection() {
        let sel = crate::types::Selection::cell_range(1, 2, 3, 5);
        assert!(is_row_in_selection(1, Some(&sel)));
        assert!(is_row_in_selection(3, Some(&sel)));
        assert!(!is_row_in_selection(0, Some(&sel)));
        assert!(!is_row_in_selection(4, Some(&sel)));
        assert!(!is_row_in_selection(2, None));
    }

    #[test]
    fn test_color_to_css() {
        assert_eq!(color_to_css([1.0, 0.0, 0.0, 1.0]), "rgba(255,0,0,1)");
        assert_eq!(color_to_css([0.0, 0.0, 0.0, 0.5]), "rgba(0,0,0,0.5)");
    }

    #[test]
    fn test_build_font_string() {
        assert_eq!(
            build_font_string(14.0, "Arial", false, false),
            "14px Arial"
        );
        assert_eq!(
            build_font_string(12.0, "Calibri", true, true),
            "italic bold 12px Calibri"
        );
        assert_eq!(
            build_font_string(11.0, "Arial", true, false),
            "bold 11px Arial"
        );
    }
}

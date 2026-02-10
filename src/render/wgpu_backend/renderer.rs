//! wgpu renderer backend for xlview.
//!
//! Implements `RenderBackend` using WebGPU via the `wgpu` crate. Text is
//! rendered to an offscreen canvas and uploaded as a texture atlas.

use super::buffers::{
    orthographic_projection, Globals, LineInstance, RectInstance, TextInstance,
};
use super::pipelines;
use super::text_atlas::{TextAtlas, TextKey};
use crate::error::Result;
use crate::render::backend::{RenderBackend, RenderParams};
use crate::render::colors::{self, parse_color_f32};
use crate::render::selection;

use web_sys::HtmlCanvasElement;

/// Maximum instances per buffer. 128k instances should cover any reasonable
/// spreadsheet viewport.
const MAX_RECTS: usize = 131_072;
const MAX_LINES: usize = 262_144;
const MAX_TEXTS: usize = 131_072;

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
/// Header border color.
const HEADER_BORDER: [f32; 4] = [0.839, 0.839, 0.839, 1.0]; // #D6D6D6
/// Header text color.
const HEADER_TEXT_COLOR: [f32; 4] = [0.333, 0.333, 0.333, 1.0]; // #555
/// Frozen pane divider color.
const FROZEN_DIVIDER_COLOR: [f32; 4] = [0.6, 0.6, 0.6, 1.0]; // #999

/// The wgpu-based spreadsheet renderer.
pub struct WgpuRenderer {
    canvas: HtmlCanvasElement,
    device: wgpu::Device,
    queue: wgpu::Queue,
    surface: wgpu::Surface<'static>,
    surface_config: wgpu::SurfaceConfiguration,
    rect_pipeline: wgpu::RenderPipeline,
    line_pipeline: wgpu::RenderPipeline,
    text_pipeline: wgpu::RenderPipeline,
    globals_buffer: wgpu::Buffer,
    globals_bind_group: wgpu::BindGroup,
    text_bind_group: wgpu::BindGroup,
    text_atlas: TextAtlas,
    rect_buffer: wgpu::Buffer,
    line_buffer: wgpu::Buffer,
    text_buffer: wgpu::Buffer,
    width: u32,
    height: u32,
    dpr: f32,
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

        // Text atlas
        let text_atlas = TextAtlas::new(&device, &queue)
            .ok_or("Failed to create text atlas (no DOM access?)")?;

        // Text bind group layout: projection + texture + sampler
        let text_bind_group_layout =
            device.create_bind_group_layout(&wgpu::BindGroupLayoutDescriptor {
                label: Some("text bind group layout"),
                entries: &[
                    // Projection uniform
                    wgpu::BindGroupLayoutEntry {
                        binding: 0,
                        visibility: wgpu::ShaderStages::VERTEX,
                        ty: wgpu::BindingType::Buffer {
                            ty: wgpu::BufferBindingType::Uniform,
                            has_dynamic_offset: false,
                            min_binding_size: None,
                        },
                        count: None,
                    },
                    // Atlas texture
                    wgpu::BindGroupLayoutEntry {
                        binding: 1,
                        visibility: wgpu::ShaderStages::FRAGMENT,
                        ty: wgpu::BindingType::Texture {
                            sample_type: wgpu::TextureSampleType::Float { filterable: true },
                            view_dimension: wgpu::TextureViewDimension::D2,
                            multisampled: false,
                        },
                        count: None,
                    },
                    // Sampler
                    wgpu::BindGroupLayoutEntry {
                        binding: 2,
                        visibility: wgpu::ShaderStages::FRAGMENT,
                        ty: wgpu::BindingType::Sampler(wgpu::SamplerBindingType::Filtering),
                        count: None,
                    },
                ],
            });

        let text_bind_group = device.create_bind_group(&wgpu::BindGroupDescriptor {
            label: Some("text bind group"),
            layout: &text_bind_group_layout,
            entries: &[
                wgpu::BindGroupEntry {
                    binding: 0,
                    resource: globals_buffer.as_entire_binding(),
                },
                wgpu::BindGroupEntry {
                    binding: 1,
                    resource: wgpu::BindingResource::TextureView(&text_atlas.view),
                },
                wgpu::BindGroupEntry {
                    binding: 2,
                    resource: wgpu::BindingResource::Sampler(&text_atlas.sampler),
                },
            ],
        });

        // Pipelines
        let rect_pipeline = pipelines::create_rect_pipeline(&device, format, &globals_layout);
        let line_pipeline = pipelines::create_line_pipeline(&device, format, &globals_layout);
        let text_pipeline =
            pipelines::create_text_pipeline(&device, format, &text_bind_group_layout);

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
        let text_buffer = device.create_buffer(&wgpu::BufferDescriptor {
            label: Some("text instances"),
            size: (MAX_TEXTS * std::mem::size_of::<TextInstance>()) as u64,
            usage: wgpu::BufferUsages::VERTEX | wgpu::BufferUsages::COPY_DST,
            mapped_at_creation: false,
        });

        Ok(Self {
            canvas,
            device,
            queue,
            surface,
            surface_config,
            rect_pipeline,
            line_pipeline,
            text_pipeline,
            globals_buffer,
            globals_bind_group,
            text_bind_group,
            text_atlas,
            rect_buffer,
            line_buffer,
            text_buffer,
            width,
            height,
            dpr,
        })
    }

    /// Set CSS dimensions on the canvas element (logical pixels).
    pub fn set_canvas_css_size(&self, css_w: f32, css_h: f32) {
        let style = self.canvas.style();
        let _ = style.set_property("width", &format!("{css_w}px"));
        let _ = style.set_property("height", &format!("{css_h}px"));
    }

    /// Collect geometry from render params and issue draw calls.
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        clippy::too_many_lines
    )]
    fn render_frame(&mut self, params: &RenderParams) -> Result<()> {
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

        let mut rects = Vec::with_capacity(params.cells.len());
        let mut lines: Vec<LineInstance> = Vec::with_capacity(params.cells.len() * 2);
        let mut texts: Vec<TextInstance> = Vec::new();

        let frozen_width = layout.frozen_cols_width();
        let frozen_height = layout.frozen_rows_height();
        let scale = viewport.scale * self.dpr;

        // --- Headers ---
        if params.show_headers {
            let header_w = params.header_config.row_header_width * self.dpr;
            let header_h = params.header_config.col_header_height * self.dpr;

            // Top-left corner
            rects.push(RectInstance {
                pos: [0.0, 0.0],
                size: [header_w, header_h],
                color: HEADER_BG,
            });

            // Column headers
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

                rects.push(RectInstance {
                    pos: [px, 0.0],
                    size: [pw, header_h],
                    color: HEADER_BG,
                });
                // Column header border
                lines.push(LineInstance {
                    start: [px + pw, 0.0],
                    end: [px + pw, header_h],
                    width_pad: [1.0, 0.0],
                    color: HEADER_BORDER,
                });
                // Column label text
                let label = col_label(col_idx);
                let key = TextKey::new(&label, "Arial", 10.0 * self.dpr, false, false);
                if let Some(entry) = self.text_atlas.get_or_insert(&key) {
                    let text_x = px + (pw - entry.pixel_width) / 2.0;
                    let text_y = (header_h - entry.pixel_height) / 2.0;
                    texts.push(TextInstance {
                        pos: [text_x, text_y],
                        size: [entry.pixel_width, entry.pixel_height],
                        uv_pos: entry.uv_pos,
                        uv_size: entry.uv_size,
                        color: HEADER_TEXT_COLOR,
                    });
                }
            }

            // Row headers
            let (start_row, end_row) = viewport.visible_rows(layout);
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

                rects.push(RectInstance {
                    pos: [0.0, py],
                    size: [header_w, ph],
                    color: HEADER_BG,
                });
                // Row header border
                lines.push(LineInstance {
                    start: [0.0, py + ph],
                    end: [header_w, py + ph],
                    width_pad: [1.0, 0.0],
                    color: HEADER_BORDER,
                });
                // Row label text
                let label = format!("{}", row_idx + 1);
                let key = TextKey::new(&label, "Arial", 10.0 * self.dpr, false, false);
                if let Some(entry) = self.text_atlas.get_or_insert(&key) {
                    let text_x = (header_w - entry.pixel_width) / 2.0;
                    let text_y = py + (ph - entry.pixel_height) / 2.0;
                    texts.push(TextInstance {
                        pos: [text_x, text_y],
                        size: [entry.pixel_width, entry.pixel_height],
                        uv_pos: entry.uv_pos,
                        uv_size: entry.uv_size,
                        color: HEADER_TEXT_COLOR,
                    });
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

        // --- Cells: backgrounds, grid lines, text ---
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

            // Cell background
            let bg = style
                .and_then(|s| s.bg_color.as_deref())
                .and_then(parse_color_f32)
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

            // Borders (override grid lines)
            if let Some(s) = style {
                self.emit_borders(s, px, py, pw, ph, &mut lines);
            }

            // Cell text
            if let Some(value) = cell.value.as_deref() {
                if !value.is_empty() {
                    let font_size = style.and_then(|s| s.font_size).unwrap_or(11.0) * self.dpr;
                    let font_family = style
                        .and_then(|s| s.font_family.as_deref())
                        .or(params.minor_font)
                        .unwrap_or("Arial");
                    let bold = style.and_then(|s| s.bold).unwrap_or(false);
                    let italic = style.and_then(|s| s.italic).unwrap_or(false);
                    let text_color = style
                        .and_then(|s| s.font_color.as_deref())
                        .and_then(parse_color_f32)
                        .unwrap_or(DEFAULT_TEXT_COLOR);

                    let key = TextKey::new(value, font_family, font_size, bold, italic);
                    if let Some(entry) = self.text_atlas.get_or_insert(&key) {
                        // Horizontal alignment
                        let align = style.and_then(|s| s.align_h.as_deref());
                        let padding = 3.0 * self.dpr;
                        let text_x = match align {
                            Some("center") | Some("centerContinuous") => {
                                px + (pw - entry.pixel_width) / 2.0
                            }
                            Some("right") => px + pw - entry.pixel_width - padding,
                            _ => {
                                // Auto: right-align numbers, left-align text
                                if cell.numeric_value.is_some() {
                                    px + pw - entry.pixel_width - padding
                                } else {
                                    px + padding
                                }
                            }
                        };
                        // Vertical center
                        let text_y = py + (ph - entry.pixel_height) / 2.0;

                        texts.push(TextInstance {
                            pos: [text_x, text_y],
                            size: [entry.pixel_width, entry.pixel_height],
                            uv_pos: entry.uv_pos,
                            uv_size: entry.uv_size,
                            color: text_color,
                        });
                    }
                }
            }
        }

        // --- Selection overlay ---
        if let Some(sel) = params.selection {
            let sel_rects = selection::selection_rects(sel, layout, viewport);
            for sr in &sel_rects {
                let sx = sr.x as f32 * self.dpr + hox;
                let sy = sr.y as f32 * self.dpr + hoy;
                let sw = sr.w as f32 * self.dpr;
                let sh = sr.h as f32 * self.dpr;

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

        // --- Upload text atlas ---
        self.text_atlas.upload_if_dirty(&self.queue);

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
        let text_count = texts.len().min(MAX_TEXTS);

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
        if text_count > 0 {
            self.queue.write_buffer(
                &self.text_buffer,
                0,
                bytemuck::cast_slice(texts.get(..text_count).unwrap_or(&[])),
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

            // Draw text
            if text_count > 0 {
                pass.set_pipeline(&self.text_pipeline);
                pass.set_bind_group(0, &self.text_bind_group, &[]);
                pass.set_vertex_buffer(0, self.text_buffer.slice(..));
                #[allow(clippy::cast_possible_truncation)]
                pass.draw(0..6, 0..text_count as u32);
            }
        }

        self.queue.submit(std::iter::once(encoder.finish()));
        output.present();

        Ok(())
    }

    /// Emit border line instances from a cell style.
    fn emit_borders(
        &self,
        style: &crate::render::backend::CellStyleData,
        px: f32,
        py: f32,
        pw: f32,
        ph: f32,
        lines: &mut Vec<LineInstance>,
    ) {
        if let Some(ref border) = style.border_top {
            if let Some(color) = border_color_f32(border) {
                #[allow(clippy::cast_possible_truncation)]
                let w = border.width() as f32 * self.dpr;
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
                let w = border.width() as f32 * self.dpr;
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
                let w = border.width() as f32 * self.dpr;
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
                let w = border.width() as f32 * self.dpr;
                lines.push(LineInstance {
                    start: [px + pw, py],
                    end: [px + pw, py + ph],
                    width_pad: [w, 0.0],
                    color,
                });
            }
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
}

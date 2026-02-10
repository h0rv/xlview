//! Text atlas for the wgpu renderer.
//!
//! Renders text to an offscreen `<canvas>` using the browser's native `fillText()`,
//! then uploads the canvas pixels to a wgpu texture. This reuses the browser's full
//! font pipeline (CJK, ligatures, kerning, web fonts) without shipping a font
//! rasterizer in WASM.

use std::collections::HashMap;
use wasm_bindgen::JsCast;
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement};

/// Atlas texture size in pixels. 2048×2048 gives ~4 MB of RGBA data and can
/// hold thousands of text entries at typical spreadsheet font sizes.
const ATLAS_SIZE: u32 = 2048;

/// Padding between atlas entries to avoid texture bleeding.
const ENTRY_PAD: u32 = 2;

/// Key for looking up cached text entries in the atlas.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct TextKey {
    pub text: String,
    pub font_family: String,
    /// Font size quantized to quarter-pixels to avoid cache explosion.
    pub size_qpx: u32,
    pub bold: bool,
    pub italic: bool,
}

impl TextKey {
    pub fn new(text: &str, font_family: &str, size_px: f32, bold: bool, italic: bool) -> Self {
        #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
        let size_qpx = (size_px * 4.0).round() as u32;
        Self {
            text: text.to_string(),
            font_family: font_family.to_string(),
            size_qpx,
            bold,
            italic,
        }
    }

    fn size_px(&self) -> f32 {
        self.size_qpx as f32 / 4.0
    }

    fn css_font(&self) -> String {
        let style = if self.italic { "italic " } else { "" };
        let weight = if self.bold { "bold " } else { "" };
        format!("{}{}{}px {}", style, weight, self.size_px(), self.font_family)
    }
}

/// Cached atlas entry describing where a rendered text string lives in the atlas.
#[derive(Clone, Debug)]
pub struct TextEntry {
    /// UV position in normalized [0,1] coordinates.
    pub uv_pos: [f32; 2],
    /// UV size in normalized [0,1] coordinates.
    pub uv_size: [f32; 2],
    /// Pixel dimensions of the rendered text.
    pub pixel_width: f32,
    pub pixel_height: f32,
}

/// Text atlas backed by an offscreen canvas and a wgpu texture.
pub struct TextAtlas {
    /// Offscreen canvas for rendering text (kept alive so ctx remains valid).
    _canvas: HtmlCanvasElement,
    ctx: CanvasRenderingContext2d,
    /// GPU texture holding the atlas.
    pub texture: wgpu::Texture,
    pub view: wgpu::TextureView,
    pub sampler: wgpu::Sampler,
    /// Entry cache: text key → atlas location.
    entries: HashMap<TextKey, TextEntry>,
    /// Row-based packing: current x cursor.
    cursor_x: u32,
    /// Current row y position.
    cursor_y: u32,
    /// Current row height (max height in current row).
    row_height: u32,
    /// Whether the canvas has been modified since the last GPU upload.
    dirty: bool,
}

impl TextAtlas {
    /// Create a new text atlas.
    pub fn new(device: &wgpu::Device, queue: &wgpu::Queue) -> Option<Self> {
        let document = web_sys::window()?.document()?;
        let canvas: HtmlCanvasElement = document
            .create_element("canvas")
            .ok()?
            .dyn_into()
            .ok()?;
        canvas.set_width(ATLAS_SIZE);
        canvas.set_height(ATLAS_SIZE);

        let ctx: CanvasRenderingContext2d = canvas
            .get_context("2d")
            .ok()??
            .dyn_into()
            .ok()?;

        let texture = device.create_texture(&wgpu::TextureDescriptor {
            label: Some("text atlas"),
            size: wgpu::Extent3d {
                width: ATLAS_SIZE,
                height: ATLAS_SIZE,
                depth_or_array_layers: 1,
            },
            mip_level_count: 1,
            sample_count: 1,
            dimension: wgpu::TextureDimension::D2,
            format: wgpu::TextureFormat::Rgba8UnormSrgb,
            usage: wgpu::TextureUsages::TEXTURE_BINDING | wgpu::TextureUsages::COPY_DST,
            view_formats: &[],
        });

        // Initialize with transparent pixels
        let zeros = vec![0u8; (ATLAS_SIZE * ATLAS_SIZE * 4) as usize];
        queue.write_texture(
            wgpu::TexelCopyTextureInfo {
                texture: &texture,
                mip_level: 0,
                origin: wgpu::Origin3d::ZERO,
                aspect: wgpu::TextureAspect::All,
            },
            &zeros,
            wgpu::TexelCopyBufferLayout {
                offset: 0,
                bytes_per_row: Some(ATLAS_SIZE * 4),
                rows_per_image: Some(ATLAS_SIZE),
            },
            wgpu::Extent3d {
                width: ATLAS_SIZE,
                height: ATLAS_SIZE,
                depth_or_array_layers: 1,
            },
        );

        let view = texture.create_view(&wgpu::TextureViewDescriptor::default());
        let sampler = device.create_sampler(&wgpu::SamplerDescriptor {
            label: Some("text atlas sampler"),
            mag_filter: wgpu::FilterMode::Linear,
            min_filter: wgpu::FilterMode::Linear,
            ..Default::default()
        });

        Some(Self {
            _canvas: canvas,
            ctx,
            texture,
            view,
            sampler,
            entries: HashMap::new(),
            cursor_x: 0,
            cursor_y: 0,
            row_height: 0,
            dirty: false,
        })
    }

    /// Look up or render text into the atlas. Returns the entry if successful.
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
    pub fn get_or_insert(&mut self, key: &TextKey) -> Option<TextEntry> {
        if let Some(entry) = self.entries.get(key) {
            return Some(entry.clone());
        }

        // Measure the text
        let font = key.css_font();
        self.ctx.set_font(&font);
        let metrics = self.ctx.measure_text(&key.text).ok()?;
        let text_width = metrics.width();
        let text_height = f64::from(key.size_px()) * 1.3; // approximate line height

        let w = (text_width.ceil() as u32).max(1) + ENTRY_PAD;
        let h = (text_height.ceil() as u32).max(1) + ENTRY_PAD;

        // Check if we need to wrap to the next row
        if self.cursor_x + w > ATLAS_SIZE {
            self.cursor_y += self.row_height + ENTRY_PAD;
            self.cursor_x = 0;
            self.row_height = 0;
        }

        // Check if atlas is full
        if self.cursor_y + h > ATLAS_SIZE {
            // Reset the atlas entirely
            self.clear();
            // Try again after clear
            if h > ATLAS_SIZE {
                return None; // Text too tall even for empty atlas
            }
        }

        let x = self.cursor_x;
        let y = self.cursor_y;

        // Render text to canvas
        self.ctx.set_font(&font);
        self.ctx.set_fill_style_str("white");
        self.ctx.set_text_baseline("top");
        self.ctx
            .fill_text(&key.text, f64::from(x), f64::from(y) + f64::from(ENTRY_PAD) / 2.0)
            .ok()?;

        // Update packing state
        self.cursor_x = x + w;
        if h > self.row_height {
            self.row_height = h;
        }
        self.dirty = true;

        let atlas_f = ATLAS_SIZE as f32;
        let entry = TextEntry {
            uv_pos: [x as f32 / atlas_f, y as f32 / atlas_f],
            uv_size: [(w - ENTRY_PAD) as f32 / atlas_f, (h - ENTRY_PAD) as f32 / atlas_f],
            pixel_width: (w - ENTRY_PAD) as f32,
            pixel_height: (h - ENTRY_PAD) as f32,
        };

        self.entries.insert(key.clone(), entry.clone());
        Some(entry)
    }

    /// Upload modified atlas pixels to the GPU texture.
    pub fn upload_if_dirty(&mut self, queue: &wgpu::Queue) {
        if !self.dirty {
            return;
        }
        self.dirty = false;

        // Read pixel data from the canvas
        let Ok(image_data) = self.ctx.get_image_data(
            0.0,
            0.0,
            f64::from(ATLAS_SIZE),
            f64::from(ATLAS_SIZE),
        ) else {
            return;
        };
        let data = image_data.data();

        queue.write_texture(
            wgpu::TexelCopyTextureInfo {
                texture: &self.texture,
                mip_level: 0,
                origin: wgpu::Origin3d::ZERO,
                aspect: wgpu::TextureAspect::All,
            },
            &data,
            wgpu::TexelCopyBufferLayout {
                offset: 0,
                bytes_per_row: Some(ATLAS_SIZE * 4),
                rows_per_image: Some(ATLAS_SIZE),
            },
            wgpu::Extent3d {
                width: ATLAS_SIZE,
                height: ATLAS_SIZE,
                depth_or_array_layers: 1,
            },
        );
    }

    /// Clear all atlas entries and reset the canvas.
    fn clear(&mut self) {
        self.entries.clear();
        self.cursor_x = 0;
        self.cursor_y = 0;
        self.row_height = 0;
        self.ctx.clear_rect(
            0.0,
            0.0,
            f64::from(ATLAS_SIZE),
            f64::from(ATLAS_SIZE),
        );
        self.dirty = true;
    }
}

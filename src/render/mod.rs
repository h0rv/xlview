//! Rendering engine with pluggable backends.
//!
//! This module provides:
//! - Backend-agnostic rendering traits and types
//! - Canvas 2D backend (primary, stable)
//! - wgpu/WebGPU backend (optional, via `wgpu-backend` feature)
//! - Color parsing utilities

pub mod backend;
pub mod blit;
pub mod cache;
pub mod canvas;
pub mod colors;
pub mod selection;

#[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
pub mod wgpu_backend;

// Re-export commonly used types
pub use crate::types::TextRunData;
pub use backend::{BorderStyleData, CellRenderData, CellStyleData, RenderBackend, RenderParams};
pub use canvas::CanvasRenderer;
pub use colors::{palette, parse_color, CssColor};

#[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
pub use wgpu_backend::WgpuRenderer;

use crate::error::Result;

/// Renderer enum wrapping available backends for runtime switching.
#[allow(clippy::large_enum_variant)]
pub enum Renderer {
    /// Canvas 2D backend (default, stable).
    Canvas(CanvasRenderer),
    /// wgpu/WebGPU backend (optional).
    #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
    Wgpu(Box<wgpu_backend::WgpuRenderer>),
}

impl Renderer {
    /// Delegate `init()` to the active backend.
    pub fn init(&mut self) -> Result<()> {
        match self {
            Self::Canvas(r) => r.init(),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.init(),
        }
    }

    /// Delegate `resize()` to the active backend.
    pub fn resize(&mut self, width: u32, height: u32, dpr: f32) {
        match self {
            Self::Canvas(r) => r.resize(width, height, dpr),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.resize(width, height, dpr),
        }
    }

    /// Full render (calls the backend's `render()`).
    pub fn render(&mut self, params: &RenderParams) -> Result<()> {
        match self {
            Self::Canvas(r) => r.render(params),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.render(params),
        }
    }

    /// Render base layer only (cells, no overlay). For wgpu, this does a full render.
    pub fn render_base(&mut self, params: &RenderParams) -> Result<()> {
        match self {
            Self::Canvas(r) => r.render_base(params),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.render(params),
        }
    }

    /// Render overlay layer only (selection, headers, frozen dividers).
    /// For wgpu, this is a no-op since everything is rendered in a single pass.
    pub fn render_overlay(&mut self, params: &RenderParams) -> Result<()> {
        match self {
            Self::Canvas(r) => r.render_overlay(params),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(_) => {
                let _ = params;
                Ok(())
            }
        }
    }

    /// Whether the renderer has deferred tile prefetch work pending.
    /// Always false for wgpu (no tile caching).
    pub fn has_deferred_prefetch_tiles(&self) -> bool {
        match self {
            Self::Canvas(r) => r.has_deferred_prefetch_tiles(),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(_) => false,
        }
    }

    /// Set the CSS size of the canvas element (logical pixels).
    pub fn set_canvas_css_size(&self, css_w: f32, css_h: f32) {
        match self {
            Self::Canvas(r) => r.set_canvas_css_size(css_w, css_h),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.set_canvas_css_size(css_w, css_h),
        }
    }

    /// Get current width.
    pub fn width(&self) -> u32 {
        match self {
            Self::Canvas(r) => r.width(),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.width(),
        }
    }

    /// Get current height.
    pub fn height(&self) -> u32 {
        match self {
            Self::Canvas(r) => r.height(),
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(r) => r.height(),
        }
    }

    /// Returns true if this is the wgpu backend.
    pub fn is_wgpu(&self) -> bool {
        match self {
            Self::Canvas(_) => false,
            #[cfg(all(feature = "wgpu-backend", target_arch = "wasm32"))]
            Self::Wgpu(_) => true,
        }
    }
}

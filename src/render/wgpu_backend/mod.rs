//! wgpu (WebGPU) rendering backend.
//!
//! This module provides spreadsheet rendering using the WebGPU API via the
//! `wgpu` crate. Text is rendered through an offscreen canvas text atlas
//! to reuse the browser's native font pipeline.

pub mod buffers;
pub mod pipelines;
pub mod renderer;
pub mod text_atlas;

pub use renderer::WgpuRenderer;

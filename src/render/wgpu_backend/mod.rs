//! wgpu (WebGPU) rendering backend.
//!
//! This module provides spreadsheet rendering using the WebGPU API via the
//! `wgpu` crate. Text is rendered on a transparent Canvas 2D overlay
//! positioned over the WebGPU canvas, reusing the browser's native font
//! pipeline for perfect text quality.

pub mod buffers;
pub mod pipelines;
pub mod renderer;

pub use renderer::WgpuRenderer;

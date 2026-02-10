//! GPU buffer types for the wgpu renderer.
//!
//! All types are `#[repr(C)]` + `bytemuck::Pod + Zeroable` for safe GPU upload.

use bytemuck::{Pod, Zeroable};

/// Global uniform: orthographic projection matrix.
#[repr(C)]
#[derive(Copy, Clone, Debug, Pod, Zeroable)]
pub struct Globals {
    pub projection: [[f32; 4]; 4],
}

/// Per-instance data for filled rectangles (cell backgrounds, selection, headers).
#[repr(C)]
#[derive(Copy, Clone, Debug, Pod, Zeroable)]
pub struct RectInstance {
    pub pos: [f32; 2],
    pub size: [f32; 2],
    pub color: [f32; 4],
}

/// Per-instance data for lines (grid lines, borders).
#[repr(C)]
#[derive(Copy, Clone, Debug, Pod, Zeroable)]
pub struct LineInstance {
    pub start: [f32; 2],
    pub end: [f32; 2],
    pub width_pad: [f32; 2],
    pub color: [f32; 4],
}

/// Build an orthographic projection matrix mapping pixel coordinates to clip space.
///
/// Maps `(0,0)` at top-left to `(width, height)` at bottom-right,
/// with z in `[0, 1]`.
pub fn orthographic_projection(width: f32, height: f32) -> [[f32; 4]; 4] {
    // NDC: x [-1, 1], y [-1, 1] (top = +1 in clip but we want top = 0 in pixels)
    let sx = 2.0 / width;
    let sy = -2.0 / height; // flip y: pixel y increases downward
    [
        [sx, 0.0, 0.0, 0.0],
        [0.0, sy, 0.0, 0.0],
        [0.0, 0.0, 1.0, 0.0],
        [-1.0, 1.0, 0.0, 1.0],
    ]
}

impl RectInstance {
    /// Vertex buffer layout for instanced attributes.
    pub fn layout() -> wgpu::VertexBufferLayout<'static> {
        wgpu::VertexBufferLayout {
            array_stride: std::mem::size_of::<Self>() as wgpu::BufferAddress,
            step_mode: wgpu::VertexStepMode::Instance,
            attributes: &[
                // pos
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x2,
                    offset: 0,
                    shader_location: 0,
                },
                // size
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x2,
                    offset: 8,
                    shader_location: 1,
                },
                // color
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x4,
                    offset: 16,
                    shader_location: 2,
                },
            ],
        }
    }
}

impl LineInstance {
    /// Vertex buffer layout for instanced attributes.
    pub fn layout() -> wgpu::VertexBufferLayout<'static> {
        wgpu::VertexBufferLayout {
            array_stride: std::mem::size_of::<Self>() as wgpu::BufferAddress,
            step_mode: wgpu::VertexStepMode::Instance,
            attributes: &[
                // start
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x2,
                    offset: 0,
                    shader_location: 0,
                },
                // end
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x2,
                    offset: 8,
                    shader_location: 1,
                },
                // width_pad
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x2,
                    offset: 16,
                    shader_location: 2,
                },
                // color
                wgpu::VertexAttribute {
                    format: wgpu::VertexFormat::Float32x4,
                    offset: 24,
                    shader_location: 3,
                },
            ],
        }
    }
}

#[cfg(test)]
#[allow(clippy::float_cmp)]
mod tests {
    use super::*;

    #[test]
    fn orthographic_maps_origin_to_top_left() {
        let proj = orthographic_projection(800.0, 600.0);
        // (0,0) should map to (-1, 1) in clip space
        let x = proj[0][0] * 0.0 + proj[3][0]; // sx*0 + (-1) = -1
        let y = proj[1][1] * 0.0 + proj[3][1]; // sy*0 + 1 = 1
        assert_eq!(x, -1.0);
        assert_eq!(y, 1.0);
    }

    #[test]
    fn orthographic_maps_bottom_right() {
        let proj = orthographic_projection(800.0, 600.0);
        // (800, 600) should map to (1, -1) in clip space
        let x = proj[0][0] * 800.0 + proj[3][0]; // 2/800*800 + (-1) = 1
        let y = proj[1][1] * 600.0 + proj[3][1]; // -2/600*600 + 1 = -1
        assert_eq!(x, 1.0);
        assert_eq!(y, -1.0);
    }
}

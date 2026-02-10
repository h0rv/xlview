//! Render pipeline creation for the wgpu backend.

use super::buffers::{LineInstance, RectInstance};

/// Create the rectangle render pipeline (instanced quads for cell backgrounds, etc.).
pub fn create_rect_pipeline(
    device: &wgpu::Device,
    format: wgpu::TextureFormat,
    globals_layout: &wgpu::BindGroupLayout,
) -> wgpu::RenderPipeline {
    let shader = device.create_shader_module(wgpu::ShaderModuleDescriptor {
        label: Some("rect shader"),
        source: wgpu::ShaderSource::Wgsl(include_str!("shaders/rect.wgsl").into()),
    });

    device.create_render_pipeline(&wgpu::RenderPipelineDescriptor {
        label: Some("rect pipeline"),
        layout: Some(
            &device.create_pipeline_layout(&wgpu::PipelineLayoutDescriptor {
                label: Some("rect pipeline layout"),
                bind_group_layouts: &[globals_layout],
                immediate_size: 0,
            }),
        ),
        vertex: wgpu::VertexState {
            module: &shader,
            entry_point: Some("vs_main"),
            buffers: &[RectInstance::layout()],
            compilation_options: wgpu::PipelineCompilationOptions::default(),
        },
        fragment: Some(wgpu::FragmentState {
            module: &shader,
            entry_point: Some("fs_main"),
            targets: &[Some(wgpu::ColorTargetState {
                format,
                blend: Some(wgpu::BlendState::ALPHA_BLENDING),
                write_mask: wgpu::ColorWrites::ALL,
            })],
            compilation_options: wgpu::PipelineCompilationOptions::default(),
        }),
        primitive: wgpu::PrimitiveState {
            topology: wgpu::PrimitiveTopology::TriangleList,
            ..Default::default()
        },
        depth_stencil: None,
        multisample: wgpu::MultisampleState::default(),
        multiview_mask: None,
        cache: None,
    })
}

/// Create the line render pipeline (thin quads for grid lines, borders).
pub fn create_line_pipeline(
    device: &wgpu::Device,
    format: wgpu::TextureFormat,
    globals_layout: &wgpu::BindGroupLayout,
) -> wgpu::RenderPipeline {
    let shader = device.create_shader_module(wgpu::ShaderModuleDescriptor {
        label: Some("line shader"),
        source: wgpu::ShaderSource::Wgsl(include_str!("shaders/line.wgsl").into()),
    });

    device.create_render_pipeline(&wgpu::RenderPipelineDescriptor {
        label: Some("line pipeline"),
        layout: Some(
            &device.create_pipeline_layout(&wgpu::PipelineLayoutDescriptor {
                label: Some("line pipeline layout"),
                bind_group_layouts: &[globals_layout],
                immediate_size: 0,
            }),
        ),
        vertex: wgpu::VertexState {
            module: &shader,
            entry_point: Some("vs_main"),
            buffers: &[LineInstance::layout()],
            compilation_options: wgpu::PipelineCompilationOptions::default(),
        },
        fragment: Some(wgpu::FragmentState {
            module: &shader,
            entry_point: Some("fs_main"),
            targets: &[Some(wgpu::ColorTargetState {
                format,
                blend: Some(wgpu::BlendState::ALPHA_BLENDING),
                write_mask: wgpu::ColorWrites::ALL,
            })],
            compilation_options: wgpu::PipelineCompilationOptions::default(),
        }),
        primitive: wgpu::PrimitiveState {
            topology: wgpu::PrimitiveTopology::TriangleList,
            ..Default::default()
        },
        depth_stencil: None,
        multisample: wgpu::MultisampleState::default(),
        multiview_mask: None,
        cache: None,
    })
}


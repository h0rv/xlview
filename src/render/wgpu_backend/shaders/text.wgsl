// Instanced textured-quad shader for text rendering.
//
// Each instance provides position, size, UV coordinates into the text atlas,
// and a text color. The atlas texture stores pre-multiplied alpha glyphs.

struct Globals {
    projection: mat4x4<f32>,
};

@group(0) @binding(0)
var<uniform> globals: Globals;

@group(0) @binding(1)
var atlas_texture: texture_2d<f32>;

@group(0) @binding(2)
var atlas_sampler: sampler;

struct TextInstance {
    @location(0) pos: vec2<f32>,
    @location(1) size: vec2<f32>,
    @location(2) uv_pos: vec2<f32>,
    @location(3) uv_size: vec2<f32>,
    @location(4) color: vec4<f32>,
};

struct VertexOutput {
    @builtin(position) position: vec4<f32>,
    @location(0) uv: vec2<f32>,
    @location(1) color: vec4<f32>,
};

@vertex
fn vs_main(@builtin(vertex_index) vi: u32, inst: TextInstance) -> VertexOutput {
    var corners = array<vec2<f32>, 6>(
        vec2<f32>(0.0, 0.0),
        vec2<f32>(1.0, 0.0),
        vec2<f32>(0.0, 1.0),
        vec2<f32>(0.0, 1.0),
        vec2<f32>(1.0, 0.0),
        vec2<f32>(1.0, 1.0),
    );
    let corner = corners[vi];
    let world_pos = inst.pos + corner * inst.size;
    let uv = inst.uv_pos + corner * inst.uv_size;

    var out: VertexOutput;
    out.position = globals.projection * vec4<f32>(world_pos, 0.0, 1.0);
    out.uv = uv;
    out.color = inst.color;
    return out;
}

@fragment
fn fs_main(in: VertexOutput) -> @location(0) vec4<f32> {
    let tex = textureSample(atlas_texture, atlas_sampler, in.uv);
    // Atlas stores white text on transparent background.
    // Use the atlas alpha for coverage, multiply by desired text color.
    return vec4<f32>(in.color.rgb, in.color.a * tex.a);
}

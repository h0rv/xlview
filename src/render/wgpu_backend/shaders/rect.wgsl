// Instanced rectangle shader.
//
// Each instance provides position, size, and color.
// A unit quad (two triangles, 6 vertices) is vertex-pulled from vertex_index.

struct Globals {
    projection: mat4x4<f32>,
};

@group(0) @binding(0)
var<uniform> globals: Globals;

struct RectInstance {
    @location(0) pos: vec2<f32>,
    @location(1) size: vec2<f32>,
    @location(2) color: vec4<f32>,
};

struct VertexOutput {
    @builtin(position) position: vec4<f32>,
    @location(0) color: vec4<f32>,
};

@vertex
fn vs_main(@builtin(vertex_index) vi: u32, inst: RectInstance) -> VertexOutput {
    // Unit quad: two triangles covering [0,0]-[1,1]
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
    var out: VertexOutput;
    out.position = globals.projection * vec4<f32>(world_pos, 0.0, 1.0);
    out.color = inst.color;
    return out;
}

@fragment
fn fs_main(in: VertexOutput) -> @location(0) vec4<f32> {
    return in.color;
}

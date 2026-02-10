// Instanced line shader using thin quads.
//
// Each instance provides start, end, line width (+ padding), and color.
// The line is extruded perpendicular to its direction by half the width.

struct Globals {
    projection: mat4x4<f32>,
};

@group(0) @binding(0)
var<uniform> globals: Globals;

struct LineInstance {
    @location(0) start: vec2<f32>,
    @location(1) end: vec2<f32>,
    @location(2) width_pad: vec2<f32>,  // x = line width, y = unused
    @location(3) color: vec4<f32>,
};

struct VertexOutput {
    @builtin(position) position: vec4<f32>,
    @location(0) color: vec4<f32>,
};

@vertex
fn vs_main(@builtin(vertex_index) vi: u32, inst: LineInstance) -> VertexOutput {
    let dir = inst.end - inst.start;
    let len = length(dir);
    // Normal perpendicular to line direction
    var normal: vec2<f32>;
    if len > 0.001 {
        normal = vec2<f32>(-dir.y, dir.x) / len;
    } else {
        normal = vec2<f32>(0.0, 1.0);
    }
    let half_w = inst.width_pad.x * 0.5;

    // 6 vertices for a quad strip along the line
    var corners = array<vec2<f32>, 6>(
        vec2<f32>(0.0, -1.0),
        vec2<f32>(1.0, -1.0),
        vec2<f32>(0.0,  1.0),
        vec2<f32>(0.0,  1.0),
        vec2<f32>(1.0, -1.0),
        vec2<f32>(1.0,  1.0),
    );
    let c = corners[vi];
    let along = mix(inst.start, inst.end, c.x);
    let world_pos = along + normal * half_w * c.y;

    var out: VertexOutput;
    out.position = globals.projection * vec4<f32>(world_pos, 0.0, 1.0);
    out.color = inst.color;
    return out;
}

@fragment
fn fs_main(in: VertexOutput) -> @location(0) vec4<f32> {
    return in.color;
}

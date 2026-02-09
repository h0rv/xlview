use serde::{Deserialize, Deserializer, Serialize, Serializer};
use std::ops::Deref;
use std::sync::Arc;

use super::DxfStyle;

/// Resolved cell style
#[derive(Debug, Serialize, Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Style {
    // Font
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<VertAlign>,

    // Fill
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pattern_type: Option<PatternType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<String>, // Pattern foreground color
    /// Gradient fill (if this cell uses a gradient instead of a solid/pattern fill)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFill>,

    // Borders
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_top: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_right: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_bottom: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_left: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_diagonal: Option<Border>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<bool>, // Line from bottom-left to top-right
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<bool>, // Line from top-left to bottom-right

    // Alignment
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align_h: Option<HAlign>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align_v: Option<VAlign>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shrink_to_fit: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u8>, // 0=context, 1=LTR, 2=RTL

    // Protection
    #[serde(skip_serializing_if = "Option::is_none")]
    pub locked: Option<bool>, // Cell is locked (default true when sheet is protected)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>, // Formula is hidden when sheet is protected
}

#[derive(Debug, Clone)]
pub struct StyleRef(pub Arc<Style>);

impl Deref for StyleRef {
    type Target = Style;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl Serialize for StyleRef {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        self.0.serialize(serializer)
    }
}

impl<'de> Deserialize<'de> for StyleRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let style = Style::deserialize(deserializer)?;
        Ok(Self(Arc::new(style)))
    }
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct Border {
    pub style: BorderStyle,
    pub color: String,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum BorderStyle {
    #[default]
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum HAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum VAlign {
    Top,
    Center, // Note: Excel uses "center" not "middle"
    Bottom,
    Justify,
    Distributed,
}

/// Pane state for frozen/split panes
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum PaneState {
    Frozen,
    FrozenSplit,
    Split,
}

/// Underline style for font formatting
#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
    None,
}

/// Vertical alignment for text (subscript/superscript)
#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum VertAlign {
    Baseline,
    Subscript,
    Superscript,
}

/// Pattern fill types from ECMA-376 Part 1, Section 18.18.55
#[derive(Debug, Serialize, Deserialize, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum PatternType {
    None,
    Solid,
    Gray125,
    Gray0625,
    DarkGray,
    MediumGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeRange {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ColWidth {
    pub col: u32,
    pub width: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct RowHeight {
    pub row: u32,
    pub height: f64,
}

/// Theme colors and fonts extracted from theme1.xml
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct Theme {
    /// 12 theme colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
    pub colors: Vec<String>,
    /// Major font (headings) from fontScheme
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_font: Option<String>,
    /// Minor font (body) from fontScheme
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_font: Option<String>,
}

// ============================================================================
// Internal types for parsing (not serialized to JSON)
// ============================================================================

/// Raw style components from styles.xml
#[allow(clippy::struct_excessive_bools)]
#[derive(Debug, Default, Clone)]
pub struct RawFont {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<ColorSpec>,
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<UnderlineStyle>,
    pub strikethrough: bool,
    pub vert_align: Option<VertAlign>,
    /// Font scheme: "minor" (body) or "major" (headings)
    pub scheme: Option<String>,
}

/// A gradient stop with position and color
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    /// Position of the stop (0.0 to 1.0)
    pub position: f64,
    /// Color at this stop position
    pub color: String,
}

/// Gradient fill definition
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientFill {
    /// Gradient type: "linear" or "path"
    pub gradient_type: String,
    /// Angle in degrees for linear gradients (0 = left-to-right, 90 = top-to-bottom)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub degree: Option<f64>,
    /// Left position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    /// Right position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    /// Top position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    /// Bottom position for path gradients (0.0 to 1.0)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
    /// Color stops defining the gradient
    pub stops: Vec<GradientStop>,
}

/// Raw gradient stop data from styles.xml (before color resolution)
#[derive(Debug, Clone)]
pub struct RawGradientStop {
    /// Position of the stop (0.0 to 1.0)
    pub position: f64,
    /// Color specification at this stop
    pub color: ColorSpec,
}

/// Raw gradient fill data from styles.xml (before color resolution)
#[derive(Debug, Clone, Default)]
pub struct RawGradientFill {
    /// Gradient type: "linear" or "path"
    pub gradient_type: Option<String>,
    /// Angle in degrees for linear gradients
    pub degree: Option<f64>,
    /// Left position for path gradients (0.0 to 1.0)
    pub left: Option<f64>,
    /// Right position for path gradients (0.0 to 1.0)
    pub right: Option<f64>,
    /// Top position for path gradients (0.0 to 1.0)
    pub top: Option<f64>,
    /// Bottom position for path gradients (0.0 to 1.0)
    pub bottom: Option<f64>,
    /// Color stops defining the gradient
    pub stops: Vec<RawGradientStop>,
}

#[derive(Debug, Default, Clone)]
pub struct RawFill {
    pub fg_color: Option<ColorSpec>,
    pub bg_color: Option<ColorSpec>,
    pub pattern_type: Option<String>,
    /// Gradient fill data (if this is a gradient fill instead of a pattern fill)
    pub gradient: Option<RawGradientFill>,
}

#[derive(Debug, Default, Clone)]
pub struct RawBorder {
    pub left: Option<RawBorderSide>,
    pub right: Option<RawBorderSide>,
    pub top: Option<RawBorderSide>,
    pub bottom: Option<RawBorderSide>,
    pub diagonal: Option<RawBorderSide>,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}

#[derive(Debug, Clone)]
pub struct RawBorderSide {
    pub style: String,
    pub color: Option<ColorSpec>,
}

#[derive(Debug, Clone)]
pub struct ColorSpec {
    pub rgb: Option<String>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
    pub indexed: Option<u32>,
    pub auto: bool,
}

#[derive(Debug, Default, Clone)]
pub struct RawAlignment {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
    pub indent: Option<u32>,
    pub text_rotation: Option<i32>,
    pub reading_order: Option<u8>, // 0=context, 1=LTR, 2=RTL
}

/// Raw protection settings from styles.xml
#[derive(Debug, Default, Clone)]
pub struct RawProtection {
    pub locked: bool, // default true in Excel
    pub hidden: bool, // default false
}

/// Cell format (xf) from cellXfs or cellStyleXfs
///
/// Per ECMA-376 Section 18.8.45, the apply* attributes default to TRUE when absent.
#[allow(clippy::struct_excessive_bools)]
#[derive(Debug, Clone)]
pub struct CellXf {
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
    pub num_fmt_id: Option<u32>,
    pub alignment: Option<RawAlignment>,
    pub apply_font: bool,
    pub apply_fill: bool,
    pub apply_border: bool,
    pub apply_alignment: bool,
    pub apply_number_format: bool,
    pub apply_protection: bool,
    pub protection: Option<RawProtection>,
    /// Reference to cellStyleXfs entry (for cellXfs only)
    pub xf_id: Option<u32>,
}

impl Default for CellXf {
    fn default() -> Self {
        Self {
            font_id: None,
            fill_id: None,
            border_id: None,
            num_fmt_id: None,
            alignment: None,
            // Per ECMA-376, apply* attributes default to TRUE when absent
            apply_font: true,
            apply_fill: true,
            apply_border: true,
            apply_alignment: true,
            apply_number_format: true,
            apply_protection: true,
            protection: None,
            xf_id: None,
        }
    }
}

/// Named style metadata from cellStyles
#[derive(Debug, Clone)]
pub struct NamedStyle {
    pub name: String,
    pub xf_id: u32,
    pub builtin_id: Option<u32>,
}

/// Complete parsed style data from styles.xml
#[derive(Debug, Default)]
pub struct StyleSheet {
    pub fonts: Vec<RawFont>,
    pub fills: Vec<RawFill>,
    pub borders: Vec<RawBorder>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<(u32, String)>, // (numFmtId, formatCode)
    /// Named styles from cellStyleXfs (base styles that cellXfs can inherit from)
    pub cell_style_xfs: Vec<CellXf>,
    /// Named style metadata from cellStyles
    pub named_styles: Vec<NamedStyle>,
    /// Custom indexed colors from `<colors><indexedColors>` (if present)
    /// Falls back to INDEXED_COLORS if None
    pub indexed_colors: Option<Vec<String>>,
    /// Default font from the "Normal" style (`cellStyleXfs[0]`)
    /// This is the workbook's default font that applies to all cells
    pub default_font: Option<RawFont>,
    /// Differential formatting styles (dxfs) for conditional formatting
    pub dxf_styles: Vec<DxfStyle>,
}

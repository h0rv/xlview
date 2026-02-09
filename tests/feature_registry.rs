//! Feature Registry - Master list of all Excel features with test coverage tracking.
//!
//! This module provides a comprehensive catalog of XLSX features that xlview supports,
//! along with their test coverage status. It serves as both documentation and a tool
//! for tracking test completeness.
//!
//! # Usage
//!
//! ```rust
//! use tests::feature_registry::{FEATURES, coverage_report};
//!
//! // Get overall coverage percentage
//! let (covered, total, pct) = coverage_stats();
//! println!("Coverage: {}/{} ({:.1}%)", covered, total, pct);
//!
//! // Generate markdown report
//! let report = coverage_report();
//! ```
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

/// Feature test status
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FeatureStatus {
    /// Feature has comprehensive tests
    Tested,
    /// Feature has partial test coverage
    Partial,
    /// Feature parsing exists but no dedicated tests
    Untested,
    /// Feature is not yet implemented
    NotImplemented,
}

impl FeatureStatus {
    pub fn as_str(&self) -> &'static str {
        match self {
            FeatureStatus::Tested => "tested",
            FeatureStatus::Partial => "partial",
            FeatureStatus::Untested => "untested",
            FeatureStatus::NotImplemented => "not_implemented",
        }
    }

    pub fn emoji(&self) -> &'static str {
        match self {
            FeatureStatus::Tested => "[x]",
            FeatureStatus::Partial => "[~]",
            FeatureStatus::Untested => "[ ]",
            FeatureStatus::NotImplemented => "[-]",
        }
    }
}

/// Feature category
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Category {
    CellTypes,
    FontStyling,
    FillPatterns,
    Borders,
    Alignment,
    NumberFormats,
    ConditionalFormatting,
    Charts,
    DataValidation,
    Drawing,
    Comments,
    Hyperlinks,
    SheetFeatures,
    WorkbookFeatures,
}

impl Category {
    pub fn as_str(&self) -> &'static str {
        match self {
            Category::CellTypes => "Cell Types",
            Category::FontStyling => "Font Styling",
            Category::FillPatterns => "Fill Patterns",
            Category::Borders => "Borders",
            Category::Alignment => "Alignment",
            Category::NumberFormats => "Number Formats",
            Category::ConditionalFormatting => "Conditional Formatting",
            Category::Charts => "Charts",
            Category::DataValidation => "Data Validation",
            Category::Drawing => "Drawing (Images/Shapes)",
            Category::Comments => "Comments",
            Category::Hyperlinks => "Hyperlinks",
            Category::SheetFeatures => "Sheet Features",
            Category::WorkbookFeatures => "Workbook Features",
        }
    }

    pub fn all() -> &'static [Category] {
        &[
            Category::CellTypes,
            Category::FontStyling,
            Category::FillPatterns,
            Category::Borders,
            Category::Alignment,
            Category::NumberFormats,
            Category::ConditionalFormatting,
            Category::Charts,
            Category::DataValidation,
            Category::Drawing,
            Category::Comments,
            Category::Hyperlinks,
            Category::SheetFeatures,
            Category::WorkbookFeatures,
        ]
    }
}

/// A feature in the registry
#[derive(Debug, Clone)]
pub struct Feature {
    /// Unique identifier (e.g., "cell_type_string")
    pub id: &'static str,
    /// Human-readable name
    pub name: &'static str,
    /// Feature category
    pub category: Category,
    /// Test file(s) that cover this feature
    pub test_files: &'static [&'static str],
    /// Current test status
    pub status: FeatureStatus,
    /// ECMA-376 section reference (if applicable)
    pub spec_ref: Option<&'static str>,
}

/// Master list of all features
pub static FEATURES: &[Feature] = &[
    // =========================================================================
    // Cell Types (6 features)
    // =========================================================================
    Feature {
        id: "cell_type_string",
        name: "String cells (shared strings)",
        category: Category::CellTypes,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_inline_string",
        name: "Inline string cells",
        category: Category::CellTypes,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_number",
        name: "Numeric cells",
        category: Category::CellTypes,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_boolean",
        name: "Boolean cells",
        category: Category::CellTypes,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_date",
        name: "Date cells (serial numbers)",
        category: Category::CellTypes,
        test_files: &["date_format_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_error",
        name: "Error cells (#DIV/0!, #N/A, etc.)",
        category: Category::CellTypes,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.4"),
    },
    Feature {
        id: "cell_type_rich_text",
        name: "Rich text cells",
        category: Category::CellTypes,
        test_files: &["rich_text_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.4.4"),
    },
    // =========================================================================
    // Font Styling (12 features)
    // =========================================================================
    Feature {
        id: "font_bold",
        name: "Bold text",
        category: Category::FontStyling,
        test_files: &["font_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.2"),
    },
    Feature {
        id: "font_italic",
        name: "Italic text",
        category: Category::FontStyling,
        test_files: &["font_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.26"),
    },
    Feature {
        id: "font_underline",
        name: "Underline text",
        category: Category::FontStyling,
        test_files: &["font_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.42"),
    },
    Feature {
        id: "font_underline_double",
        name: "Double underline",
        category: Category::FontStyling,
        test_files: &["font_tests.rs"],
        status: FeatureStatus::Partial,
        spec_ref: Some("18.8.42"),
    },
    Feature {
        id: "font_strikethrough",
        name: "Strikethrough text",
        category: Category::FontStyling,
        test_files: &["font_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.37"),
    },
    Feature {
        id: "font_name",
        name: "Font family name",
        category: Category::FontStyling,
        test_files: &["font_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.29"),
    },
    Feature {
        id: "font_size",
        name: "Font size",
        category: Category::FontStyling,
        test_files: &["font_tests.rs", "font_rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.38"),
    },
    Feature {
        id: "font_color_rgb",
        name: "Font color (RGB)",
        category: Category::FontStyling,
        test_files: &["font_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.3"),
    },
    Feature {
        id: "font_color_theme",
        name: "Font color (theme)",
        category: Category::FontStyling,
        test_files: &["theme_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.3"),
    },
    Feature {
        id: "font_color_indexed",
        name: "Font color (indexed)",
        category: Category::FontStyling,
        test_files: &["font_tests.rs"],
        status: FeatureStatus::Partial,
        spec_ref: Some("18.8.3"),
    },
    Feature {
        id: "font_subscript",
        name: "Subscript",
        category: Category::FontStyling,
        test_files: &["rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.44"),
    },
    Feature {
        id: "font_superscript",
        name: "Superscript",
        category: Category::FontStyling,
        test_files: &["rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.44"),
    },
    // =========================================================================
    // Fill Patterns (19 features - all ECMA-376 patterns)
    // =========================================================================
    Feature {
        id: "fill_none",
        name: "No fill (none)",
        category: Category::FillPatterns,
        test_files: &["fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_solid",
        name: "Solid fill",
        category: Category::FillPatterns,
        test_files: &["fill_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_gray125",
        name: "Gray 12.5% pattern",
        category: Category::FillPatterns,
        test_files: &["fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_gray0625",
        name: "Gray 6.25% pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_gray",
        name: "Dark gray pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_medium_gray",
        name: "Medium gray pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_gray",
        name: "Light gray pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_horizontal",
        name: "Dark horizontal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_vertical",
        name: "Dark vertical pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_down",
        name: "Dark down diagonal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_up",
        name: "Dark up diagonal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_grid",
        name: "Dark grid pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_dark_trellis",
        name: "Dark trellis pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_horizontal",
        name: "Light horizontal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_vertical",
        name: "Light vertical pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_down",
        name: "Light down diagonal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_up",
        name: "Light up diagonal pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_grid",
        name: "Light grid pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_light_trellis",
        name: "Light trellis pattern",
        category: Category::FillPatterns,
        test_files: &["pattern_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.55"),
    },
    Feature {
        id: "fill_gradient",
        name: "Gradient fill",
        category: Category::FillPatterns,
        test_files: &["gradient_fill_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.24"),
    },
    // =========================================================================
    // Borders (13 styles)
    // =========================================================================
    Feature {
        id: "border_none",
        name: "No border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_thin",
        name: "Thin border",
        category: Category::Borders,
        test_files: &["border_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_medium",
        name: "Medium border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_thick",
        name: "Thick border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_dashed",
        name: "Dashed border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_dotted",
        name: "Dotted border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_double",
        name: "Double border",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_hair",
        name: "Hair border",
        category: Category::Borders,
        test_files: &["border_tests.rs", "border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_medium_dashed",
        name: "Medium dashed border",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_dash_dot",
        name: "Dash-dot border",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_medium_dash_dot",
        name: "Medium dash-dot border",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_dash_dot_dot",
        name: "Dash-dot-dot border",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_slant_dash_dot",
        name: "Slant dash-dot border",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.18.3"),
    },
    Feature {
        id: "border_diagonal",
        name: "Diagonal borders",
        category: Category::Borders,
        test_files: &["border_style_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.4"),
    },
    Feature {
        id: "border_color_rgb",
        name: "Border color (RGB)",
        category: Category::Borders,
        test_files: &["border_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.3"),
    },
    Feature {
        id: "border_color_theme",
        name: "Border color (theme)",
        category: Category::Borders,
        test_files: &["theme_tests.rs"],
        status: FeatureStatus::Partial,
        spec_ref: Some("18.8.3"),
    },
    // =========================================================================
    // Alignment (10 features)
    // =========================================================================
    Feature {
        id: "align_h_left",
        name: "Horizontal align: left",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_h_center",
        name: "Horizontal align: center",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_h_right",
        name: "Horizontal align: right",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_h_justify",
        name: "Horizontal align: justify",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_h_fill",
        name: "Horizontal align: fill",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_v_top",
        name: "Vertical align: top",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_v_center",
        name: "Vertical align: center",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_v_bottom",
        name: "Vertical align: bottom",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_wrap_text",
        name: "Text wrapping",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_indent",
        name: "Text indent",
        category: Category::Alignment,
        test_files: &["alignment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_rotation",
        name: "Text rotation",
        category: Category::Alignment,
        test_files: &["text_rotation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    Feature {
        id: "align_shrink_to_fit",
        name: "Shrink to fit",
        category: Category::Alignment,
        test_files: &["shrink_to_fit_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.1"),
    },
    // =========================================================================
    // Number Formats (15+ features)
    // =========================================================================
    Feature {
        id: "numfmt_general",
        name: "General format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_number",
        name: "Number format (0.00)",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_currency",
        name: "Currency format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_accounting",
        name: "Accounting format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_date",
        name: "Date format",
        category: Category::NumberFormats,
        test_files: &["date_format_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_time",
        name: "Time format",
        category: Category::NumberFormats,
        test_files: &["date_format_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_percentage",
        name: "Percentage format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_fraction",
        name: "Fraction format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Partial,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_scientific",
        name: "Scientific format",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_text",
        name: "Text format (@)",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    Feature {
        id: "numfmt_custom",
        name: "Custom number formats",
        category: Category::NumberFormats,
        test_files: &["numfmt_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.30"),
    },
    // =========================================================================
    // Conditional Formatting (30+ features)
    // =========================================================================
    Feature {
        id: "cf_color_scale_2",
        name: "2-color scale",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.16"),
    },
    Feature {
        id: "cf_color_scale_3",
        name: "3-color scale",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.16"),
    },
    Feature {
        id: "cf_data_bar",
        name: "Data bars",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.28"),
    },
    Feature {
        id: "cf_data_bar_negative",
        name: "Data bars (negative values)",
        category: Category::ConditionalFormatting,
        test_files: &["cf_data_bar_negative_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.28"),
    },
    Feature {
        id: "cf_icon_set_3arrows",
        name: "Icon set: 3 Arrows",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_3traffic",
        name: "Icon set: 3 Traffic Lights",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_3symbols",
        name: "Icon set: 3 Symbols",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_3flags",
        name: "Icon set: 3 Flags",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_4arrows",
        name: "Icon set: 4 Arrows",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_4traffic",
        name: "Icon set: 4 Traffic Lights",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_5arrows",
        name: "Icon set: 5 Arrows",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_icon_set_5rating",
        name: "Icon set: 5 Ratings",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.49"),
    },
    Feature {
        id: "cf_cell_is_equal",
        name: "Cell Is: equal to",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.10"),
    },
    Feature {
        id: "cf_cell_is_not_equal",
        name: "Cell Is: not equal to",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.10"),
    },
    Feature {
        id: "cf_cell_is_greater",
        name: "Cell Is: greater than",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.10"),
    },
    Feature {
        id: "cf_cell_is_less",
        name: "Cell Is: less than",
        category: Category::ConditionalFormatting,
        test_files: &["conditional_formatting_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.10"),
    },
    Feature {
        id: "cf_cell_is_between",
        name: "Cell Is: between",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.10"),
    },
    Feature {
        id: "cf_top_10",
        name: "Top/Bottom 10",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.92"),
    },
    Feature {
        id: "cf_above_average",
        name: "Above/Below Average",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.3"),
    },
    Feature {
        id: "cf_duplicate_values",
        name: "Duplicate Values",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_unique_values",
        name: "Unique Values",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_text_contains",
        name: "Text Contains",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_text_begins",
        name: "Text Begins With",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_text_ends",
        name: "Text Ends With",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_date_occurring",
        name: "Date Occurring",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.91"),
    },
    Feature {
        id: "cf_blanks",
        name: "Blanks/No Blanks",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_errors",
        name: "Errors/No Errors",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.12"),
    },
    Feature {
        id: "cf_formula",
        name: "Formula-based rule",
        category: Category::ConditionalFormatting,
        test_files: &["cf_rule_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.43"),
    },
    // =========================================================================
    // Charts (12+ types)
    // =========================================================================
    Feature {
        id: "chart_bar_clustered",
        name: "Bar chart (clustered)",
        category: Category::Charts,
        test_files: &["charts_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.16"),
    },
    Feature {
        id: "chart_bar_stacked",
        name: "Bar chart (stacked)",
        category: Category::Charts,
        test_files: &["charts_tests.rs", "chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.16"),
    },
    Feature {
        id: "chart_column",
        name: "Column chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.16"),
    },
    Feature {
        id: "chart_line",
        name: "Line chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.97"),
    },
    Feature {
        id: "chart_pie",
        name: "Pie chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.141"),
    },
    Feature {
        id: "chart_doughnut",
        name: "Doughnut chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs", "chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.50"),
    },
    Feature {
        id: "chart_area",
        name: "Area chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.9"),
    },
    Feature {
        id: "chart_scatter",
        name: "Scatter chart",
        category: Category::Charts,
        test_files: &["charts_tests.rs", "chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.162"),
    },
    Feature {
        id: "chart_bubble",
        name: "Bubble chart",
        category: Category::Charts,
        test_files: &["chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.20"),
    },
    Feature {
        id: "chart_radar",
        name: "Radar chart",
        category: Category::Charts,
        test_files: &["chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.153"),
    },
    Feature {
        id: "chart_stock",
        name: "Stock chart",
        category: Category::Charts,
        test_files: &["chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.177"),
    },
    Feature {
        id: "chart_surface",
        name: "Surface chart",
        category: Category::Charts,
        test_files: &["chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("21.2.2.178"),
    },
    Feature {
        id: "chart_combo",
        name: "Combination chart",
        category: Category::Charts,
        test_files: &["chart_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: None,
    },
    // =========================================================================
    // Data Validation (8 features)
    // =========================================================================
    Feature {
        id: "dv_list",
        name: "List validation (dropdown)",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_whole_number",
        name: "Whole number validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_decimal",
        name: "Decimal validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_date",
        name: "Date validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_time",
        name: "Time validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_text_length",
        name: "Text length validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_custom",
        name: "Custom formula validation",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    Feature {
        id: "dv_error_alert",
        name: "Error alert messages",
        category: Category::DataValidation,
        test_files: &["data_validation_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.32"),
    },
    // =========================================================================
    // Drawing (Images & Shapes)
    // =========================================================================
    Feature {
        id: "drawing_image_png",
        name: "Embedded PNG images",
        category: Category::Drawing,
        test_files: &["drawing_tests.rs", "images_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.5"),
    },
    Feature {
        id: "drawing_image_jpeg",
        name: "Embedded JPEG images",
        category: Category::Drawing,
        test_files: &["drawing_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.5"),
    },
    Feature {
        id: "drawing_image_positioning",
        name: "Image positioning (anchors)",
        category: Category::Drawing,
        test_files: &["drawing_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.5.2"),
    },
    Feature {
        id: "drawing_shapes",
        name: "Auto shapes",
        category: Category::Drawing,
        test_files: &["drawing_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.1"),
    },
    Feature {
        id: "drawing_text_boxes",
        name: "Text boxes",
        category: Category::Drawing,
        test_files: &["text_box_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.1"),
    },
    // =========================================================================
    // Comments
    // =========================================================================
    Feature {
        id: "comment_basic",
        name: "Cell comments",
        category: Category::Comments,
        test_files: &["comment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.7"),
    },
    Feature {
        id: "comment_author",
        name: "Comment author",
        category: Category::Comments,
        test_files: &["comment_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.7.2"),
    },
    Feature {
        id: "comment_rich_text",
        name: "Comment rich text",
        category: Category::Comments,
        test_files: &["comment_rich_text_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.7.3"),
    },
    Feature {
        id: "comment_indicator",
        name: "Comment indicator rendering",
        category: Category::Comments,
        test_files: &["comment_indicator_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: None,
    },
    // =========================================================================
    // Hyperlinks
    // =========================================================================
    Feature {
        id: "hyperlink_external",
        name: "External hyperlinks (URLs)",
        category: Category::Hyperlinks,
        test_files: &["hyperlink_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.47"),
    },
    Feature {
        id: "hyperlink_internal",
        name: "Internal hyperlinks (sheet refs)",
        category: Category::Hyperlinks,
        test_files: &["hyperlink_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.47"),
    },
    Feature {
        id: "hyperlink_email",
        name: "Email hyperlinks (mailto:)",
        category: Category::Hyperlinks,
        test_files: &["hyperlink_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.47"),
    },
    Feature {
        id: "hyperlink_rendering",
        name: "Hyperlink rendering (blue, underline)",
        category: Category::Hyperlinks,
        test_files: &["hyperlink_rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: None,
    },
    // =========================================================================
    // Sheet Features
    // =========================================================================
    Feature {
        id: "sheet_frozen_rows",
        name: "Frozen rows",
        category: Category::SheetFeatures,
        test_files: &["frozen_panes_tests.rs", "frozen_panes_rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.66"),
    },
    Feature {
        id: "sheet_frozen_cols",
        name: "Frozen columns",
        category: Category::SheetFeatures,
        test_files: &["frozen_panes_tests.rs", "frozen_panes_rendering_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.66"),
    },
    Feature {
        id: "sheet_merged_cells",
        name: "Merged cells",
        category: Category::SheetFeatures,
        test_files: &["merge_tests.rs", "integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.55"),
    },
    Feature {
        id: "sheet_col_width",
        name: "Column widths",
        category: Category::SheetFeatures,
        test_files: &["dimension_tests.rs", "layout_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.13"),
    },
    Feature {
        id: "sheet_row_height",
        name: "Row heights",
        category: Category::SheetFeatures,
        test_files: &["dimension_tests.rs", "layout_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.73"),
    },
    Feature {
        id: "sheet_hidden_rows",
        name: "Hidden rows",
        category: Category::SheetFeatures,
        test_files: &["dimension_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.73"),
    },
    Feature {
        id: "sheet_hidden_cols",
        name: "Hidden columns",
        category: Category::SheetFeatures,
        test_files: &["dimension_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.13"),
    },
    Feature {
        id: "sheet_auto_filter",
        name: "Auto filter",
        category: Category::SheetFeatures,
        test_files: &["auto_filter_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.2"),
    },
    Feature {
        id: "sheet_outline",
        name: "Row/column grouping (outline)",
        category: Category::SheetFeatures,
        test_files: &["outline_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.63"),
    },
    Feature {
        id: "sheet_protection",
        name: "Sheet protection",
        category: Category::SheetFeatures,
        test_files: &["protection_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.85"),
    },
    Feature {
        id: "sheet_page_setup",
        name: "Page setup (print settings)",
        category: Category::SheetFeatures,
        test_files: &["page_setup_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.64"),
    },
    Feature {
        id: "sheet_sparklines",
        name: "Sparklines",
        category: Category::SheetFeatures,
        test_files: &["sparkline_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.96"),
    },
    // =========================================================================
    // Workbook Features
    // =========================================================================
    Feature {
        id: "workbook_sheet_names",
        name: "Sheet names",
        category: Category::WorkbookFeatures,
        test_files: &["integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.2.19"),
    },
    Feature {
        id: "workbook_tab_colors",
        name: "Tab colors",
        category: Category::WorkbookFeatures,
        test_files: &["integration_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.3.1.93"),
    },
    Feature {
        id: "workbook_hidden_sheets",
        name: "Hidden sheets",
        category: Category::WorkbookFeatures,
        test_files: &["sheet_visibility_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.2.19"),
    },
    Feature {
        id: "workbook_very_hidden_sheets",
        name: "Very hidden sheets",
        category: Category::WorkbookFeatures,
        test_files: &["sheet_visibility_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.2.19"),
    },
    Feature {
        id: "workbook_defined_names",
        name: "Defined names (named ranges)",
        category: Category::WorkbookFeatures,
        test_files: &["defined_names_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.2.6"),
    },
    Feature {
        id: "workbook_theme_colors",
        name: "Theme colors",
        category: Category::WorkbookFeatures,
        test_files: &["theme_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("20.1.6"),
    },
    Feature {
        id: "workbook_named_styles",
        name: "Named cell styles",
        category: Category::WorkbookFeatures,
        test_files: &["named_styles_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.8.10"),
    },
    Feature {
        id: "workbook_shared_strings",
        name: "Shared string table",
        category: Category::WorkbookFeatures,
        test_files: &["cell_type_tests.rs"],
        status: FeatureStatus::Tested,
        spec_ref: Some("18.4"),
    },
];

// ============================================================================
// Coverage Statistics Functions
// ============================================================================

/// Calculate coverage statistics
pub fn coverage_stats() -> (usize, usize, f64) {
    let total = FEATURES.len();
    let covered = FEATURES
        .iter()
        .filter(|f| matches!(f.status, FeatureStatus::Tested | FeatureStatus::Partial))
        .count();
    let percentage = if total > 0 {
        (covered as f64 / total as f64) * 100.0
    } else {
        0.0
    };
    (covered, total, percentage)
}

/// Calculate coverage by category
pub fn coverage_by_category(category: Category) -> (usize, usize, f64) {
    let features: Vec<_> = FEATURES.iter().filter(|f| f.category == category).collect();
    let total = features.len();
    let covered = features
        .iter()
        .filter(|f| matches!(f.status, FeatureStatus::Tested | FeatureStatus::Partial))
        .count();
    let percentage = if total > 0 {
        (covered as f64 / total as f64) * 100.0
    } else {
        0.0
    };
    (covered, total, percentage)
}

/// Generate a markdown coverage report
pub fn coverage_report() -> String {
    let mut report = String::new();

    report.push_str("# xlview Feature Coverage Report\n\n");

    // Overall summary
    let (covered, total, pct) = coverage_stats();
    report.push_str(&format!(
        "## Summary\n\n**Overall Coverage: {}/{} ({:.1}%)**\n\n",
        covered, total, pct
    ));

    // Status legend
    report.push_str("### Status Legend\n\n");
    report.push_str("- `[x]` Tested - Feature has comprehensive tests\n");
    report.push_str("- `[~]` Partial - Feature has partial test coverage\n");
    report.push_str("- `[ ]` Untested - Feature exists but lacks dedicated tests\n");
    report.push_str("- `[-]` Not Implemented - Feature is not yet supported\n\n");

    // Coverage by category
    report.push_str("## Coverage by Category\n\n");
    report.push_str("| Category | Covered | Total | Percentage |\n");
    report.push_str("|----------|---------|-------|------------|\n");

    for category in Category::all() {
        let (c, t, p) = coverage_by_category(*category);
        report.push_str(&format!(
            "| {} | {} | {} | {:.1}% |\n",
            category.as_str(),
            c,
            t,
            p
        ));
    }
    report.push('\n');

    // Detailed feature list by category
    report.push_str("## Feature Details\n\n");

    for category in Category::all() {
        let features: Vec<_> = FEATURES
            .iter()
            .filter(|f| f.category == *category)
            .collect();
        if features.is_empty() {
            continue;
        }

        report.push_str(&format!("### {}\n\n", category.as_str()));
        report.push_str("| Status | Feature | Test Files | Spec Ref |\n");
        report.push_str("|--------|---------|------------|----------|\n");

        for feature in features {
            let test_files = if feature.test_files.is_empty() {
                "-".to_string()
            } else {
                feature.test_files.join(", ")
            };
            let spec_ref = feature.spec_ref.unwrap_or("-");
            report.push_str(&format!(
                "| {} | {} | {} | {} |\n",
                feature.status.emoji(),
                feature.name,
                test_files,
                spec_ref
            ));
        }
        report.push('\n');
    }

    // Untested features summary
    let untested: Vec<_> = FEATURES
        .iter()
        .filter(|f| matches!(f.status, FeatureStatus::Untested))
        .collect();

    if !untested.is_empty() {
        report.push_str("## Untested Features (Need Coverage)\n\n");
        for feature in untested {
            report.push_str(&format!(
                "- {} ({})\n",
                feature.name,
                feature.category.as_str()
            ));
        }
        report.push('\n');
    }

    // Not implemented features
    let not_impl: Vec<_> = FEATURES
        .iter()
        .filter(|f| matches!(f.status, FeatureStatus::NotImplemented))
        .collect();

    if !not_impl.is_empty() {
        report.push_str("## Not Implemented Features\n\n");
        for feature in not_impl {
            report.push_str(&format!(
                "- {} ({})\n",
                feature.name,
                feature.category.as_str()
            ));
        }
    }

    report
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_feature_count() {
        // Sanity check - we should have a substantial number of features
        assert!(
            FEATURES.len() > 100,
            "Expected 100+ features, got {}",
            FEATURES.len()
        );
    }

    #[test]
    fn test_coverage_stats() {
        let (covered, total, pct) = coverage_stats();
        assert!(total > 0);
        assert!(covered <= total);
        assert!((0.0..=100.0).contains(&pct));
    }

    #[test]
    fn test_all_categories_represented() {
        for category in Category::all() {
            let count = FEATURES.iter().filter(|f| f.category == *category).count();
            assert!(count > 0, "Category {} has no features", category.as_str());
        }
    }

    #[test]
    fn test_unique_feature_ids() {
        let mut ids: Vec<_> = FEATURES.iter().map(|f| f.id).collect();
        ids.sort();
        let len_before = ids.len();
        ids.dedup();
        assert_eq!(len_before, ids.len(), "Feature IDs must be unique");
    }

    #[test]
    fn test_coverage_report_generation() {
        let report = coverage_report();
        assert!(report.contains("# xlview Feature Coverage Report"));
        assert!(report.contains("## Summary"));
        assert!(report.contains("## Coverage by Category"));
        assert!(report.contains("## Feature Details"));
    }

    #[test]
    #[ignore] // Run with: cargo test print_coverage -- --ignored --nocapture
    fn print_coverage() {
        println!("{}", coverage_report());
    }

    #[test]
    #[ignore] // Run with: cargo test generate_coverage_md -- --ignored
    fn generate_coverage_md() {
        use std::fs;
        let report = coverage_report();
        fs::write("COVERAGE.md", &report).expect("Failed to write COVERAGE.md");
        println!("Generated COVERAGE.md with {} bytes", report.len());
    }
}

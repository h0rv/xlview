//! Sparkline parsing module
//! This module handles parsing of sparklines from XLSX extension elements.

use crate::color::resolve_color;
use crate::types::{
    Sparkline, SparklineAxisType, SparklineColors, SparklineEmptyCells, SparklineGroup,
    SparklineType,
};
use crate::xml_helpers::parse_color_attrs;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

/// Parse sparklines from extension elements
///
/// Sparklines in XLSX files are stored in extension elements within the worksheet XML:
/// ```xml
/// <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
///   <x14:sparklineGroups xmlns:xm="...">
///     <x14:sparklineGroup type="line" displayEmptyCellsAs="gap">
///       <x14:colorSeries theme="4"/>
///       <x14:colorNegative theme="5"/>
///       <x14:sparklines>
///         <x14:sparkline>
///           <xm:f>Sheet1!A1:A10</xm:f>
///           <xm:sqref>B1</xm:sqref>
///         </x14:sparkline>
///       </x14:sparklines>
///     </x14:sparklineGroup>
///   </x14:sparklineGroups>
/// </ext>
/// ```
#[allow(clippy::too_many_lines)]
pub fn parse_ext_sparklines<R: BufRead>(
    _start_element: &BytesStart,
    xml: &mut Reader<R>,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Vec<SparklineGroup> {
    let mut groups = Vec::new();
    let mut buf = Vec::new();

    // Track parsing state
    let mut in_sparkline_groups = false;
    let mut current_group: Option<SparklineGroupBuilder> = None;
    let mut in_sparklines = false;
    let mut current_sparkline: Option<SparklineBuilder> = None;
    let mut in_f = false; // Inside xm:f (formula/data range)
    let mut in_sqref = false; // Inside xm:sqref (location)

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "sparklineGroups" => {
                        in_sparkline_groups = true;
                    }
                    "sparklineGroup" if in_sparkline_groups => {
                        current_group = Some(parse_sparkline_group_attrs(e));
                    }
                    // Color elements
                    "colorSeries" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.series =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorNegative" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.negative =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorAxis" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.axis =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorMarkers" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.markers =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorFirst" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.first =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorLast" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.last =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorHigh" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.high =
                                parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "colorLow" if current_group.is_some() => {
                        if let Some(ref mut group) = current_group {
                            group.colors.low = parse_color_element(e, theme_colors, indexed_colors);
                        }
                    }
                    "sparklines" if current_group.is_some() => {
                        in_sparklines = true;
                    }
                    "sparkline" if in_sparklines => {
                        current_sparkline = Some(SparklineBuilder::default());
                    }
                    "f" if current_sparkline.is_some() => {
                        in_f = true;
                    }
                    "sqref" if current_sparkline.is_some() => {
                        in_sqref = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Ok(text) = e.unescape() {
                    let text = text.trim().to_string();
                    if in_f {
                        if let Some(ref mut sparkline) = current_sparkline {
                            sparkline.data_range = Some(text);
                        }
                    } else if in_sqref {
                        if let Some(ref mut sparkline) = current_sparkline {
                            sparkline.location = Some(text);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "sparklineGroups" => {
                        in_sparkline_groups = false;
                    }
                    "sparklineGroup" => {
                        if let Some(builder) = current_group.take() {
                            groups.push(builder.build());
                        }
                    }
                    "sparklines" => {
                        in_sparklines = false;
                    }
                    "sparkline" => {
                        if let Some(sparkline_builder) = current_sparkline.take() {
                            if let Some(sparkline) = sparkline_builder.build() {
                                if let Some(ref mut group) = current_group {
                                    group.sparklines.push(sparkline);
                                }
                            }
                        }
                    }
                    "f" => {
                        in_f = false;
                    }
                    "sqref" => {
                        in_sqref = false;
                    }
                    "ext" => {
                        // End of extension element
                        break;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    groups
}

/// Builder for `SparklineGroup` during parsing
#[derive(Default)]
struct SparklineGroupBuilder {
    sparkline_type: SparklineType,
    sparklines: Vec<Sparkline>,
    colors: SparklineColors,
    display_empty_cells_as: SparklineEmptyCells,
    show_markers: Option<bool>,
    show_high: Option<bool>,
    show_low: Option<bool>,
    show_first: Option<bool>,
    show_last: Option<bool>,
    show_negative: Option<bool>,
    show_axis: Option<bool>,
    display_hidden: Option<bool>,
    right_to_left: Option<bool>,
    line_weight: Option<f64>,
    min_axis_type: Option<SparklineAxisType>,
    max_axis_type: Option<SparklineAxisType>,
    manual_min: Option<f64>,
    manual_max: Option<f64>,
    date_axis: Option<bool>,
}

impl SparklineGroupBuilder {
    fn build(self) -> SparklineGroup {
        SparklineGroup {
            sparkline_type: match self.sparkline_type {
                SparklineType::Line => "line".to_string(),
                SparklineType::Column => "column".to_string(),
                SparklineType::Stacked => "stacked".to_string(),
            },
            sparklines: self.sparklines,
            colors: self.colors,
            display_empty_cells_as: Some(match self.display_empty_cells_as {
                SparklineEmptyCells::Gap => "gap".to_string(),
                SparklineEmptyCells::Zero => "zero".to_string(),
                SparklineEmptyCells::Connect => "span".to_string(),
            }),
            markers: self.show_markers.unwrap_or(false),
            high_point: self.show_high.unwrap_or(false),
            low_point: self.show_low.unwrap_or(false),
            first_point: self.show_first.unwrap_or(false),
            last_point: self.show_last.unwrap_or(false),
            negative_points: self.show_negative.unwrap_or(false),
            display_x_axis: self.show_axis.unwrap_or(false),
            display_hidden: self.display_hidden,
            right_to_left: self.right_to_left.unwrap_or(false),
            line_weight: self.line_weight,
            min_axis_type: self.min_axis_type.map(|t| match t {
                SparklineAxisType::Individual => "individual".to_string(),
                SparklineAxisType::Group => "group".to_string(),
                SparklineAxisType::Custom => "custom".to_string(),
            }),
            max_axis_type: self.max_axis_type.map(|t| match t {
                SparklineAxisType::Individual => "individual".to_string(),
                SparklineAxisType::Group => "group".to_string(),
                SparklineAxisType::Custom => "custom".to_string(),
            }),
            manual_min: self.manual_min,
            manual_max: self.manual_max,
            date_axis: self.date_axis,
        }
    }
}

/// Builder for `Sparkline` during parsing
#[derive(Default)]
struct SparklineBuilder {
    data_range: Option<String>,
    location: Option<String>,
}

impl SparklineBuilder {
    fn build(self) -> Option<Sparkline> {
        match (self.data_range, self.location) {
            (Some(data_range), Some(location)) => Some(Sparkline {
                data_range,
                location,
            }),
            _ => None,
        }
    }
}

/// Parse attributes from a sparklineGroup element
fn parse_sparkline_group_attrs(e: &BytesStart) -> SparklineGroupBuilder {
    let mut builder = SparklineGroupBuilder::default();

    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let value = std::str::from_utf8(&attr.value).unwrap_or("");

        match key {
            b"type" => {
                builder.sparkline_type = match value {
                    "column" => SparklineType::Column,
                    "stacked" => SparklineType::Stacked,
                    _ => SparklineType::Line, // "line" or default
                };
            }
            b"displayEmptyCellsAs" => {
                builder.display_empty_cells_as = match value {
                    "zero" => SparklineEmptyCells::Zero,
                    "span" | "connect" => SparklineEmptyCells::Connect,
                    _ => SparklineEmptyCells::Gap, // "gap" or default
                };
            }
            b"markers" => {
                builder.show_markers = Some(value == "1" || value == "true");
            }
            b"high" => {
                builder.show_high = Some(value == "1" || value == "true");
            }
            b"low" => {
                builder.show_low = Some(value == "1" || value == "true");
            }
            b"first" => {
                builder.show_first = Some(value == "1" || value == "true");
            }
            b"last" => {
                builder.show_last = Some(value == "1" || value == "true");
            }
            b"negative" => {
                builder.show_negative = Some(value == "1" || value == "true");
            }
            b"displayXAxis" => {
                builder.show_axis = Some(value == "1" || value == "true");
            }
            b"displayHidden" => {
                builder.display_hidden = Some(value == "1" || value == "true");
            }
            b"rightToLeft" => {
                builder.right_to_left = Some(value == "1" || value == "true");
            }
            b"lineWeight" => {
                builder.line_weight = value.parse().ok();
            }
            b"minAxisType" => {
                builder.min_axis_type = Some(parse_axis_type(value));
            }
            b"maxAxisType" => {
                builder.max_axis_type = Some(parse_axis_type(value));
            }
            b"manualMin" => {
                builder.manual_min = value.parse().ok();
            }
            b"manualMax" => {
                builder.manual_max = value.parse().ok();
            }
            b"dateAxis" => {
                builder.date_axis = Some(value == "1" || value == "true");
            }
            _ => {}
        }
    }

    builder
}

/// Parse axis type from string
fn parse_axis_type(value: &str) -> SparklineAxisType {
    match value {
        "group" => SparklineAxisType::Group,
        "custom" => SparklineAxisType::Custom,
        _ => SparklineAxisType::Individual, // "individual" or default
    }
}

/// Parse a color element (colorSeries, colorNegative, etc.) and resolve it
fn parse_color_element(
    e: &BytesStart,
    theme_colors: &[String],
    indexed_colors: Option<&Vec<String>>,
) -> Option<String> {
    let color_spec = parse_color_attrs(e);
    resolve_color(&color_spec, theme_colors, indexed_colors)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::bool_assert_comparison,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_axis_type() {
        assert_eq!(parse_axis_type("individual"), SparklineAxisType::Individual);
        assert_eq!(parse_axis_type("group"), SparklineAxisType::Group);
        assert_eq!(parse_axis_type("custom"), SparklineAxisType::Custom);
        assert_eq!(parse_axis_type("unknown"), SparklineAxisType::Individual);
    }

    #[test]
    fn test_sparkline_builder() {
        let builder = SparklineBuilder {
            data_range: Some("Sheet1!A1:A10".to_string()),
            location: Some("B1".to_string()),
        };
        let sparkline = builder.build().expect("should build sparkline");
        assert_eq!(sparkline.data_range, "Sheet1!A1:A10");
        assert_eq!(sparkline.location, "B1");
    }

    #[test]
    fn test_sparkline_builder_incomplete() {
        let builder = SparklineBuilder {
            data_range: Some("Sheet1!A1:A10".to_string()),
            location: None,
        };
        assert!(builder.build().is_none());
    }

    #[test]
    fn test_parse_sparkline_group_attrs() {
        let xml =
            br#"<sparklineGroup type="column" displayEmptyCellsAs="zero" markers="1" high="1"/>"#;
        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let mut buf = Vec::new();
        if let Ok(Event::Empty(ref e)) = reader.read_event_into(&mut buf) {
            let builder = parse_sparkline_group_attrs(e);
            assert_eq!(builder.sparkline_type, SparklineType::Column);
            assert_eq!(builder.display_empty_cells_as, SparklineEmptyCells::Zero);
            assert_eq!(builder.show_markers, Some(true));
            assert_eq!(builder.show_high, Some(true));
        }
    }

    #[test]
    fn test_parse_ext_sparklines_basic() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                <x14:sparklineGroup type="line" displayEmptyCellsAs="gap">
                    <x14:colorSeries rgb="FF376092"/>
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!A1:A10</xm:f>
                            <xm:sqref>B1</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec!["#000000".to_string(), "#FFFFFF".to_string()];

        // Skip to start of ext element
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];
                        assert_eq!(group.sparkline_type, "line");
                        assert_eq!(group.display_empty_cells_as, Some("gap".to_string()));
                        assert_eq!(group.colors.series, Some("#376092".to_string()));
                        assert_eq!(group.sparklines.len(), 1);
                        assert_eq!(group.sparklines[0].data_range, "Sheet1!A1:A10");
                        assert_eq!(group.sparklines[0].location, "B1");
                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_sparkline_group_with_all_colors() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="column" displayEmptyCellsAs="zero" negative="1" high="1" low="1" first="1" last="1">
                    <x14:colorSeries theme="4"/>
                    <x14:colorNegative theme="5"/>
                    <x14:colorAxis rgb="FF000000"/>
                    <x14:colorMarkers rgb="FFFF0000"/>
                    <x14:colorFirst rgb="FF00FF00"/>
                    <x14:colorLast rgb="FF0000FF"/>
                    <x14:colorHigh rgb="FFFFFF00"/>
                    <x14:colorLow rgb="FFFF00FF"/>
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!B1:B10</xm:f>
                            <xm:sqref>A1</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec![
            "#000000".to_string(), // dk1
            "#FFFFFF".to_string(), // lt1
            "#44546A".to_string(), // dk2
            "#E7E6E6".to_string(), // lt2
            "#4472C4".to_string(), // accent1 (theme 4)
            "#ED7D31".to_string(), // accent2 (theme 5)
        ];

        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];

                        // Check type and settings
                        assert_eq!(group.sparkline_type, "column");
                        assert_eq!(group.display_empty_cells_as, Some("zero".to_string()));
                        assert_eq!(group.negative_points, true);
                        assert_eq!(group.high_point, true);
                        assert_eq!(group.low_point, true);
                        assert_eq!(group.first_point, true);
                        assert_eq!(group.last_point, true);

                        // Check colors
                        assert_eq!(group.colors.series, Some("#4472C4".to_string())); // theme 4
                        assert_eq!(group.colors.negative, Some("#ED7D31".to_string())); // theme 5
                        assert_eq!(group.colors.axis, Some("#000000".to_string()));
                        assert_eq!(group.colors.markers, Some("#FF0000".to_string()));
                        assert_eq!(group.colors.first, Some("#00FF00".to_string()));
                        assert_eq!(group.colors.last, Some("#0000FF".to_string()));
                        assert_eq!(group.colors.high, Some("#FFFF00".to_string()));
                        assert_eq!(group.colors.low, Some("#FF00FF".to_string()));

                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_multiple_sparklines_in_group() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="line">
                    <x14:colorSeries rgb="FF0000FF"/>
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!A1:A10</xm:f>
                            <xm:sqref>B1</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                            <xm:f>Sheet1!A2:A11</xm:f>
                            <xm:sqref>B2</xm:sqref>
                        </x14:sparkline>
                        <x14:sparkline>
                            <xm:f>Sheet1!A3:A12</xm:f>
                            <xm:sqref>B3</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec![];

        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];

                        assert_eq!(group.sparklines.len(), 3);
                        assert_eq!(group.sparklines[0].location, "B1");
                        assert_eq!(group.sparklines[0].data_range, "Sheet1!A1:A10");
                        assert_eq!(group.sparklines[1].location, "B2");
                        assert_eq!(group.sparklines[1].data_range, "Sheet1!A2:A11");
                        assert_eq!(group.sparklines[2].location, "B3");
                        assert_eq!(group.sparklines[2].data_range, "Sheet1!A3:A12");

                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_stacked_sparkline() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="stacked" displayEmptyCellsAs="span">
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!C1:C10</xm:f>
                            <xm:sqref>D1</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec![];

        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];

                        assert_eq!(group.sparkline_type, "stacked");
                        assert_eq!(group.display_empty_cells_as, Some("span".to_string()));

                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_sparkline_with_axis_settings() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="line" minAxisType="custom" maxAxisType="group" manualMin="-10" manualMax="100" displayXAxis="1" lineWeight="1.5">
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!E1:E10</xm:f>
                            <xm:sqref>F1</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec![];

        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];

                        assert_eq!(group.min_axis_type, Some("custom".to_string()));
                        assert_eq!(group.max_axis_type, Some("group".to_string()));
                        assert_eq!(group.manual_min, Some(-10.0));
                        assert_eq!(group.manual_max, Some(100.0));
                        assert_eq!(group.display_x_axis, true);
                        assert_eq!(group.line_weight, Some(1.5));

                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }

    #[test]
    fn test_parse_sparkline_with_display_hidden() {
        let xml = br#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="line" displayHidden="1" rightToLeft="1">
                    <x14:sparklines>
                        <x14:sparkline>
                            <xm:f>Sheet1!G1:G10</xm:f>
                            <xm:sqref>H1</xm:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>"#;

        let mut reader = Reader::from_reader(&xml[..]);
        reader.trim_text(true);

        let theme_colors: Vec<String> = vec![];

        let mut buf = Vec::new();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"ext" {
                        let groups = parse_ext_sparklines(e, &mut reader, &theme_colors, None);
                        assert_eq!(groups.len(), 1);
                        let group = &groups[0];

                        assert_eq!(group.display_hidden, Some(true));
                        assert_eq!(group.right_to_left, true);

                        break;
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
    }
}

//! Chart parsing module
//!
//! This module handles parsing of charts from XLSX files.
#![allow(clippy::indexing_slicing)] // Safe: vectors are pre-extended before indexing
//! Charts are stored in xl/charts/chartN.xml files and referenced
//! from drawing XML files via relationships.
//!
//! # XLSX Chart Structure
//!
//! Charts in XLSX files follow this structure:
//! - Drawing XML (xl/drawings/drawingN.xml) contains anchor positions and references to charts
//! - Chart XML (xl/charts/chartN.xml) contains the actual chart definition
//! - Relationships (xl/drawings/_rels/drawingN.xml.rels) map rIds to chart paths
//!
//! ## Chart XML Structure
//! ```xml
//! <c:chartSpace xmlns:c="...">
//!   <c:chart>
//!     <c:title><c:tx><c:rich>...</c:rich></c:tx></c:title>
//!     <c:plotArea>
//!       <c:barChart>  <!-- or lineChart, pieChart, etc. -->
//!         <c:barDir val="col"/>
//!         <c:grouping val="clustered"/>
//!         <c:ser>
//!           <c:idx val="0"/>
//!           <c:order val="0"/>
//!           <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>
//!           <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f></c:strRef></c:cat>
//!           <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:val>
//!         </c:ser>
//!       </c:barChart>
//!       <c:catAx>...</c:catAx>
//!       <c:valAx>...</c:valAx>
//!     </c:plotArea>
//!     <c:legend>...</c:legend>
//!   </c:chart>
//! </c:chartSpace>
//! ```

use crate::types::{
    BarDirection, Chart, ChartAxis, ChartDataRef, ChartGrouping, ChartLegend, ChartSeries,
    ChartType,
};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use zip::ZipArchive;

/// Parse a chart from the chart XML file
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `chart_path` - Path to the chart file (e.g., "xl/charts/chart1.xml")
///
/// # Returns
/// Parsed Chart object, or None if parsing fails
pub fn parse_chart<R: Read + Seek>(archive: &mut ZipArchive<R>, chart_path: &str) -> Option<Chart> {
    let normalized_path = chart_path.trim_start_matches('/');

    let Ok(file) = archive.by_name(normalized_path) else {
        return None;
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    // Chart parsing state
    let mut chart = ChartBuilder::default();
    let mut in_chart = false;
    let mut in_plot_area = false;
    let mut in_title = false;
    let mut in_legend = false;
    let mut current_chart_type: Option<ChartType> = None;
    let mut current_series: Option<SeriesBuilder> = None;
    let mut current_axis: Option<AxisBuilder> = None;
    let mut in_ser = false;
    let mut in_cat = false;
    let mut in_val = false;
    let mut in_x_val = false;
    let mut in_bubble_size = false;
    let mut in_tx = false; // series name
    let mut in_str_ref = false;
    let mut in_num_ref = false;
    let mut in_f = false; // formula
    let mut in_v = false; // value
    let mut in_pt = false; // point
    let mut current_pt_idx: Option<u32> = None;
    let mut in_axis_title = false;
    let mut in_scaling = false;
    let mut title_text = String::new();
    let mut axis_title_text = String::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "chart" => in_chart = true,
                    "plotArea" if in_chart => in_plot_area = true,
                    "title" if in_chart && !in_plot_area && !in_ser => in_title = true,
                    "title" if current_axis.is_some() => in_axis_title = true,
                    "legend" if in_chart => {
                        in_legend = true;
                        chart.legend = Some(ChartLegend {
                            position: "r".to_string(), // default
                            overlay: false,
                        });
                    }
                    "legendPos" if in_legend => {
                        if let Some(ref mut legend) = chart.legend {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    legend.position =
                                        std::str::from_utf8(&attr.value).unwrap_or("r").to_string();
                                }
                            }
                        }
                    }
                    "overlay" if in_legend => {
                        if let Some(ref mut legend) = chart.legend {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    legend.overlay =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                            }
                        }
                    }

                    // Chart type elements
                    "barChart" | "bar3DChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Bar);
                        chart.chart_type = ChartType::Bar;
                    }
                    "lineChart" | "line3DChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Line);
                        chart.chart_type = ChartType::Line;
                    }
                    "pieChart" | "pie3DChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Pie);
                        chart.chart_type = ChartType::Pie;
                    }
                    "areaChart" | "area3DChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Area);
                        chart.chart_type = ChartType::Area;
                    }
                    "scatterChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Scatter);
                        chart.chart_type = ChartType::Scatter;
                    }
                    "doughnutChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Doughnut);
                        chart.chart_type = ChartType::Doughnut;
                    }
                    "radarChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Radar);
                        chart.chart_type = ChartType::Radar;
                    }
                    "bubbleChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Bubble);
                        chart.chart_type = ChartType::Bubble;
                    }
                    "stockChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Stock);
                        chart.chart_type = ChartType::Stock;
                    }
                    "surfaceChart" | "surface3DChart" if in_plot_area => {
                        current_chart_type = Some(ChartType::Surface);
                        chart.chart_type = ChartType::Surface;
                    }

                    // Bar direction
                    "barDir" if current_chart_type == Some(ChartType::Bar) => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let val = std::str::from_utf8(&attr.value).unwrap_or("col");
                                chart.bar_direction = Some(match val {
                                    "bar" => BarDirection::Bar,
                                    _ => BarDirection::Col,
                                });
                            }
                        }
                    }

                    // Grouping
                    "grouping" if current_chart_type.is_some() => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                let val = std::str::from_utf8(&attr.value).unwrap_or("standard");
                                chart.grouping = Some(match val {
                                    "stacked" => ChartGrouping::Stacked,
                                    "percentStacked" => ChartGrouping::PercentStacked,
                                    "clustered" => ChartGrouping::Clustered,
                                    _ => ChartGrouping::Standard,
                                });
                            }
                        }
                    }

                    // Vary colors
                    "varyColors" if current_chart_type.is_some() => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                chart.vary_colors =
                                    Some(std::str::from_utf8(&attr.value).unwrap_or("0") != "0");
                            }
                        }
                    }

                    // Series
                    "ser" if current_chart_type.is_some() => {
                        in_ser = true;
                        current_series = Some(SeriesBuilder::default());
                    }
                    "idx" if in_ser && current_series.is_some() => {
                        if let Some(ref mut ser) = current_series {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    ser.idx = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                            }
                        }
                    }
                    "order" if in_ser && current_series.is_some() => {
                        if let Some(ref mut ser) = current_series {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    ser.order = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                            }
                        }
                    }
                    "tx" if in_ser => in_tx = true,
                    "cat" if in_ser => in_cat = true,
                    "val" if in_ser => in_val = true,
                    "xVal" if in_ser => in_x_val = true,
                    "bubbleSize" if in_ser => in_bubble_size = true,
                    "strRef" => in_str_ref = true,
                    "numRef" => in_num_ref = true,
                    "f" if in_str_ref || in_num_ref => in_f = true,
                    "v" if in_pt => in_v = true,
                    "pt" if in_str_ref || in_num_ref => {
                        in_pt = true;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"idx" {
                                current_pt_idx = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                            }
                        }
                    }

                    // Axes
                    "catAx" if in_plot_area => {
                        current_axis = Some(AxisBuilder {
                            axis_type: "cat".to_string(),
                            ..Default::default()
                        });
                    }
                    "valAx" if in_plot_area => {
                        current_axis = Some(AxisBuilder {
                            axis_type: "val".to_string(),
                            ..Default::default()
                        });
                    }
                    "dateAx" if in_plot_area => {
                        current_axis = Some(AxisBuilder {
                            axis_type: "date".to_string(),
                            ..Default::default()
                        });
                    }
                    "serAx" if in_plot_area => {
                        current_axis = Some(AxisBuilder {
                            axis_type: "ser".to_string(),
                            ..Default::default()
                        });
                    }
                    "axId" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                }
                            }
                        }
                    }
                    "axPos" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.position = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "scaling" if current_axis.is_some() => in_scaling = true,
                    "min" if in_scaling && current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.min = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }
                    "max" if in_scaling && current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.max = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }
                    "majorUnit" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.major_unit = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }
                    "minorUnit" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.minor_unit = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }
                    "majorGridlines" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            axis.major_gridlines = true;
                        }
                    }
                    "minorGridlines" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            axis.minor_gridlines = true;
                        }
                    }
                    "crossAx" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.crosses_ax = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                            }
                        }
                    }
                    "numFmt" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"formatCode" {
                                    axis.num_fmt = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "delete" if current_axis.is_some() => {
                        if let Some(ref mut axis) = current_axis {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    axis.deleted =
                                        std::str::from_utf8(&attr.value).unwrap_or("0") == "1";
                                }
                            }
                        }
                    }

                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Ok(text) = e.unescape() {
                    let text_str = text.trim();

                    if in_f && current_series.is_some() {
                        if let Some(ref mut ser) = current_series {
                            if in_tx {
                                ser.name_ref = Some(text_str.to_string());
                            } else if in_cat {
                                ser.categories
                                    .get_or_insert_with(ChartDataRef::default)
                                    .formula = Some(text_str.to_string());
                            } else if in_val {
                                ser.values.get_or_insert_with(ChartDataRef::default).formula =
                                    Some(text_str.to_string());
                            } else if in_x_val {
                                ser.x_values
                                    .get_or_insert_with(ChartDataRef::default)
                                    .formula = Some(text_str.to_string());
                            } else if in_bubble_size {
                                ser.bubble_sizes
                                    .get_or_insert_with(ChartDataRef::default)
                                    .formula = Some(text_str.to_string());
                            }
                        }
                    } else if in_v && in_pt && current_series.is_some() {
                        if let Some(ref mut ser) = current_series {
                            let idx = current_pt_idx.unwrap_or(0) as usize;

                            if in_str_ref {
                                // String value (category or series name)
                                if in_tx {
                                    ser.name = Some(text_str.to_string());
                                } else if in_cat {
                                    let cat =
                                        ser.categories.get_or_insert_with(ChartDataRef::default);
                                    // Extend vector if needed
                                    while cat.str_values.len() <= idx {
                                        cat.str_values.push(String::new());
                                    }
                                    cat.str_values[idx] = text_str.to_string();
                                }
                            } else if in_num_ref {
                                // Numeric value
                                let num_val = text_str.parse::<f64>().ok();
                                if in_val {
                                    let val = ser.values.get_or_insert_with(ChartDataRef::default);
                                    while val.num_values.len() <= idx {
                                        val.num_values.push(None);
                                    }
                                    val.num_values[idx] = num_val;
                                } else if in_x_val {
                                    let x = ser.x_values.get_or_insert_with(ChartDataRef::default);
                                    while x.num_values.len() <= idx {
                                        x.num_values.push(None);
                                    }
                                    x.num_values[idx] = num_val;
                                } else if in_bubble_size {
                                    let b =
                                        ser.bubble_sizes.get_or_insert_with(ChartDataRef::default);
                                    while b.num_values.len() <= idx {
                                        b.num_values.push(None);
                                    }
                                    b.num_values[idx] = num_val;
                                }
                            }
                        }
                    } else if in_title && !in_axis_title {
                        // Chart title text
                        if !text_str.is_empty() {
                            if !title_text.is_empty() {
                                title_text.push(' ');
                            }
                            title_text.push_str(text_str);
                        }
                    } else if in_axis_title && current_axis.is_some() {
                        // Axis title text
                        if !text_str.is_empty() {
                            if !axis_title_text.is_empty() {
                                axis_title_text.push(' ');
                            }
                            axis_title_text.push_str(text_str);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "chart" => in_chart = false,
                    "plotArea" => in_plot_area = false,
                    "title" if !in_axis_title => {
                        in_title = false;
                        if !title_text.is_empty() {
                            chart.title = Some(title_text.clone());
                            title_text.clear();
                        }
                    }
                    "title" if in_axis_title => {
                        in_axis_title = false;
                        if !axis_title_text.is_empty() {
                            if let Some(ref mut axis) = current_axis {
                                axis.title = Some(axis_title_text.clone());
                            }
                            axis_title_text.clear();
                        }
                    }
                    "legend" => in_legend = false,

                    // Chart type elements end
                    "barChart" | "bar3DChart" | "lineChart" | "line3DChart" | "pieChart"
                    | "pie3DChart" | "areaChart" | "area3DChart" | "scatterChart"
                    | "doughnutChart" | "radarChart" | "bubbleChart" | "stockChart"
                    | "surfaceChart" | "surface3DChart" => {
                        current_chart_type = None;
                    }

                    "ser" => {
                        if let Some(ser_builder) = current_series.take() {
                            chart.series.push(ser_builder.build());
                        }
                        in_ser = false;
                    }
                    "tx" => in_tx = false,
                    "cat" => in_cat = false,
                    "val" => in_val = false,
                    "xVal" => in_x_val = false,
                    "bubbleSize" => in_bubble_size = false,
                    "strRef" => in_str_ref = false,
                    "numRef" => in_num_ref = false,
                    "f" => in_f = false,
                    "v" => in_v = false,
                    "pt" => {
                        in_pt = false;
                        current_pt_idx = None;
                    }

                    // Axes end
                    "catAx" | "valAx" | "dateAx" | "serAx" => {
                        if let Some(axis_builder) = current_axis.take() {
                            chart.axes.push(axis_builder.build());
                        }
                    }
                    "scaling" => in_scaling = false,

                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Some(chart.build())
}

/// Get chart file paths from drawing relationships
///
/// Parses xl/drawings/_rels/drawingN.xml.rels to get chart paths
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `drawing_path` - Path to the drawing file (e.g., "xl/drawings/drawing1.xml")
///
/// # Returns
/// HashMap mapping rId -> chart path (e.g., {"rId1" -> "xl/charts/chart1.xml"})
pub fn get_chart_paths<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let mut paths = HashMap::new();

    let drawing_path = drawing_path.trim_start_matches('/');

    // Build rels path
    let rels_path = if let Some(pos) = drawing_path.rfind('/') {
        let dir = &drawing_path[..pos];
        let filename = &drawing_path[pos + 1..];
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{drawing_path}.rels")
    };

    let Ok(file) = archive.by_name(&rels_path) else {
        return paths;
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e) | Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = String::new();
                    let mut target = String::new();
                    let mut rel_type = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => {
                                id = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Target" => {
                                target = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            b"Type" => {
                                rel_type =
                                    std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                            }
                            _ => {}
                        }
                    }

                    // Check if this is a chart relationship
                    if !id.is_empty() && !target.is_empty() && rel_type.contains("chart") {
                        // Resolve relative path
                        let base_dir = if let Some(pos) = drawing_path.rfind('/') {
                            &drawing_path[..pos]
                        } else {
                            ""
                        };

                        let full_path = resolve_relative_path(base_dir, &target);
                        paths.insert(id, full_path);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    paths
}

/// Parse chart references from drawing XML
///
/// Extracts chart references and their anchor positions from drawingN.xml
///
/// # Arguments
/// * `archive` - The ZIP archive containing the XLSX file
/// * `drawing_path` - Path to the drawing file
///
/// # Returns
/// Vector of (rId, from_col, from_row, to_col, to_row, name)
pub fn parse_chart_refs_from_drawing<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    drawing_path: &str,
) -> Vec<ChartRef> {
    let mut refs = Vec::new();

    let normalized_path = drawing_path.trim_start_matches('/');

    let Ok(file) = archive.by_name(normalized_path) else {
        return refs;
    };

    let reader = BufReader::new(file);
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);

    let mut buf = Vec::new();

    // Parsing state
    let mut current_ref: Option<ChartRefBuilder> = None;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_graphic_frame = false;
    let mut current_element: Option<String> = None;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "twoCellAnchor" | "oneCellAnchor" => {
                        current_ref = Some(ChartRefBuilder::default());
                    }
                    "from" => in_from = true,
                    "to" => in_to = true,
                    "graphicFrame" => in_graphic_frame = true,
                    "cNvPr" if in_graphic_frame => {
                        if let Some(ref mut r) = current_ref {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"name" {
                                    r.name = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "chart" if in_graphic_frame => {
                        if let Some(ref mut r) = current_ref {
                            for attr in e.attributes().flatten() {
                                let key = attr.key.as_ref();
                                if key == b"r:id"
                                    || key == b"id"
                                    || (key.len() > 3 && key.ends_with(b":id"))
                                {
                                    r.r_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .map(ToString::to_string);
                                }
                            }
                        }
                    }
                    "col" | "row" => {
                        current_element = Some(name.to_string());
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                if name == "chart" && in_graphic_frame {
                    if let Some(ref mut r) = current_ref {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            if key == b"r:id"
                                || key == b"id"
                                || (key.len() > 3 && key.ends_with(b":id"))
                            {
                                r.r_id = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                    }
                } else if name == "cNvPr" && in_graphic_frame {
                    if let Some(ref mut r) = current_ref {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"name" {
                                r.name = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .map(ToString::to_string);
                            }
                        }
                    }
                }
            }
            Ok(Event::Text(ref e)) => {
                if let (Some(ref element), Some(ref mut r)) = (&current_element, &mut current_ref) {
                    if let Ok(text) = e.unescape() {
                        let value: u32 = text.parse().unwrap_or(0);
                        match element.as_str() {
                            "col" if in_from => r.from_col = Some(value),
                            "row" if in_from => r.from_row = Some(value),
                            "col" if in_to => r.to_col = Some(value),
                            "row" if in_to => r.to_row = Some(value),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local_name = e.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).unwrap_or("");

                match name {
                    "twoCellAnchor" | "oneCellAnchor" => {
                        if let Some(builder) = current_ref.take() {
                            if let Some(chart_ref) = builder.build() {
                                refs.push(chart_ref);
                            }
                        }
                    }
                    "from" => in_from = false,
                    "to" => in_to = false,
                    "graphicFrame" => in_graphic_frame = false,
                    "col" | "row" => current_element = None,
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    refs
}

/// Chart reference from drawing XML
#[derive(Debug, Clone)]
pub struct ChartRef {
    /// Relationship ID (e.g., "rId1")
    pub r_id: String,
    /// Starting column (0-indexed)
    pub from_col: u32,
    /// Starting row (0-indexed)
    pub from_row: u32,
    /// Ending column (0-indexed) - None for oneCellAnchor
    pub to_col: Option<u32>,
    /// Ending row (0-indexed) - None for oneCellAnchor
    pub to_row: Option<u32>,
    /// Chart name from drawing
    pub name: Option<String>,
}

/// Builder for ChartRef
#[derive(Debug, Default)]
struct ChartRefBuilder {
    r_id: Option<String>,
    from_col: Option<u32>,
    from_row: Option<u32>,
    to_col: Option<u32>,
    to_row: Option<u32>,
    name: Option<String>,
}

impl ChartRefBuilder {
    fn build(self) -> Option<ChartRef> {
        // Need at least the rId to be a valid chart reference
        let r_id = self.r_id?;

        Some(ChartRef {
            r_id,
            from_col: self.from_col.unwrap_or(0),
            from_row: self.from_row.unwrap_or(0),
            to_col: self.to_col, // Keep as Option - None for oneCellAnchor
            to_row: self.to_row, // Keep as Option - None for oneCellAnchor
            name: self.name,
        })
    }
}

/// Builder for Chart during parsing
#[derive(Debug, Default)]
struct ChartBuilder {
    chart_type: ChartType,
    bar_direction: Option<BarDirection>,
    grouping: Option<ChartGrouping>,
    title: Option<String>,
    series: Vec<ChartSeries>,
    axes: Vec<ChartAxis>,
    legend: Option<ChartLegend>,
    vary_colors: Option<bool>,
}

impl ChartBuilder {
    fn build(self) -> Chart {
        Chart {
            chart_type: self.chart_type,
            bar_direction: self.bar_direction,
            grouping: self.grouping,
            title: self.title,
            series: self.series,
            axes: self.axes,
            legend: self.legend,
            vary_colors: self.vary_colors,
            from_col: None,
            from_row: None,
            to_col: None,
            to_row: None,
            name: None,
        }
    }
}

/// Builder for ChartSeries during parsing
#[derive(Debug, Default)]
struct SeriesBuilder {
    idx: u32,
    order: u32,
    name: Option<String>,
    name_ref: Option<String>,
    categories: Option<ChartDataRef>,
    values: Option<ChartDataRef>,
    x_values: Option<ChartDataRef>,
    bubble_sizes: Option<ChartDataRef>,
    fill_color: Option<String>,
    line_color: Option<String>,
}

impl SeriesBuilder {
    fn build(self) -> ChartSeries {
        ChartSeries {
            idx: self.idx,
            order: self.order,
            name: self.name,
            name_ref: self.name_ref,
            categories: self.categories,
            values: self.values,
            x_values: self.x_values,
            bubble_sizes: self.bubble_sizes,
            fill_color: self.fill_color,
            line_color: self.line_color,
            series_type: None,
        }
    }
}

/// Builder for ChartAxis during parsing
#[derive(Debug, Default)]
struct AxisBuilder {
    id: u32,
    axis_type: String,
    position: Option<String>,
    title: Option<String>,
    min: Option<f64>,
    max: Option<f64>,
    major_unit: Option<f64>,
    minor_unit: Option<f64>,
    major_gridlines: bool,
    minor_gridlines: bool,
    crosses_ax: Option<u32>,
    num_fmt: Option<String>,
    deleted: bool,
}

impl AxisBuilder {
    fn build(self) -> ChartAxis {
        ChartAxis {
            id: self.id,
            axis_type: self.axis_type,
            position: self.position,
            title: self.title,
            min: self.min,
            max: self.max,
            major_unit: self.major_unit,
            minor_unit: self.minor_unit,
            major_gridlines: self.major_gridlines,
            minor_gridlines: self.minor_gridlines,
            crosses_ax: self.crosses_ax,
            num_fmt: self.num_fmt,
            deleted: self.deleted,
        }
    }
}

/// Resolve a relative path against a base directory
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    // If path is absolute (starts with /), just remove the leading slash
    if let Some(stripped) = relative.strip_prefix('/') {
        return stripped.to_string();
    }

    // Split base directory into components
    let mut components: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();

    // Process relative path
    for part in relative.split('/') {
        match part {
            ".." => {
                components.pop();
            }
            "." | "" => {}
            _ => components.push(part),
        }
    }

    components.join("/")
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_relative_path() {
        assert_eq!(
            resolve_relative_path("xl/drawings", "../charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );

        assert_eq!(
            resolve_relative_path("xl/drawings", "chart1.xml"),
            "xl/drawings/chart1.xml"
        );

        assert_eq!(
            resolve_relative_path("xl/drawings", "/xl/charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn test_chart_ref_builder() {
        let builder = ChartRefBuilder {
            r_id: Some("rId1".to_string()),
            from_col: Some(0),
            from_row: Some(0),
            to_col: Some(5),
            to_row: Some(10),
            name: Some("Chart 1".to_string()),
        };

        let chart_ref = builder.build().unwrap();
        assert_eq!(chart_ref.r_id, "rId1");
        assert_eq!(chart_ref.from_col, 0);
        assert_eq!(chart_ref.from_row, 0);
        assert_eq!(chart_ref.to_col, Some(5));
        assert_eq!(chart_ref.to_row, Some(10));
        assert_eq!(chart_ref.name, Some("Chart 1".to_string()));
    }

    #[test]
    fn test_chart_ref_builder_missing_rid() {
        let builder = ChartRefBuilder {
            r_id: None,
            from_col: Some(0),
            from_row: Some(0),
            ..Default::default()
        };

        assert!(builder.build().is_none());
    }

    #[test]
    fn test_series_builder() {
        let builder = SeriesBuilder {
            idx: 0,
            order: 0,
            name: Some("Sales".to_string()),
            name_ref: Some("Sheet1!$A$1".to_string()),
            values: Some(ChartDataRef {
                formula: Some("Sheet1!$B$2:$B$5".to_string()),
                num_values: vec![Some(10.0), Some(20.0), Some(30.0), Some(40.0)],
                str_values: Vec::new(),
            }),
            ..Default::default()
        };

        let series = builder.build();
        assert_eq!(series.idx, 0);
        assert_eq!(series.name, Some("Sales".to_string()));
        assert!(series.values.is_some());

        let values = series.values.unwrap();
        assert_eq!(values.formula, Some("Sheet1!$B$2:$B$5".to_string()));
        assert_eq!(values.num_values.len(), 4);
    }

    #[test]
    fn test_axis_builder() {
        let builder = AxisBuilder {
            id: 1,
            axis_type: "val".to_string(),
            position: Some("l".to_string()),
            title: Some("Values".to_string()),
            min: Some(0.0),
            max: Some(100.0),
            major_gridlines: true,
            ..Default::default()
        };

        let axis = builder.build();
        assert_eq!(axis.id, 1);
        assert_eq!(axis.axis_type, "val");
        assert_eq!(axis.position, Some("l".to_string()));
        assert_eq!(axis.title, Some("Values".to_string()));
        assert_eq!(axis.min, Some(0.0));
        assert_eq!(axis.max, Some(100.0));
        assert!(axis.major_gridlines);
    }

    #[test]
    fn test_chart_builder() {
        let builder = ChartBuilder {
            chart_type: ChartType::Bar,
            bar_direction: Some(BarDirection::Col),
            grouping: Some(ChartGrouping::Clustered),
            title: Some("Sales Chart".to_string()),
            series: vec![ChartSeries {
                idx: 0,
                order: 0,
                name: Some("Q1".to_string()),
                name_ref: None,
                categories: None,
                values: Some(ChartDataRef {
                    formula: Some("Sheet1!$B$2:$B$5".to_string()),
                    num_values: Vec::new(),
                    str_values: Vec::new(),
                }),
                x_values: None,
                bubble_sizes: None,
                fill_color: None,
                line_color: None,
                series_type: None,
            }],
            axes: Vec::new(),
            legend: Some(ChartLegend {
                position: "r".to_string(),
                overlay: false,
            }),
            vary_colors: None,
        };

        let chart = builder.build();
        assert_eq!(chart.chart_type, ChartType::Bar);
        assert_eq!(chart.bar_direction, Some(BarDirection::Col));
        assert_eq!(chart.title, Some("Sales Chart".to_string()));
        assert_eq!(chart.series.len(), 1);
        assert!(chart.legend.is_some());
    }
}

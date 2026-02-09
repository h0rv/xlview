//! Tests for drawings/images parsing in XLSX files.
//!
//! Drawings in XLSX files include embedded images, charts, and shapes. They are stored
//! in xl/drawings/drawingN.xml files and linked via sheet relationships.
//!
//! Drawing XML structure (xl/drawings/drawing1.xml):
//! ```xml
//! <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
//!   <xdr:twoCellAnchor>
//!     <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>...</xdr:from>
//!     <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>...</xdr:to>
//!     <xdr:pic>
//!       <xdr:nvPicPr>
//!         <xdr:cNvPr id="1" name="Picture 1" descr="Alt text"/>
//!       </xdr:nvPicPr>
//!       <xdr:blipFill>
//!         <a:blip r:embed="rId1"/>
//!       </xdr:blipFill>
//!     </xdr:pic>
//!   </xdr:twoCellAnchor>
//! </xdr:wsDr>
//! ```
//!
//! Anchor types:
//! - twoCellAnchor: Image spans from one cell to another (resizes with cells)
//! - oneCellAnchor: Image anchored to one cell with fixed size
//! - absoluteAnchor: Image at absolute position (doesn't move with cells)
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

mod common;
mod fixtures;

use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

// =============================================================================
// Helper: Create base XLSX structure
// =============================================================================

fn create_base_xlsx_with_drawing(drawing_xml: &str, drawing_rels_xml: Option<&str>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    let _ = zip.start_file("[Content_Types].xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>"#);

    // _rels/.rels
    let _ = zip.start_file("_rels/.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#);

    // xl/_rels/workbook.xml.rels
    let _ = zip.start_file("xl/_rels/workbook.xml.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#);

    // xl/workbook.xml
    let _ = zip.start_file("xl/workbook.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#);

    // xl/styles.xml
    let _ = zip.start_file("xl/styles.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#);

    // xl/worksheets/_rels/sheet1.xml.rels (links to drawing)
    let _ = zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#);

    // xl/worksheets/sheet1.xml
    let _ = zip.start_file("xl/worksheets/sheet1.xml", options);
    let _ = zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Sheet with drawing</t></is></c></row>
</sheetData>
<drawing r:id="rId1"/>
</worksheet>"#);

    // xl/drawings/drawing1.xml
    let _ = zip.start_file("xl/drawings/drawing1.xml", options);
    let _ = zip.write_all(drawing_xml.as_bytes());

    // xl/drawings/_rels/drawing1.xml.rels (optional - links to images)
    if let Some(rels) = drawing_rels_xml {
        let _ = zip.start_file("xl/drawings/_rels/drawing1.xml.rels", options);
        let _ = zip.write_all(rels.as_bytes());
    }

    // Add a dummy image file
    let _ = zip.start_file("xl/media/image1.png", options);
    let _ = zip.write_all(&[0x89, 0x50, 0x4E, 0x47]); // PNG magic bytes

    let cursor = zip.finish().expect("Failed to finish ZIP");
    cursor.into_inner()
}

// =============================================================================
// Test: Two-cell anchor image (spans cells)
// =============================================================================

#[test]
fn test_two_cell_anchor_image() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>38100</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>19050</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col>
      <xdr:colOff>304800</xdr:colOff>
      <xdr:row>10</xdr:row>
      <xdr:rowOff>152400</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1" descr="Company Logo"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="2743200" cy="1524000"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    // Check drawings were parsed
    assert!(!sheet.drawings.is_empty(), "Should have drawings");
    assert_eq!(sheet.drawings.len(), 1, "Should have 1 drawing");

    let drawing = &sheet.drawings[0];

    // Check anchor type
    assert_eq!(drawing.anchor_type, "twoCellAnchor");

    // Check from position
    assert_eq!(drawing.from_col, Some(1));
    assert_eq!(drawing.from_row, Some(2));
    assert_eq!(drawing.from_col_off, Some(38100));
    assert_eq!(drawing.from_row_off, Some(19050));

    // Check to position
    assert_eq!(drawing.to_col, Some(5));
    assert_eq!(drawing.to_row, Some(10));
    assert_eq!(drawing.to_col_off, Some(304800));
    assert_eq!(drawing.to_row_off, Some(152400));

    // Check drawing type
    assert_eq!(drawing.drawing_type, "picture");

    // Check image details
    assert_eq!(drawing.name.as_deref(), Some("Picture 1"));
    assert_eq!(drawing.description.as_deref(), Some("Company Logo"));
    // image_id is resolved from rId1 to the actual path during parsing
    assert_eq!(drawing.image_id.as_deref(), Some("xl/media/image1.png"));
}

// =============================================================================
// Test: One-cell anchor image (fixed position)
// =============================================================================

#[test]
fn test_one_cell_anchor_image() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="1905000" cy="952500"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="3" name="Image 2"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 1);

    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.anchor_type, "oneCellAnchor");
    assert_eq!(drawing.from_col, Some(0));
    assert_eq!(drawing.from_row, Some(0));

    // One-cell anchor has extent (cx, cy) instead of to position
    assert_eq!(drawing.extent_cx, Some(1905000));
    assert_eq!(drawing.extent_cy, Some(952500));

    // No to position for one-cell anchor
    assert!(drawing.to_col.is_none());
    assert!(drawing.to_row.is_none());

    assert_eq!(drawing.name.as_deref(), Some("Image 2"));
}

// =============================================================================
// Test: Absolute anchor image
// =============================================================================

#[test]
fn test_absolute_anchor_image() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="914400" y="457200"/>
    <xdr:ext cx="2286000" cy="1143000"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="4" name="Logo"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 1);

    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.anchor_type, "absoluteAnchor");

    // Absolute position in EMUs
    assert_eq!(drawing.pos_x, Some(914400));
    assert_eq!(drawing.pos_y, Some(457200));

    // Extent
    assert_eq!(drawing.extent_cx, Some(2286000));
    assert_eq!(drawing.extent_cy, Some(1143000));

    // No cell references for absolute anchor
    assert!(drawing.from_col.is_none());
    assert!(drawing.from_row.is_none());

    assert_eq!(drawing.name.as_deref(), Some("Logo"));
}

// =============================================================================
// Test: Image with description/alt text
// =============================================================================

#[test]
fn test_image_with_description() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="5" name="Accessible Image" descr="A chart showing quarterly sales data for 2024. Q1: $1.2M, Q2: $1.5M, Q3: $1.8M, Q4: $2.1M." title="Quarterly Sales Chart"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let drawing = &sheet.drawings[0];

    assert_eq!(drawing.name.as_deref(), Some("Accessible Image"));
    assert_eq!(
        drawing.description.as_deref(),
        Some("A chart showing quarterly sales data for 2024. Q1: $1.2M, Q2: $1.5M, Q3: $1.8M, Q4: $2.1M.")
    );
    assert_eq!(drawing.title.as_deref(), Some("Quarterly Sales Chart"));
}

// =============================================================================
// Test: Multiple images in sheet
// =============================================================================

#[test]
fn test_multiple_images() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="Header Logo"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="914400" cy="914400"/>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Icon 1"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>25</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Main Chart"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId3"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 3, "Should have 3 drawings");

    // Check each drawing
    let names: Vec<_> = sheet
        .drawings
        .iter()
        .filter_map(|d| d.name.as_deref())
        .collect();
    assert!(names.contains(&"Header Logo"));
    assert!(names.contains(&"Icon 1"));
    assert!(names.contains(&"Main Chart"));

    // Check anchor types
    let anchor_types: Vec<_> = sheet
        .drawings
        .iter()
        .map(|d| d.anchor_type.as_str())
        .collect();
    assert_eq!(
        anchor_types
            .iter()
            .filter(|&&t| t == "twoCellAnchor")
            .count(),
        2
    );
    assert_eq!(
        anchor_types
            .iter()
            .filter(|&&t| t == "oneCellAnchor")
            .count(),
        1
    );
}

// =============================================================================
// Test: Chart placeholder
// =============================================================================

#[test]
fn test_chart_placeholder() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="6096000" cy="3810000"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 1);

    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.drawing_type, "chart");
    assert_eq!(drawing.name.as_deref(), Some("Chart 1"));
    assert_eq!(drawing.chart_id.as_deref(), Some("rId1"));

    // Chart spans cells
    assert_eq!(drawing.from_col, Some(0));
    assert_eq!(drawing.from_row, Some(0));
    assert_eq!(drawing.to_col, Some(10));
    assert_eq!(drawing.to_row, Some(20));
}

// =============================================================================
// Test: Shape placeholder
// =============================================================================

#[test]
fn test_shape_placeholder() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="Rectangle 1" descr="A blue rectangle shape"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="1219200" y="952500"/>
          <a:ext cx="2438400" cy="952500"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill>
          <a:srgbClr val="4472C4"/>
        </a:solidFill>
        <a:ln>
          <a:solidFill><a:srgbClr val="2F5496"/></a:solidFill>
        </a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:t>Shape Text</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, None);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 1);

    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.drawing_type, "shape");
    assert_eq!(drawing.name.as_deref(), Some("Rectangle 1"));
    assert_eq!(
        drawing.description.as_deref(),
        Some("A blue rectangle shape")
    );
    assert_eq!(drawing.shape_type.as_deref(), Some("rect"));

    // Shape text content
    assert_eq!(drawing.text_content.as_deref(), Some("Shape Text"));

    // Shape fill color
    assert_eq!(drawing.fill_color.as_deref(), Some("#4472C4"));
}

// =============================================================================
// Test: Image positioning (offset within cell)
// =============================================================================

#[test]
fn test_image_positioning_offsets() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>3</xdr:col>
      <xdr:colOff>152400</xdr:colOff>
      <xdr:row>7</xdr:row>
      <xdr:rowOff>76200</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>7</xdr:col>
      <xdr:colOff>457200</xdr:colOff>
      <xdr:row>15</xdr:row>
      <xdr:rowOff>228600</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="6" name="Offset Image"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let drawing = &sheet.drawings[0];

    // From position with offset
    assert_eq!(drawing.from_col, Some(3));
    assert_eq!(drawing.from_col_off, Some(152400)); // 152400 EMUs = ~0.167 inches
    assert_eq!(drawing.from_row, Some(7));
    assert_eq!(drawing.from_row_off, Some(76200)); // 76200 EMUs = ~0.083 inches

    // To position with offset
    assert_eq!(drawing.to_col, Some(7));
    assert_eq!(drawing.to_col_off, Some(457200)); // 457200 EMUs = ~0.5 inches
    assert_eq!(drawing.to_row, Some(15));
    assert_eq!(drawing.to_row_off, Some(228600)); // 228600 EMUs = ~0.25 inches

    // Edit mode
    assert_eq!(drawing.edit_as.as_deref(), Some("oneCell"));
}

// =============================================================================
// Test: Sheet without drawings
// =============================================================================

#[test]
fn test_sheet_without_drawings() {
    use fixtures::XlsxBuilder;

    let xlsx = XlsxBuilder::new()
        .add_sheet("Sheet1")
        .add_cell("A1", "No drawings here", None)
        .build();

    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");
    let sheet = &workbook.sheets[0];

    assert!(sheet.drawings.is_empty(), "Sheet should have no drawings");
}

// =============================================================================
// Test: Drawing serialization to JSON
// =============================================================================

#[test]
fn test_drawing_serialization() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="Test Image" descr="Test description"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    // Serialize to JSON
    let json = serde_json::to_value(&workbook).expect("Failed to serialize");

    // Check drawings array exists
    let drawings = &json["sheets"][0]["drawings"];
    assert!(drawings.is_array(), "drawings should be an array");
    assert_eq!(drawings.as_array().unwrap().len(), 1);

    // Check drawing structure
    let drawing = &drawings[0];
    assert_eq!(drawing["anchorType"], "twoCellAnchor");
    assert_eq!(drawing["drawingType"], "picture");
    assert_eq!(drawing["name"], "Test Image");
    assert_eq!(drawing["description"], "Test description");
    assert_eq!(drawing["fromCol"], 0);
    assert_eq!(drawing["fromRow"], 0);
    assert_eq!(drawing["toCol"], 5);
    assert_eq!(drawing["toRow"], 10);
}

// =============================================================================
// Test: Connector shape
// =============================================================================

#[test]
fn test_connector_shape() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="4" name="Connector 1"/>
        <xdr:cNvCxnSpPr/>
      </xdr:nvCxnSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="609600" y="190500"/>
          <a:ext cx="2438400" cy="762000"/>
        </a:xfrm>
        <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
        <a:ln w="12700">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:tailEnd type="triangle"/>
        </a:ln>
      </xdr:spPr>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, None);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert_eq!(sheet.drawings.len(), 1);

    let drawing = &sheet.drawings[0];
    assert_eq!(drawing.drawing_type, "connector");
    assert_eq!(drawing.name.as_deref(), Some("Connector 1"));
    assert_eq!(drawing.shape_type.as_deref(), Some("straightConnector1"));
}

// =============================================================================
// Test: Drawing with hyperlink
// Note: Hyperlink parsing in drawings is not fully implemented yet.
// This test verifies the drawing is parsed but hyperlink is not extracted.
// =============================================================================

#[test]
fn test_drawing_with_hyperlink_not_implemented() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Clickable Image">
          <a:hlinkClick r:id="rId2" tooltip="Click to visit website"/>
        </xdr:cNvPr>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let drawing = &sheet.drawings[0];

    // The drawing should be parsed successfully
    assert_eq!(drawing.name.as_deref(), Some("Clickable Image"));
    assert_eq!(drawing.drawing_type, "picture");

    // Note: Hyperlink parsing in drawings is not yet implemented
    // When implemented, this test should be updated to verify hyperlink.is_some()
    // Currently hyperlinks in drawings are not extracted
}

// =============================================================================
// Test: Group shape
// =============================================================================

#[test]
fn test_group_shape() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr>
        <xdr:cNvPr id="5" name="Group 1"/>
        <xdr:cNvGrpSpPr/>
      </xdr:nvGrpSpPr>
      <xdr:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="4876800" cy="1524000"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="4876800" cy="1524000"/>
        </a:xfrm>
      </xdr:grpSpPr>
      <xdr:sp>
        <xdr:nvSpPr><xdr:cNvPr id="6" name="Shape in Group"/><xdr:cNvSpPr/></xdr:nvSpPr>
        <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
      </xdr:sp>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, None);
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    assert!(!sheet.drawings.is_empty());

    // Find the group
    let group = sheet.drawings.iter().find(|d| d.drawing_type == "group");
    assert!(group.is_some(), "Should have a group drawing");

    let group = group.unwrap();
    assert_eq!(group.name.as_deref(), Some("Group 1"));
}

// =============================================================================
// Test: Drawing with rotation
// =============================================================================

#[test]
fn test_drawing_with_rotation() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="Rotated Image"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr>
        <a:xfrm rot="5400000" flipH="1">
          <a:off x="0" y="0"/>
          <a:ext cx="1828800" cy="1828800"/>
        </a:xfrm>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let drawing_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let xlsx = create_base_xlsx_with_drawing(drawing_xml, Some(drawing_rels));
    let workbook = xlview::parser::parse(&xlsx).expect("Failed to parse XLSX");

    let sheet = &workbook.sheets[0];
    let drawing = &sheet.drawings[0];

    // Rotation is in 1/60000 of a degree, so 5400000 = 90 degrees
    assert_eq!(drawing.rotation, Some(5400000));
    assert_eq!(drawing.flip_h, Some(true));
}

// =============================================================================
// Test: Drawing types struct
// =============================================================================

#[test]
fn test_drawing_types() {
    use xlview::types::{Drawing, Hyperlink};

    // Test Drawing struct can be created with expected fields
    let drawing = Drawing {
        anchor_type: "twoCellAnchor".to_string(),
        drawing_type: "picture".to_string(),
        name: Some("Test Image".to_string()),
        description: Some("Alt text".to_string()),
        title: None,
        from_col: Some(0),
        from_row: Some(0),
        from_col_off: Some(0),
        from_row_off: Some(0),
        to_col: Some(5),
        to_row: Some(10),
        to_col_off: Some(0),
        to_row_off: Some(0),
        pos_x: None,
        pos_y: None,
        extent_cx: None,
        extent_cy: None,
        edit_as: Some("oneCell".to_string()),
        image_id: Some("rId1".to_string()),
        chart_id: None,
        shape_type: None,
        fill_color: None,
        line_color: None,
        text_content: None,
        rotation: None,
        flip_h: None,
        flip_v: None,
        hyperlink: Some(Hyperlink {
            target: "https://example.com".to_string(),
            location: None,
            tooltip: Some("Click me".to_string()),
            is_external: true,
        }),
        xfrm_x: None,
        xfrm_y: None,
        xfrm_cx: None,
        xfrm_cy: None,
    };

    assert_eq!(drawing.anchor_type, "twoCellAnchor");
    assert_eq!(drawing.drawing_type, "picture");
    assert!(drawing.hyperlink.is_some());
}

// =============================================================================
// Tests: Real XLSX Files - kitchen_sink_v2.xlsx
// =============================================================================

/// Parse a real XLSX file and return the workbook
#[allow(clippy::expect_used)]
fn parse_real_file(path: &str) -> xlview::types::Workbook {
    let data = std::fs::read(path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    xlview::parser::parse(&data).unwrap_or_else(|_| panic!("Failed to parse XLSX file: {}", path))
}

#[test]
fn test_kitchen_sink_v2_has_drawings() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // kitchen_sink_v2.xlsx has drawings in Sheet1 (charts) and Sheet2 (images)
    // Sheet1 has 4 charts, Sheet2 has 4 images
    let total_drawings: usize = workbook.sheets.iter().map(|s| s.drawings.len()).sum();

    assert!(
        total_drawings > 0,
        "kitchen_sink_v2.xlsx should have drawings"
    );
}

#[test]
fn test_kitchen_sink_v2_embedded_images() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // kitchen_sink_v2.xlsx has 4 embedded images in xl/media/
    assert!(
        !workbook.images.is_empty(),
        "kitchen_sink_v2.xlsx should have embedded images"
    );

    // Check that we have at least 4 images
    assert!(
        workbook.images.len() >= 4,
        "Expected at least 4 images, got {}",
        workbook.images.len()
    );

    // Check that images have valid properties
    for image in &workbook.images {
        // ID should be the path
        assert!(
            image.id.contains("media/image"),
            "Image ID should contain media/image: {}",
            image.id
        );

        // Should have base64 data
        assert!(!image.data.is_empty(), "Image data should not be empty");

        // MIME type should be set (PNG for these test images)
        assert!(
            image.mime_type == "image/png" || image.mime_type == "image/jpeg",
            "Image MIME type should be image/png or image/jpeg, got: {}",
            image.mime_type
        );

        // Filename should be present
        assert!(image.filename.is_some(), "Image filename should be present");
    }
}

#[test]
fn test_kitchen_sink_v2_image_drawings_sheet2() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // Find Sheet2 (Images) - it has the image drawings
    let images_sheet = workbook.sheets.iter().find(|s| s.name == "Images");

    assert!(
        images_sheet.is_some(),
        "Should have an 'Images' sheet in kitchen_sink_v2.xlsx"
    );

    let images_sheet = images_sheet.unwrap();

    // This sheet should have 4 picture drawings
    let picture_drawings: Vec<_> = images_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "picture")
        .collect();

    assert_eq!(
        picture_drawings.len(),
        4,
        "Images sheet should have 4 picture drawings"
    );

    // Verify each image drawing has proper properties
    for (i, drawing) in picture_drawings.iter().enumerate() {
        assert_eq!(
            drawing.anchor_type,
            "oneCellAnchor",
            "Image {} should use oneCellAnchor",
            i + 1
        );
        assert_eq!(drawing.drawing_type, "picture", "Should be a picture type");

        // Should have a name like "Image 1", "Image 2", etc.
        assert!(drawing.name.is_some(), "Image {} should have a name", i + 1);

        // Should have position info
        assert!(
            drawing.from_col.is_some(),
            "Image {} should have from_col",
            i + 1
        );
        assert!(
            drawing.from_row.is_some(),
            "Image {} should have from_row",
            i + 1
        );

        // One-cell anchors should have extent (dimensions)
        assert!(
            drawing.extent_cx.is_some(),
            "Image {} should have extent_cx for oneCellAnchor",
            i + 1
        );
        assert!(
            drawing.extent_cy.is_some(),
            "Image {} should have extent_cy for oneCellAnchor",
            i + 1
        );

        // Should have image relationship ID
        assert!(
            drawing.image_id.is_some(),
            "Image {} should have image_id",
            i + 1
        );
    }
}

#[test]
fn test_kitchen_sink_v2_image_positions() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // Find the Images sheet
    let images_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Images")
        .expect("Should have Images sheet");

    let pictures: Vec<_> = images_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "picture")
        .collect();

    // Verify positions - images are at columns 1, 3, 5, 7 (0-indexed)
    // based on the drawing2.xml content
    let expected_cols = [1u32, 3, 5, 7];
    let expected_row = 2u32;

    for (i, picture) in pictures.iter().enumerate() {
        if i < expected_cols.len() {
            assert_eq!(
                picture.from_col,
                Some(expected_cols[i]),
                "Image {} should be at column {}",
                i + 1,
                expected_cols[i]
            );
            assert_eq!(
                picture.from_row,
                Some(expected_row),
                "Image {} should be at row {}",
                i + 1,
                expected_row
            );
        }
    }
}

#[test]
fn test_kitchen_sink_v2_image_dimensions() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let images_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Images")
        .expect("Should have Images sheet");

    let pictures: Vec<_> = images_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "picture")
        .collect();

    // All images have the same dimensions: 762000 x 762000 EMUs (80x80 pixels at 96 DPI)
    // 762000 EMUs = 80 pixels (at 96 DPI) = ~0.83 inches
    for picture in &pictures {
        assert_eq!(
            picture.extent_cx,
            Some(762000),
            "Image width should be 762000 EMUs"
        );
        assert_eq!(
            picture.extent_cy,
            Some(762000),
            "Image height should be 762000 EMUs"
        );
    }
}

#[test]
fn test_kitchen_sink_v2_image_descriptions() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let images_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Images")
        .expect("Should have Images sheet");

    let pictures: Vec<_> = images_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "picture")
        .collect();

    // All images have description "Picture" according to the drawing XML
    for picture in &pictures {
        assert_eq!(
            picture.description.as_deref(),
            Some("Picture"),
            "Image should have description 'Picture'"
        );
    }
}

#[test]
fn test_kitchen_sink_v2_chart_drawings() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // Find Charts sheet (Sheet1)
    let charts_sheet = workbook.sheets.iter().find(|s| s.name == "Charts");

    assert!(
        charts_sheet.is_some(),
        "Should have a 'Charts' sheet in kitchen_sink_v2.xlsx"
    );

    let charts_sheet = charts_sheet.unwrap();

    // This sheet should have 4 chart drawings
    let chart_drawings: Vec<_> = charts_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "chart")
        .collect();

    assert_eq!(
        chart_drawings.len(),
        4,
        "Charts sheet should have 4 chart drawings"
    );

    // Verify chart drawing properties
    for (i, drawing) in chart_drawings.iter().enumerate() {
        assert_eq!(
            drawing.anchor_type,
            "oneCellAnchor",
            "Chart {} should use oneCellAnchor",
            i + 1
        );
        assert_eq!(drawing.drawing_type, "chart", "Should be a chart type");

        // Charts should have chart_id (relationship to chart XML)
        assert!(
            drawing.chart_id.is_some(),
            "Chart {} should have chart_id",
            i + 1
        );

        // Should have position info
        assert!(
            drawing.from_col.is_some(),
            "Chart {} should have from_col",
            i + 1
        );
        assert!(
            drawing.from_row.is_some(),
            "Chart {} should have from_row",
            i + 1
        );

        // Charts use one-cell anchor with extent
        assert!(
            drawing.extent_cx.is_some(),
            "Chart {} should have extent_cx",
            i + 1
        );
        assert!(
            drawing.extent_cy.is_some(),
            "Chart {} should have extent_cy",
            i + 1
        );
    }
}

#[test]
fn test_kitchen_sink_v2_chart_dimensions() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let charts_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Charts")
        .expect("Should have Charts sheet");

    let charts: Vec<_> = charts_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "chart")
        .collect();

    // All charts have dimensions: 5400000 x 2700000 EMUs
    for chart in &charts {
        assert_eq!(
            chart.extent_cx,
            Some(5400000),
            "Chart width should be 5400000 EMUs"
        );
        assert_eq!(
            chart.extent_cy,
            Some(2700000),
            "Chart height should be 2700000 EMUs"
        );
    }
}

// =============================================================================
// Tests: Real XLSX Files - ms_cf_samples.xlsx (has many drawings/images)
// =============================================================================

#[test]
fn test_ms_cf_samples_has_drawings() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    // ms_cf_samples.xlsx has 16 drawings with images
    let total_drawings: usize = workbook.sheets.iter().map(|s| s.drawings.len()).sum();

    assert!(
        total_drawings > 0,
        "ms_cf_samples.xlsx should have drawings"
    );
}

#[test]
fn test_ms_cf_samples_has_embedded_images() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    // ms_cf_samples.xlsx has 16 images in xl/media/
    assert!(
        !workbook.images.is_empty(),
        "ms_cf_samples.xlsx should have embedded images"
    );

    // Check that we have multiple images
    assert!(
        workbook.images.len() >= 10,
        "Expected at least 10 images, got {}",
        workbook.images.len()
    );
}

#[test]
fn test_ms_cf_samples_image_mime_types() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    for image in &workbook.images {
        // All images in ms_cf_samples.xlsx are PNG
        assert_eq!(
            image.mime_type, "image/png",
            "Image MIME type should be image/png for {}",
            image.id
        );
    }
}

#[test]
fn test_ms_cf_samples_image_data_is_valid_base64() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    use base64::{engine::general_purpose::STANDARD as BASE64, Engine};

    for image in &workbook.images {
        // Verify base64 data can be decoded
        let decoded = BASE64.decode(&image.data);
        assert!(
            decoded.is_ok(),
            "Image {} should have valid base64 data",
            image.id
        );

        let decoded = decoded.unwrap();
        assert!(
            !decoded.is_empty(),
            "Decoded image {} should not be empty",
            image.id
        );

        // Verify PNG magic bytes
        if image.mime_type == "image/png" {
            assert!(
                decoded.starts_with(&[0x89, 0x50, 0x4E, 0x47]),
                "PNG image {} should have PNG magic bytes",
                image.id
            );
        }
    }
}

#[test]
fn test_ms_cf_samples_drawing_anchor_types() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    let mut anchor_types: std::collections::HashSet<String> = std::collections::HashSet::new();

    for sheet in &workbook.sheets {
        for drawing in &sheet.drawings {
            anchor_types.insert(drawing.anchor_type.clone());
        }
    }

    // Should have at least one anchor type
    assert!(
        !anchor_types.is_empty(),
        "Should have at least one anchor type"
    );

    // Verify anchor types are valid
    for anchor_type in &anchor_types {
        assert!(
            anchor_type == "twoCellAnchor"
                || anchor_type == "oneCellAnchor"
                || anchor_type == "absoluteAnchor",
            "Invalid anchor type: {}",
            anchor_type
        );
    }
}

#[test]
fn test_ms_cf_samples_drawing_types() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    let mut drawing_types: std::collections::HashSet<String> = std::collections::HashSet::new();

    for sheet in &workbook.sheets {
        for drawing in &sheet.drawings {
            drawing_types.insert(drawing.drawing_type.clone());
        }
    }

    // Should have picture type
    assert!(
        drawing_types.contains("picture"),
        "Should have picture drawing type"
    );
}

// =============================================================================
// Tests: Drawing Parsing Does Not Panic
// =============================================================================

#[test]
fn test_kitchen_sink_v2_parsing_does_not_panic() {
    // This test ensures the parser doesn't panic on real-world XLSX files
    let result = std::panic::catch_unwind(|| {
        parse_real_file("test/kitchen_sink_v2.xlsx");
    });

    assert!(
        result.is_ok(),
        "Parsing kitchen_sink_v2.xlsx should not panic"
    );
}

#[test]
fn test_ms_cf_samples_parsing_does_not_panic() {
    // This test ensures the parser doesn't panic on real-world XLSX files
    let result = std::panic::catch_unwind(|| {
        parse_real_file("test/ms_cf_samples.xlsx");
    });

    assert!(
        result.is_ok(),
        "Parsing ms_cf_samples.xlsx should not panic"
    );
}

#[test]
fn test_kitchen_sink_original_parsing_does_not_panic() {
    // kitchen_sink.xlsx has no drawings, but should still parse fine
    let result = std::panic::catch_unwind(|| {
        parse_real_file("test/kitchen_sink.xlsx");
    });

    assert!(result.is_ok(), "Parsing kitchen_sink.xlsx should not panic");
}

#[test]
fn test_kitchen_sink_original_has_no_drawings() {
    // Verify the original kitchen_sink.xlsx has no drawings
    let workbook = parse_real_file("test/kitchen_sink.xlsx");

    let total_drawings: usize = workbook.sheets.iter().map(|s| s.drawings.len()).sum();

    assert_eq!(
        total_drawings, 0,
        "kitchen_sink.xlsx should have no drawings"
    );

    assert!(
        workbook.images.is_empty(),
        "kitchen_sink.xlsx should have no embedded images"
    );
}

// =============================================================================
// Tests: Drawing Relationship IDs
// =============================================================================

#[test]
fn test_kitchen_sink_v2_image_ids_resolved_to_paths() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let images_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Images")
        .expect("Should have Images sheet");

    let pictures: Vec<_> = images_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "picture")
        .collect();

    // Each image should have an image_id that was resolved to the actual file path
    let image_ids: Vec<_> = pictures
        .iter()
        .filter_map(|p| p.image_id.as_ref())
        .collect();

    assert_eq!(
        image_ids.len(),
        pictures.len(),
        "All pictures should have image_id"
    );

    // IDs should be unique
    let unique_ids: std::collections::HashSet<_> = image_ids.iter().collect();
    assert_eq!(
        unique_ids.len(),
        image_ids.len(),
        "Image IDs should be unique"
    );

    // IDs should be resolved paths like "xl/media/image1.png"
    for id in &image_ids {
        assert!(
            id.contains("media/image"),
            "Image ID should be resolved to path containing 'media/image': {}",
            id
        );
    }
}

#[test]
fn test_kitchen_sink_v2_chart_relationship_ids() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let charts_sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Charts")
        .expect("Should have Charts sheet");

    let charts: Vec<_> = charts_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "chart")
        .collect();

    // Each chart should have a chart_id
    let chart_ids: Vec<_> = charts.iter().filter_map(|c| c.chart_id.as_ref()).collect();

    assert_eq!(
        chart_ids.len(),
        charts.len(),
        "All charts should have chart_id"
    );

    // IDs should follow rIdN pattern
    for id in &chart_ids {
        assert!(
            id.starts_with("rId"),
            "Chart ID should start with 'rId': {}",
            id
        );
    }
}

// =============================================================================
// Tests: JSON Serialization of Drawings
// =============================================================================

#[test]
fn test_kitchen_sink_v2_drawings_serialize_to_json() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // Serialize to JSON
    let json = serde_json::to_value(&workbook).expect("Failed to serialize workbook to JSON");

    // Find the Images sheet in JSON
    let sheets = json["sheets"].as_array().expect("sheets should be array");

    let images_sheet = sheets.iter().find(|s| s["name"] == "Images");
    assert!(images_sheet.is_some(), "Should have Images sheet in JSON");

    let images_sheet = images_sheet.unwrap();
    let drawings = images_sheet["drawings"]
        .as_array()
        .expect("drawings should be array");

    assert_eq!(drawings.len(), 4, "Should have 4 drawings in Images sheet");

    // Check first drawing structure
    let first_drawing = &drawings[0];
    assert_eq!(first_drawing["anchorType"], "oneCellAnchor");
    assert_eq!(first_drawing["drawingType"], "picture");
    assert!(first_drawing["fromCol"].is_number());
    assert!(first_drawing["fromRow"].is_number());
    assert!(first_drawing["extentCx"].is_number());
    assert!(first_drawing["extentCy"].is_number());
}

#[test]
fn test_ms_cf_samples_drawings_serialize_to_json() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    // Serialize to JSON - this should not panic
    let json = serde_json::to_value(&workbook).expect("Failed to serialize workbook to JSON");

    // Verify images array exists
    let images = json["images"].as_array();
    assert!(
        images.is_some(),
        "Workbook should have images array in JSON"
    );

    let images = images.unwrap();
    assert!(!images.is_empty(), "Images array should not be empty");

    // Check first image structure
    let first_image = &images[0];
    assert!(first_image["id"].is_string());
    assert!(first_image["mimeType"].is_string());
    assert!(first_image["data"].is_string());
}

// =============================================================================
// Tests: Image Format Detection
// =============================================================================

#[test]
fn test_image_format_from_extension() {
    use xlview::types::ImageFormat;

    assert_eq!(ImageFormat::from_extension("png"), ImageFormat::Png);
    assert_eq!(ImageFormat::from_extension("PNG"), ImageFormat::Png);
    assert_eq!(ImageFormat::from_extension("jpg"), ImageFormat::Jpeg);
    assert_eq!(ImageFormat::from_extension("jpeg"), ImageFormat::Jpeg);
    assert_eq!(ImageFormat::from_extension("JPEG"), ImageFormat::Jpeg);
    assert_eq!(ImageFormat::from_extension("gif"), ImageFormat::Gif);
    assert_eq!(ImageFormat::from_extension("bmp"), ImageFormat::Bmp);
    assert_eq!(ImageFormat::from_extension("tif"), ImageFormat::Tiff);
    assert_eq!(ImageFormat::from_extension("tiff"), ImageFormat::Tiff);
    assert_eq!(ImageFormat::from_extension("webp"), ImageFormat::Webp);
    assert_eq!(ImageFormat::from_extension("emf"), ImageFormat::Emf);
    assert_eq!(ImageFormat::from_extension("wmf"), ImageFormat::Wmf);
    assert_eq!(ImageFormat::from_extension("xyz"), ImageFormat::Unknown);
}

#[test]
fn test_image_format_from_magic_bytes() {
    use xlview::types::ImageFormat;

    // PNG magic bytes
    assert_eq!(
        ImageFormat::from_magic_bytes(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
        ImageFormat::Png
    );

    // JPEG magic bytes
    assert_eq!(
        ImageFormat::from_magic_bytes(&[0xFF, 0xD8, 0xFF, 0xE0]),
        ImageFormat::Jpeg
    );

    // GIF magic bytes
    assert_eq!(ImageFormat::from_magic_bytes(b"GIF89a"), ImageFormat::Gif);
    assert_eq!(ImageFormat::from_magic_bytes(b"GIF87a"), ImageFormat::Gif);

    // BMP magic bytes
    assert_eq!(
        ImageFormat::from_magic_bytes(b"BM\x00\x00"),
        ImageFormat::Bmp
    );

    // Unknown
    assert_eq!(
        ImageFormat::from_magic_bytes(&[0x00, 0x00, 0x00]),
        ImageFormat::Unknown
    );

    // Too short
    assert_eq!(
        ImageFormat::from_magic_bytes(&[0x89, 0x50]),
        ImageFormat::Unknown
    );
}

#[test]
fn test_image_format_mime_types() {
    use xlview::types::ImageFormat;

    assert_eq!(ImageFormat::Png.mime_type(), "image/png");
    assert_eq!(ImageFormat::Jpeg.mime_type(), "image/jpeg");
    assert_eq!(ImageFormat::Gif.mime_type(), "image/gif");
    assert_eq!(ImageFormat::Bmp.mime_type(), "image/bmp");
    assert_eq!(ImageFormat::Tiff.mime_type(), "image/tiff");
    assert_eq!(ImageFormat::Webp.mime_type(), "image/webp");
    assert_eq!(ImageFormat::Emf.mime_type(), "image/x-emf");
    assert_eq!(ImageFormat::Wmf.mime_type(), "image/x-wmf");
    assert_eq!(ImageFormat::Unknown.mime_type(), "application/octet-stream");
}

// =============================================================================
// Tests: Drawing Count per Sheet
// =============================================================================

#[test]
fn test_kitchen_sink_v2_drawing_counts_per_sheet() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    let mut drawing_counts: Vec<(String, usize)> = Vec::new();

    for sheet in &workbook.sheets {
        drawing_counts.push((sheet.name.clone(), sheet.drawings.len()));
    }

    // Charts sheet should have 4 drawings (charts)
    let charts_count = drawing_counts
        .iter()
        .find(|(name, _)| name == "Charts")
        .map(|(_, count)| *count)
        .unwrap_or(0);
    assert_eq!(charts_count, 4, "Charts sheet should have 4 drawings");

    // Images sheet should have 4 drawings (images)
    let images_count = drawing_counts
        .iter()
        .find(|(name, _)| name == "Images")
        .map(|(_, count)| *count)
        .unwrap_or(0);
    assert_eq!(images_count, 4, "Images sheet should have 4 drawings");
}

// =============================================================================
// Tests: EMU (English Metric Unit) Conversions
// =============================================================================

#[test]
fn test_emu_values_are_reasonable() {
    let workbook = parse_real_file("test/kitchen_sink_v2.xlsx");

    // 1 EMU = 1/914400 inches
    // 914400 EMUs = 1 inch
    // Common values:
    // - 762000 EMUs = ~0.83 inches (typical small image)
    // - 5400000 EMUs = ~5.9 inches (typical chart width)

    for sheet in &workbook.sheets {
        for drawing in &sheet.drawings {
            // Check extent_cx (width) if present
            if let Some(cx) = drawing.extent_cx {
                // Width should be positive and less than 20 inches (reasonable max)
                assert!(
                    cx > 0 && cx < 20 * 914400,
                    "extent_cx {} EMUs is unreasonable",
                    cx
                );
            }

            // Check extent_cy (height) if present
            if let Some(cy) = drawing.extent_cy {
                // Height should be positive and less than 20 inches
                assert!(
                    cy > 0 && cy < 20 * 914400,
                    "extent_cy {} EMUs is unreasonable",
                    cy
                );
            }

            // Check column/row positions
            if let Some(col) = drawing.from_col {
                // Column should be within Excel's limit (16384 columns)
                assert!(col < 16384, "from_col {} is unreasonable", col);
            }

            if let Some(row) = drawing.from_row {
                // Row should be within Excel's limit (1048576 rows)
                assert!(row < 1048576, "from_row {} is unreasonable", row);
            }
        }
    }
}

// =============================================================================
// Shape positioning regression tests
// =============================================================================

/// Regression test: Shapes should have xfrm transform values parsed for precise positioning.
/// This test verifies that shapes in ms_cf_samples.xlsx have xfrm values that can be used
/// for accurate rendering instead of relying solely on cell anchor calculations.
///
/// Issue: Shapes were overlapping cells because the cell-anchor based positioning
/// didn't account for Excel's precise pre-calculated positions stored in xfrm.
#[test]
fn test_shape_xfrm_positioning() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");

    // ms_cf_samples.xlsx has shapes (rounded rectangles with text like "Products1")
    // on the Home sheet (first sheet)
    let home_sheet = &workbook.sheets[0];

    // Filter for shape drawings
    let shapes: Vec<_> = home_sheet
        .drawings
        .iter()
        .filter(|d| d.drawing_type == "shape")
        .collect();

    assert!(
        !shapes.is_empty(),
        "ms_cf_samples.xlsx Home sheet should have shape drawings"
    );

    // Verify that at least some shapes have xfrm values parsed
    let shapes_with_xfrm: Vec<_> = shapes
        .iter()
        .filter(|d| d.xfrm_x.is_some() && d.xfrm_y.is_some())
        .collect();

    assert!(
        !shapes_with_xfrm.is_empty(),
        "At least some shapes should have xfrm position values parsed"
    );

    // Verify that shapes with xfrm also have size values
    for shape in &shapes_with_xfrm {
        assert!(
            shape.xfrm_cx.is_some(),
            "Shape with xfrm position should have xfrm_cx (width)"
        );
        assert!(
            shape.xfrm_cy.is_some(),
            "Shape with xfrm position should have xfrm_cy (height)"
        );

        // Sanity check: xfrm values should be positive and reasonable
        let x = shape.xfrm_x.unwrap();
        let y = shape.xfrm_y.unwrap();
        let cx = shape.xfrm_cx.unwrap();
        let cy = shape.xfrm_cy.unwrap();

        // EMU values: 1 inch = 914400 EMUs, positions should be < 100 inches
        const MAX_EMU: i64 = 100 * 914400;
        assert!(
            (0..MAX_EMU).contains(&x),
            "xfrm_x {} should be reasonable",
            x
        );
        assert!(
            (0..MAX_EMU).contains(&y),
            "xfrm_y {} should be reasonable",
            y
        );
        assert!(
            (1..MAX_EMU).contains(&cx),
            "xfrm_cx {} should be positive and reasonable",
            cx
        );
        assert!(
            (1..MAX_EMU).contains(&cy),
            "xfrm_cy {} should be positive and reasonable",
            cy
        );
    }
}

/// Verify specific shape positions match expected values from the XML.
/// The "Products1" button shape in ms_cf_samples.xlsx has known xfrm values.
#[test]
fn test_shape_xfrm_specific_values() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");
    let home_sheet = &workbook.sheets[0];

    // Find the "Products1" shape by its text content
    let products1_shape = home_sheet
        .drawings
        .iter()
        .find(|d| d.drawing_type == "shape" && d.text_content.as_deref() == Some("Products1"));

    assert!(
        products1_shape.is_some(),
        "Should find Products1 shape in ms_cf_samples.xlsx"
    );

    let shape = products1_shape.unwrap();

    // From the XML, Products1 has:
    // <a:off x="752474" y="990600"/>
    // <a:ext cx="1280160" cy="140304"/>
    assert_eq!(
        shape.xfrm_x,
        Some(752474),
        "Products1 xfrm_x should match XML"
    );
    assert_eq!(
        shape.xfrm_y,
        Some(990600),
        "Products1 xfrm_y should match XML"
    );
    assert_eq!(
        shape.xfrm_cx,
        Some(1280160),
        "Products1 xfrm_cx should match XML"
    );
    assert_eq!(
        shape.xfrm_cy,
        Some(140304),
        "Products1 xfrm_cy should match XML"
    );
}

/// Debug test to verify column widths are being parsed correctly
#[test]
fn test_debug_column_widths() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");
    let home_sheet = &workbook.sheets[0];

    println!("\n=== Column Widths for ms_cf_samples.xlsx ===");
    for cw in &home_sheet.col_widths {
        let px = cw.width * 7.0;
        println!("  col {}: {:.2} chars = {:.1} px", cw.col, cw.width, px);
    }

    // Expected values from XML (converted to 0-based):
    // col 0 (A): 10.71 chars = 75 px
    // col 1 (B): 19.29 chars = 135 px
    // col 2 (C): 23.43 chars = 164 px
    // col 3 (D): 66.43 chars = 465 px

    // Find Products1 shape and print its anchor data
    if let Some(shape) = home_sheet
        .drawings
        .iter()
        .find(|d| d.text_content.as_deref() == Some("Products1"))
    {
        println!("\n=== Products1 Shape Anchor Data ===");
        println!(
            "  from_col: {:?}, from_col_off: {:?} EMUs",
            shape.from_col, shape.from_col_off
        );
        println!(
            "  to_col: {:?}, to_col_off: {:?} EMUs",
            shape.to_col, shape.to_col_off
        );
        println!("  xfrm_x: {:?}, xfrm_cx: {:?}", shape.xfrm_x, shape.xfrm_cx);

        // Calculate expected positions
        if let (Some(_from_col), Some(_to_col)) = (shape.from_col, shape.to_col) {
            let from_off_px = shape.from_col_off.unwrap_or(0) as f64 / 9525.0;
            let to_off_px = shape.to_col_off.unwrap_or(0) as f64 / 9525.0;
            println!("  from_col_off in px: {:.1}", from_off_px);
            println!("  to_col_off in px: {:.1}", to_off_px);
        }
    }

    assert!(
        home_sheet.col_widths.len() >= 4,
        "Should have at least 4 column widths"
    );
}

/// Verify that the yellow header box has xfrm values.
/// This is a oneCellAnchor shape with text "Conditionally Formatting Data: Examples and Guidelines"
#[test]
fn test_onecellanchor_shape_xfrm() {
    let workbook = parse_real_file("test/ms_cf_samples.xlsx");
    let home_sheet = &workbook.sheets[0];

    // Find the header shape by anchor type and presence of "Conditionally" text
    let header_shape = home_sheet.drawings.iter().find(|d| {
        d.anchor_type == "oneCellAnchor"
            && d.text_content
                .as_ref()
                .is_some_and(|t| t.contains("Conditionally"))
    });

    assert!(
        header_shape.is_some(),
        "Should find the header shape in ms_cf_samples.xlsx"
    );

    let shape = header_shape.unwrap();

    // From the XML:
    // <a:off x="752474" y="133350"/>
    // <a:ext cx="7610475" cy="448563"/>
    assert_eq!(shape.xfrm_x, Some(752474), "Header xfrm_x should match XML");
    assert_eq!(shape.xfrm_y, Some(133350), "Header xfrm_y should match XML");
    assert_eq!(
        shape.xfrm_cx,
        Some(7610475),
        "Header xfrm_cx should match XML"
    );
    assert_eq!(
        shape.xfrm_cy,
        Some(448563),
        "Header xfrm_cy should match XML"
    );

    // The shape should also have anchor-level extent values
    assert!(shape.extent_cx.is_some(), "Should have anchor extent_cx");
    assert!(shape.extent_cy.is_some(), "Should have anchor extent_cy");
}

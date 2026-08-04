use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::{Component, Path, PathBuf};

use base64::Engine;
use roxmltree::{Document, Node};
use serde::Serialize;
use zip::ZipArchive;

use crate::SidecarError;

const MAX_MEDIA_BYTES: u64 = 20 * 1024 * 1024;

#[derive(Clone, Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub wrap_text: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_top: Option<BorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_bottom: Option<BorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_left: Option<BorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_right: Option<BorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_diagonal: Option<BorderEdge>,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}

impl CellStyle {
    /// True when a value-less cell carrying this style is worth keeping:
    /// either the style paints something visible (fill/border), or it differs
    /// from the workbook default in formatting that takes effect the moment
    /// the user types into the cell — number format, font, alignment (#169).
    /// Comparing against the default xf keeps the payload bounded: fontId=0
    /// materializes the default font into every style, so presence alone
    /// would mark every cell as styled.
    pub fn styles_blank_cell(&self, default: &CellStyle) -> bool {
        self.fill_color.is_some()
            || self.border_top.is_some()
            || self.border_bottom.is_some()
            || self.border_left.is_some()
            || self.border_right.is_some()
            || self.border_diagonal.is_some()
            || self.number_format != default.number_format
            || self.font_family != default.font_family
            || self.font_size != default.font_size
            || self.bold != default.bold
            || self.italic != default.italic
            || self.underline != default.underline
            || self.strikethrough != default.strikethrough
            || self.font_color != default.font_color
            || self.horizontal_alignment != default.horizontal_alignment
            || self.vertical_alignment != default.vertical_alignment
            || self.indent != default.indent
            || self.wrap_text != default.wrap_text
    }
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    pub style: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DrawingAnchor {
    pub from_row: usize,
    pub from_column: usize,
    pub from_row_offset: i64,
    pub from_column_offset: i64,
    pub to_row: usize,
    pub to_column: usize,
    pub to_row_offset: i64,
    pub to_column_offset: i64,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: String,
    pub categories: Vec<String>,
    pub values: Vec<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    /// numCache formatCode of the category (or scatter X) data (#182).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub trendline: Option<String>,
    /// `c:f` range references, so the renderer can offer data-range editing.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub values_ref: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub categories_ref: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub point_colors: Option<Vec<PointColor>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub explosion_pct: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub point_explosions: Option<Vec<PointExplosion>>,
}

/// Per-point fill override from `c:dPt`, e.g. pie slice colors.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PointColor {
    pub index: u32,
    pub color: String,
}

/// Per-slice `c:dPt/c:explosion` (% of radius).
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PointExplosion {
    pub index: u32,
    pub pct: u32,
}

/// Explicit `c:scaling` bounds; absent keys mean auto.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ValueAxisBounds {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AxisTitles {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartMetadata {
    pub chart_types: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_direction: Option<String>,
    pub title: String,
    /// Always present ("none" when the legend is absent) so the editor can
    /// echo the current state back.
    pub legend: String,
    pub data_labels: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_titles: Option<AxisTitles>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grouping: Option<String>,
    /// Only emitted when a value axis exists (pie/doughnut have none).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gridlines: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value_axis: Option<ValueAxisBounds>,
    /// `c:numFmt` on the category/date axis; wins over the series-level
    /// numCache formatCode (#182).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category_axis_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gap_width_pct: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hole_size_pct: Option<u32>,
    pub series: Vec<ChartSeries>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct VisualObject {
    pub id: String,
    pub sheet_id: String,
    pub kind: String,
    pub anchor: DrawingAnchor,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart: Option<ChartMetadata>,
    /// ZIP entry path of the chart part, e.g. `xl/charts/chart1.xml`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// Degrees clockwise.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    /// ZIP entry path of the drawing part this visual lives in, plus its
    /// anchor index within that part — the save-side edit locator.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub drawing_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub drawing_index: Option<usize>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaResult {
    pub media_type: String,
    pub base64: String,
}

#[derive(Clone)]
pub struct SheetVisualSource {
    pub sheet_id: String,
    pub worksheet_path: String,
}

#[derive(Clone)]
struct Relationship {
    target: String,
    relationship_type: String,
}

#[derive(Clone, Default)]
struct FontStyle {
    family: Option<String>,
    size: Option<f64>,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    color: Option<String>,
}

#[derive(Clone, Default)]
struct BorderSet {
    top: Option<BorderEdge>,
    bottom: Option<BorderEdge>,
    left: Option<BorderEdge>,
    right: Option<BorderEdge>,
    diagonal: Option<BorderEdge>,
    diagonal_up: bool,
    diagonal_down: bool,
}

/// Theme palette in `theme` attribute index order (0↔1 and 2↔3 are swapped
/// versus the clrScheme document order, per the xlsx theme index mapping).
#[derive(Clone, Default)]
pub struct ColorContext {
    theme: Vec<(u8, u8, u8)>,
}

pub fn read_styles(
    archive: &mut ZipArchive<File>,
    colors: &ColorContext,
) -> Result<(Vec<CellStyle>, Vec<CellStyle>), SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/styles.xml")? else {
        return Ok((vec![CellStyle::default()], Vec::new()));
    };
    let document = parse_document(&xml, "styles.xml")?;
    let custom_formats = document
        .descendants()
        .filter(|node| node.has_tag_name("numFmt"))
        .filter_map(|node| {
            Some((
                node.attribute("numFmtId")?.parse::<u32>().ok()?,
                node.attribute("formatCode")?.to_owned(),
            ))
        })
        .collect::<HashMap<_, _>>();
    let fonts = document
        .descendants()
        .find(|node| node.has_tag_name("fonts"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("font"))
                .map(|font| parse_font(font, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let fills = document
        .descendants()
        .find(|node| node.has_tag_name("fills"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("fill"))
                .map(|fill| parse_fill(fill, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let borders = document
        .descendants()
        .find(|node| node.has_tag_name("borders"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("border"))
                .map(|border| parse_border(border, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let styles = document
        .descendants()
        .find(|node| node.has_tag_name("cellXfs"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("xf"))
                .map(|xf| {
                    let font = numeric_attribute(xf, "fontId")
                        .and_then(|index| fonts.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let fill_color = numeric_attribute(xf, "fillId")
                        .and_then(|index| fills.get(index))
                        .cloned()
                        .flatten();
                    let border = numeric_attribute(xf, "borderId")
                        .and_then(|index| borders.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let number_format = numeric_attribute(xf, "numFmtId")
                        .and_then(|id| {
                            custom_formats
                                .get(&(id as u32))
                                .cloned()
                                .or_else(|| builtin_number_format(id as u32).map(ToOwned::to_owned))
                        });
                    let alignment = xf.children().find(|child| child.has_tag_name("alignment"));
                    CellStyle {
                        font_family: font.family,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        underline: font.underline,
                        strikethrough: font.strikethrough,
                        wrap_text: alignment
                            .and_then(|node| node.attribute("wrapText"))
                            .is_some_and(|value| value == "1" || value == "true"),
                        font_color: font.color,
                        fill_color,
                        horizontal_alignment: alignment
                            .and_then(|node| node.attribute("horizontal"))
                            .map(ToOwned::to_owned),
                        vertical_alignment: alignment
                            .and_then(|node| node.attribute("vertical"))
                            .map(ToOwned::to_owned),
                        indent: alignment
                            .and_then(|node| node.attribute("indent"))
                            .and_then(|value| value.parse::<u32>().ok())
                            .filter(|steps| *steps > 0),
                        number_format,
                        border_top: border.top,
                        border_bottom: border.bottom,
                        border_left: border.left,
                        border_right: border.right,
                        border_diagonal: border.diagonal,
                        diagonal_up: border.diagonal_up,
                        diagonal_down: border.diagonal_down,
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let dxfs = document
        .descendants()
        .find(|node| node.has_tag_name("dxfs"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("dxf"))
                .map(|dxf| parse_dxf(dxf, colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let styles = if styles.is_empty() {
        vec![CellStyle::default()]
    } else {
        styles
    };
    Ok((styles, dxfs))
}

/// Differential (dxf) styles referenced by conditional-formatting rules.
/// Solid dxf fills carry the color in bgColor, unlike cell fills.
fn parse_dxf(dxf: Node<'_, '_>, colors: &ColorContext) -> CellStyle {
    let font = dxf
        .children()
        .find(|node| node.has_tag_name("font"))
        .map(|node| parse_font(node, colors))
        .unwrap_or_default();
    let fill_color = dxf
        .children()
        .find(|node| node.has_tag_name("fill"))
        .and_then(|fill| {
            let pattern = fill
                .children()
                .find(|node| node.has_tag_name("patternFill"))?;
            pattern
                .children()
                .find(|node| node.has_tag_name("bgColor"))
                .or_else(|| {
                    pattern
                        .children()
                        .find(|node| node.has_tag_name("fgColor"))
                })
                .and_then(|node| parse_color(node, colors))
        });
    let border = dxf
        .children()
        .find(|node| node.has_tag_name("border"))
        .map(|node| parse_border(node, colors))
        .unwrap_or_default();
    CellStyle {
        font_family: font.family,
        font_size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        wrap_text: false,
        font_color: font.color,
        fill_color,
        horizontal_alignment: None,
        vertical_alignment: None,
        indent: None,
        number_format: None,
        border_top: border.top,
        border_bottom: border.bottom,
        border_left: border.left,
        border_right: border.right,
        border_diagonal: border.diagonal,
        diagonal_up: border.diagonal_up,
        diagonal_down: border.diagonal_down,
    }
}

pub fn read_visual_objects(
    archive: &mut ZipArchive<File>,
    sheets: &[SheetVisualSource],
    colors: &ColorContext,
) -> Result<Vec<VisualObject>, SidecarError> {
    let mut visuals = Vec::new();
    for sheet in sheets {
        let sheet_relationships = read_relationships(archive, &sheet.worksheet_path)?;
        let Some(drawing_relationship) = sheet_relationships
            .values()
            .find(|relationship| relationship.relationship_type.ends_with("/drawing"))
        else {
            continue;
        };
        let drawing_path = resolve_part_target(
            &sheet.worksheet_path,
            &drawing_relationship.target,
        )?;
        visuals.extend(read_drawing(
            archive,
            &drawing_path,
            &sheet.sheet_id,
            visuals.len(),
            colors,
        )?);
    }
    Ok(visuals)
}

/// DrawingML solid fill: srgbClr, or schemeClr resolved via the theme palette.
fn drawing_fill_color(node: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    let fill = node
        .descendants()
        .find(|child| child.has_tag_name("solidFill"))
        .or_else(|| {
            // Gradient fills approximate to their first stop color.
            node.descendants()
                .find(|child| child.has_tag_name("gradFill"))
                .and_then(|grad| grad.descendants().find(|child| child.has_tag_name("gs")))
        })?;
    if let Some(srgb) = fill.descendants().find(|child| child.has_tag_name("srgbClr")) {
        return srgb.attribute("val").map(|value| format!("#{value}"));
    }
    let scheme = fill.descendants().find(|child| child.has_tag_name("schemeClr"))?;
    let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
    Some(tint_to_hex(base, 0.0))
}

fn scheme_color_rgb(name: &str, colors: &ColorContext) -> Option<(u8, u8, u8)> {
    if let Some(rest) = name.strip_prefix("accent") {
        return theme_accent(colors, rest.parse::<usize>().ok()?);
    }
    let index = match name {
        "lt1" | "bg1" => 0,
        "dk1" | "tx1" => 1,
        "lt2" | "bg2" => 2,
        "dk2" | "tx2" => 3,
        _ => return None,
    };
    colors.theme.get(index).copied()
}

pub fn read_media(
    archive: &mut ZipArchive<File>,
    media_path: &str,
) -> Result<MediaResult, SidecarError> {
    let mut entry = archive.by_name(media_path)?;
    if entry.size() > MAX_MEDIA_BYTES {
        return Err(SidecarError::Workbook(
            "Embedded image exceeds the media response limit.".into(),
        ));
    }
    let media_type = media_type_for_path(media_path)
        .ok_or_else(|| SidecarError::Workbook("Unsupported embedded image type.".into()))?;
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes)?;
    Ok(MediaResult {
        media_type: media_type.to_owned(),
        base64: base64::engine::general_purpose::STANDARD.encode(bytes),
    })
}

fn read_drawing(
    archive: &mut ZipArchive<File>,
    drawing_path: &str,
    sheet_id: &str,
    id_offset: usize,
    colors: &ColorContext,
) -> Result<Vec<VisualObject>, SidecarError> {
    let xml = read_xml(archive, drawing_path)?;
    let document = parse_document(&xml, drawing_path)?;
    let relationships = read_relationships(archive, drawing_path)?;
    let mut visuals = Vec::new();
    for (index, anchor_node) in document
        .descendants()
        .filter(|node| {
            node.has_tag_name("twoCellAnchor")
                || node.has_tag_name("oneCellAnchor")
                || node.has_tag_name("absoluteAnchor")
        })
        .enumerate()
    {
        let Some(anchor) = parse_anchor(anchor_node) else {
            continue;
        };
        let visual_id = format!("visual-{}", id_offset + index + 1);
        if let Some(chart_node) = anchor_node.descendants().find(|node| node.has_tag_name("chart")) {
            let Some(id) = relationship_id(chart_node) else {
                continue;
            };
            let Some(relationship) = relationships.get(&id) else {
                continue;
            };
            let chart_path = resolve_part_target(drawing_path, &relationship.target)?;
            visuals.push(VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "chart".into(),
                anchor,
                chart: Some(read_chart(archive, &chart_path, colors)?),
                chart_path: Some(chart_path.clone()),
                media_path: None,
                media_type: None,
                name: drawing_name(anchor_node),
                shape_type: None,
                fill_color: None,
                text: None,
                rotation: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        if let Some(blip_node) = anchor_node.descendants().find(|node| node.has_tag_name("blip")) {
            let Some(id) = relationship_id(blip_node) else {
                continue;
            };
            let Some(relationship) = relationships.get(&id) else {
                continue;
            };
            let media_path = resolve_part_target(drawing_path, &relationship.target)?;
            visuals.push(VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "image".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_type: media_type_for_path(&media_path).map(ToOwned::to_owned),
                media_path: Some(media_path),
                name: drawing_name(anchor_node),
                shape_type: None,
                fill_color: None,
                text: None,
                rotation: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        if let Some(shape_node) = anchor_node.descendants().find(|node| node.has_tag_name("sp")) {
            let shape_type = shape_node
                .descendants()
                .find(|node| node.has_tag_name("prstGeom"))
                .and_then(|node| node.attribute("prst"))
                .map(ToOwned::to_owned);
            let fill_color = shape_node
                .children()
                .find(|node| node.has_tag_name("spPr"))
                .and_then(|sppr| drawing_fill_color(sppr, colors));
            let rotation = shape_node
                .descendants()
                .find(|node| node.has_tag_name("xfrm"))
                .and_then(|node| node.attribute("rot"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 60_000.0);
            let text = shape_node
                .descendants()
                .find(|node| node.has_tag_name("txBody"))
                .map(|body| {
                    body.descendants()
                        .filter(|node| node.has_tag_name("t"))
                        .filter_map(|node| node.text())
                        .collect::<String>()
                })
                .filter(|value| !value.is_empty());
            visuals.push(VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "shape".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_path: None,
                media_type: None,
                name: drawing_name(anchor_node),
                shape_type,
                fill_color,
                text,
                rotation,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
        }
    }
    Ok(visuals)
}

const CHART_TYPE_NAMES: [&str; 7] = [
    "barChart",
    "lineChart",
    "pieChart",
    "doughnutChart",
    "areaChart",
    "scatterChart",
    "radarChart",
];

fn read_chart(
    archive: &mut ZipArchive<File>,
    chart_path: &str,
    colors: &ColorContext,
) -> Result<ChartMetadata, SidecarError> {
    let xml = read_xml(archive, chart_path)?;
    let document = parse_document(&xml, chart_path)?;
    Ok(chart_metadata(&document, colors))
}

fn chart_metadata(document: &Document<'_>, colors: &ColorContext) -> ChartMetadata {
    let chart_types = CHART_TYPE_NAMES
        .iter()
        .filter(|name| document.descendants().any(|node| node.has_tag_name(**name)))
        .map(|name| (*name).to_owned())
        .collect::<Vec<_>>();
    // Only the chart-level title — axes carry their own c:title deeper down.
    let title = document
        .descendants()
        .find(|node| node.has_tag_name("chart"))
        .and_then(|chart| direct_child(chart, "title"))
        .map(|node| {
            let rich = node
                .descendants()
                .filter(|child| child.has_tag_name("t"))
                .filter_map(|child| child.text())
                .collect::<String>();
            if !rich.is_empty() {
                return rich;
            }
            // Cell-linked title (<c:tx><c:strRef>): the strCache <c:v> holds
            // the cached cell text — show that instead of a placeholder (#181)
            node.descendants()
                .filter(|child| child.has_tag_name("v"))
                .filter_map(|child| child.text())
                .collect::<String>()
        })
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "Chart".into());
    let bar_direction = document
        .descendants()
        .find(|node| node.has_tag_name("barDir"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let series = document
        .descendants()
        .filter(|node| node.has_tag_name("ser"))
        .enumerate()
        .map(|(index, node)| parse_chart_series(node, index, colors))
        .collect::<Vec<_>>();
    ChartMetadata {
        chart_types,
        bar_direction,
        title,
        legend: legend_position(document),
        data_labels: data_labels(document),
        data_label_position: data_label_position(document),
        data_label_format: data_label_format(document),
        axis_titles: axis_titles(document),
        grouping: plot_grouping(document),
        gridlines: value_axis(document).map(|axis| direct_child(axis, "majorGridlines").is_some()),
        value_axis: value_axis_bounds(document),
        category_axis_format: category_axis_format(document),
        gap_width_pct: plot_val_attribute(document, "barChart", "gapWidth"),
        hole_size_pct: plot_val_attribute(document, "doughnutChart", "holeSize"),
        series,
    }
}

/// Scatter plots carry two valAx (X on the bottom, Y on the left); the left
/// one is the value axis the metadata (gridlines/bounds) should describe.
fn value_axis<'a>(document: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    let axes: Vec<_> = document
        .descendants()
        .filter(|node| node.has_tag_name("valAx"))
        .collect();
    axes.iter()
        .find(|axis| {
            direct_child(**axis, "axPos").and_then(|node| node.attribute("val")) == Some("l")
        })
        .or_else(|| axes.first())
        .copied()
}

fn category_axis_format(document: &Document<'_>) -> Option<String> {
    let axis = document
        .descendants()
        .find(|node| node.has_tag_name("catAx") || node.has_tag_name("dateAx"))?;
    direct_child(axis, "numFmt")?
        .attribute("formatCode")
        .map(ToOwned::to_owned)
}

fn value_axis_bounds(document: &Document<'_>) -> Option<ValueAxisBounds> {
    let scaling = direct_child(value_axis(document)?, "scaling")?;
    let bound = |name: &str| {
        direct_child(scaling, name)
            .and_then(|node| node.attribute("val"))
            .and_then(|value| value.parse::<f64>().ok())
    };
    let min = bound("min");
    let max = bound("max");
    (min.is_some() || max.is_some()).then_some(ValueAxisBounds { min, max })
}

fn plot_val_attribute(document: &Document<'_>, plot: &str, name: &str) -> Option<u32> {
    let plot = document.descendants().find(|node| node.has_tag_name(plot))?;
    direct_child(plot, name)
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok())
}

fn legend_position(document: &Document<'_>) -> String {
    let Some(legend) = document.descendants().find(|node| node.has_tag_name("legend")) else {
        return "none".into();
    };
    match direct_child(legend, "legendPos").and_then(|node| node.attribute("val")) {
        Some("b") => "bottom",
        Some("t") => "top",
        Some("l") => "left",
        // "r", "tr", or absent all render on the right, the OOXML default.
        _ => "right",
    }
    .into()
}

/// Plot-level dLbls, falling back to the first series' dLbls.
fn data_labels_node<'a>(document: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    document
        .descendants()
        .find(|node| CHART_TYPE_NAMES.iter().any(|name| node.has_tag_name(*name)))
        .and_then(|plot| direct_child(plot, "dLbls"))
        .or_else(|| {
            document
                .descendants()
                .find(|node| node.has_tag_name("ser"))
                .and_then(|series| direct_child(series, "dLbls"))
        })
}

fn data_labels(document: &Document<'_>) -> String {
    let Some(labels) = data_labels_node(document) else {
        return "none".into();
    };
    let shown = |name: &str| {
        direct_child(labels, name)
            .and_then(|node| node.attribute("val"))
            .is_some_and(|value| value == "1" || value == "true")
    };
    if shown("delete") {
        return "none".into();
    }
    if shown("showPercent") {
        return if shown("showCatName") {
            "category-percent"
        } else {
            "percent"
        }
        .into();
    }
    if shown("showVal") {
        return "value".into();
    }
    "none".into()
}

fn data_label_position(document: &Document<'_>) -> Option<String> {
    let position = direct_child(data_labels_node(document)?, "dLblPos")?.attribute("val")?;
    match position {
        "ctr" => Some("center".into()),
        "inEnd" => Some("inside-end".into()),
        "outEnd" => Some("outside-end".into()),
        _ => None,
    }
}

fn data_label_format(document: &Document<'_>) -> Option<String> {
    direct_child(data_labels_node(document)?, "numFmt")?
        .attribute("formatCode")
        .map(ToOwned::to_owned)
}

fn axis_titles(document: &Document<'_>) -> Option<AxisTitles> {
    let title_of = |names: &[&str]| -> Option<String> {
        let axis = names
            .iter()
            .find_map(|name| document.descendants().find(|node| node.has_tag_name(*name)))?;
        let text = direct_child(axis, "title")?
            .descendants()
            .filter(|node| node.has_tag_name("t"))
            .filter_map(|node| node.text())
            .collect::<String>();
        (!text.is_empty()).then_some(text)
    };
    let category = title_of(&["catAx", "dateAx"]);
    let value = title_of(&["valAx"]);
    if category.is_none() && value.is_none() {
        return None;
    }
    Some(AxisTitles { category, value })
}

fn plot_grouping(document: &Document<'_>) -> Option<String> {
    let plot = document.descendants().find(|node| {
        ["barChart", "areaChart", "lineChart"]
            .iter()
            .any(|name| node.has_tag_name(*name))
    })?;
    direct_child(plot, "grouping")
        .and_then(|node| node.attribute("val"))
        .filter(|value| {
            matches!(*value, "clustered" | "stacked" | "percentStacked" | "standard")
        })
        .map(ToOwned::to_owned)
}

fn parse_chart_series(
    series: Node<'_, '_>,
    index: usize,
    colors: &ColorContext,
) -> ChartSeries {
    let name = direct_child(series, "tx")
        .and_then(|node| first_cached_value(node))
        .unwrap_or_else(|| "Series".into());
    // Explicit series fill/line color, else the theme accent cycle Excel uses
    // for automatic chart colors.
    let color = direct_child(series, "spPr")
        .and_then(|sppr| drawing_fill_color(sppr, colors))
        .or_else(|| {
            theme_accent(colors, index % 6 + 1).map(|base| tint_to_hex(base, 0.0))
        });
    let category_node = direct_child(series, "cat").or_else(|| direct_child(series, "xVal"));
    let categories = category_node.map(cached_values).unwrap_or_default();
    let category_format = category_node.and_then(cache_format_code);
    let value_node = direct_child(series, "val").or_else(|| direct_child(series, "yVal"));
    let values = value_node
        .map(cached_values)
        .unwrap_or_default()
        .into_iter()
        .filter_map(|value| value.parse::<f64>().ok())
        .collect();
    let number_format = value_node.and_then(cache_format_code);
    let trendline = series
        .descendants()
        .find(|node| node.has_tag_name("trendlineType"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let values_ref = value_node.and_then(formula_ref);
    let categories_ref = category_node.and_then(formula_ref);
    let explosion_pct = direct_child(series, "explosion")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok());
    let point_colors = data_points(series)
        .filter_map(|(index, point)| {
            Some(PointColor {
                index,
                color: direct_child(point, "spPr")
                    .and_then(|sppr| drawing_fill_color(sppr, colors))?,
            })
        })
        .collect::<Vec<_>>();
    let point_explosions = data_points(series)
        .filter_map(|(index, point)| {
            Some(PointExplosion {
                index,
                pct: direct_child(point, "explosion")?
                    .attribute("val")?
                    .parse::<u32>()
                    .ok()?,
            })
        })
        .collect::<Vec<_>>();
    ChartSeries {
        name,
        categories,
        values,
        number_format,
        category_format,
        color,
        trendline,
        values_ref,
        categories_ref,
        point_colors: (!point_colors.is_empty()).then_some(point_colors),
        explosion_pct,
        point_explosions: (!point_explosions.is_empty()).then_some(point_explosions),
    }
}

fn cache_format_code(node: Node<'_, '_>) -> Option<String> {
    node.descendants()
        .find(|child| child.has_tag_name("formatCode"))
        .and_then(|child| child.text())
        .map(ToOwned::to_owned)
}

/// `c:dPt` entries paired with their `c:idx` value.
fn data_points<'a>(series: Node<'a, 'a>) -> impl Iterator<Item = (u32, Node<'a, 'a>)> {
    series
        .children()
        .filter(|node| node.has_tag_name("dPt"))
        .filter_map(|point| {
            let index = direct_child(point, "idx")?
                .attribute("val")?
                .parse::<u32>()
                .ok()?;
            Some((index, point))
        })
}

fn formula_ref(node: Node<'_, '_>) -> Option<String> {
    node.descendants()
        .find(|child| child.has_tag_name("f"))
        .and_then(|child| child.text())
        .map(ToOwned::to_owned)
        .filter(|value| !value.is_empty())
}

fn parse_anchor(anchor: Node<'_, '_>) -> Option<DrawingAnchor> {
    let from = direct_child(anchor, "from")?;
    let to = direct_child(anchor, "to").unwrap_or(from);
    Some(DrawingAnchor {
        from_row: marker_value(from, "row")?,
        from_column: marker_value(from, "col")?,
        from_row_offset: marker_signed_value(from, "rowOff").unwrap_or(0),
        from_column_offset: marker_signed_value(from, "colOff").unwrap_or(0),
        to_row: marker_value(to, "row").unwrap_or_else(|| marker_value(from, "row").unwrap_or(0) + 20),
        to_column: marker_value(to, "col")
            .unwrap_or_else(|| marker_value(from, "col").unwrap_or(0) + 8),
        to_row_offset: marker_signed_value(to, "rowOff").unwrap_or(0),
        to_column_offset: marker_signed_value(to, "colOff").unwrap_or(0),
    })
}

fn parse_font(font: Node<'_, '_>, colors: &ColorContext) -> FontStyle {
    FontStyle {
        family: font
            .children()
            .find(|node| node.has_tag_name("name"))
            .and_then(|node| node.attribute("val"))
            .map(ToOwned::to_owned),
        size: font
            .children()
            .find(|node| node.has_tag_name("sz"))
            .and_then(|node| node.attribute("val"))
            .and_then(|value| value.parse::<f64>().ok()),
        bold: font.children().any(|node| node.has_tag_name("b")),
        italic: font.children().any(|node| node.has_tag_name("i")),
        underline: font
            .children()
            .find(|node| node.has_tag_name("u"))
            .is_some_and(|node| node.attribute("val") != Some("none")),
        strikethrough: font
            .children()
            .find(|node| node.has_tag_name("strike"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        color: font
            .children()
            .find(|node| node.has_tag_name("color"))
            .and_then(|node| parse_color(node, colors)),
    }
}

fn parse_fill(fill: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    let pattern = fill
        .children()
        .find(|node| node.has_tag_name("patternFill"))?;
    let pattern_type = pattern.attribute("patternType");
    if pattern_type == Some("none") {
        return None;
    }
    let foreground = pattern
        .children()
        .find(|node| node.has_tag_name("fgColor"))
        .and_then(|node| parse_color(node, colors));
    let background = pattern
        .children()
        .find(|node| node.has_tag_name("bgColor"))
        .and_then(|node| parse_color(node, colors));
    // Textured patterns (gray125, stripes, …) render as the per-channel blend
    // of both colors — the closest flat-color approximation of the texture.
    if pattern_type.is_some_and(|value| value != "solid") {
        if let (Some(fg), Some(bg)) = (foreground.as_deref(), background.as_deref()) {
            if let Some(mixed) = mix_hex(fg, bg) {
                return Some(mixed);
            }
        }
    }
    foreground.or(background)
}

fn mix_hex(first: &str, second: &str) -> Option<String> {
    let parse = |hex: &str| -> Option<(u8, u8, u8)> {
        let value = hex.strip_prefix('#')?;
        Some((
            u8::from_str_radix(value.get(0..2)?, 16).ok()?,
            u8::from_str_radix(value.get(2..4)?, 16).ok()?,
            u8::from_str_radix(value.get(4..6)?, 16).ok()?,
        ))
    };
    let (r1, g1, b1) = parse(first)?;
    let (r2, g2, b2) = parse(second)?;
    Some(format!(
        "#{:02X}{:02X}{:02X}",
        (u16::from(r1) + u16::from(r2)) / 2,
        (u16::from(g1) + u16::from(g2)) / 2,
        (u16::from(b1) + u16::from(b2)) / 2,
    ))
}

fn parse_border(border: Node<'_, '_>, colors: &ColorContext) -> BorderSet {
    let edge = |name: &str| -> Option<BorderEdge> {
        let node = border.children().find(|child| child.has_tag_name(name))?;
        let style = node.attribute("style")?;
        if style == "none" {
            return None;
        }
        Some(BorderEdge {
            style: style.to_owned(),
            color: node
                .children()
                .find(|child| child.has_tag_name("color"))
                .and_then(|child| parse_color(child, colors)),
        })
    };
    BorderSet {
        top: edge("top"),
        bottom: edge("bottom"),
        left: edge("left"),
        right: edge("right"),
        diagonal: edge("diagonal"),
        diagonal_up: border
            .attribute("diagonalUp")
            .is_some_and(|value| value == "1" || value == "true"),
        diagonal_down: border
            .attribute("diagonalDown")
            .is_some_and(|value| value == "1" || value == "true"),
    }
}

/// Legacy indexed palette, ECMA-376 §18.8.27. Indexes 64/65 are the system
/// window text/background colors.
const INDEXED_COLORS: [&str; 66] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    "000000", "FFFFFF",
];

fn parse_color(node: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    resolve_color(
        node.attribute("rgb"),
        node.attribute("indexed"),
        node.attribute("theme"),
        node.attribute("tint"),
        colors,
    )
}

pub fn resolve_color(
    rgb: Option<&str>,
    indexed: Option<&str>,
    theme: Option<&str>,
    tint: Option<&str>,
    colors: &ColorContext,
) -> Option<String> {
    if let Some(rgb) = rgb {
        let value = if rgb.len() == 8 { &rgb[2..] } else { rgb };
        return Some(format!("#{value}"));
    }
    if let Some(indexed) = indexed {
        let index = indexed.parse::<usize>().ok()?;
        return INDEXED_COLORS.get(index).map(|value| format!("#{value}"));
    }
    let theme = theme?.parse::<usize>().ok()?;
    let base = *colors.theme.get(theme)?;
    let tint = tint
        .and_then(|value| value.parse::<f64>().ok())
        .unwrap_or(0.0);
    let (red, green, blue) = apply_tint(base, tint);
    Some(format!("#{red:02X}{green:02X}{blue:02X}"))
}

/// Theme accent color (1-6) as rgb, if the palette was loaded.
pub fn theme_accent(colors: &ColorContext, accent: usize) -> Option<(u8, u8, u8)> {
    // Effective palette order: [lt1, dk1, lt2, dk2, accent1-6, ...]
    colors.theme.get(3 + accent).copied()
}

pub fn tint_to_hex(base: (u8, u8, u8), tint: f64) -> String {
    let (red, green, blue) = apply_tint(base, tint);
    format!("#{red:02X}{green:02X}{blue:02X}")
}

pub fn read_theme_palette(archive: &mut ZipArchive<File>) -> Result<ColorContext, SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/theme/theme1.xml")? else {
        return Ok(ColorContext::default());
    };
    let document = parse_document(&xml, "theme1.xml")?;
    let Some(scheme) = document
        .descendants()
        .find(|node| node.has_tag_name("clrScheme"))
    else {
        return Ok(ColorContext::default());
    };
    let slot = |name: &str| -> Option<(u8, u8, u8)> {
        let node = scheme.children().find(|child| child.has_tag_name(name))?;
        let hex = node
            .children()
            .find(|child| child.has_tag_name("srgbClr"))
            .and_then(|child| child.attribute("val"))
            .or_else(|| {
                node.children()
                    .find(|child| child.has_tag_name("sysClr"))
                    .and_then(|child| child.attribute("lastClr"))
            })?;
        parse_hex_rgb(hex)
    };
    // The `theme` attribute indexes [lt1, dk1, lt2, dk2, accent1-6, hlink,
    // folHlink] — light/dark pairs swapped versus clrScheme document order.
    let order = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    let mut theme = Vec::with_capacity(order.len());
    for name in order {
        match slot(name) {
            Some(color) => theme.push(color),
            None => return Ok(ColorContext::default()),
        }
    }
    Ok(ColorContext { theme })
}

fn parse_hex_rgb(hex: &str) -> Option<(u8, u8, u8)> {
    let value = if hex.len() == 8 { &hex[2..] } else { hex };
    if value.len() != 6 {
        return None;
    }
    Some((
        u8::from_str_radix(&value[0..2], 16).ok()?,
        u8::from_str_radix(&value[2..4], 16).ok()?,
        u8::from_str_radix(&value[4..6], 16).ok()?,
    ))
}

/// Excel's tint transform: scale HSL luminance toward black (tint < 0) or
/// white (tint > 0).
fn apply_tint(rgb: (u8, u8, u8), tint: f64) -> (u8, u8, u8) {
    if tint == 0.0 {
        return rgb;
    }
    let (hue, saturation, luminance) = rgb_to_hsl(rgb);
    let luminance = if tint < 0.0 {
        luminance * (1.0 + tint)
    } else {
        luminance * (1.0 - tint) + tint
    };
    hsl_to_rgb(hue, saturation, luminance.clamp(0.0, 1.0))
}

fn rgb_to_hsl((red, green, blue): (u8, u8, u8)) -> (f64, f64, f64) {
    let red = f64::from(red) / 255.0;
    let green = f64::from(green) / 255.0;
    let blue = f64::from(blue) / 255.0;
    let maximum = red.max(green).max(blue);
    let minimum = red.min(green).min(blue);
    let luminance = (maximum + minimum) / 2.0;
    if maximum == minimum {
        return (0.0, 0.0, luminance);
    }
    let delta = maximum - minimum;
    let saturation = if luminance > 0.5 {
        delta / (2.0 - maximum - minimum)
    } else {
        delta / (maximum + minimum)
    };
    let hue = if maximum == red {
        (green - blue) / delta + if green < blue { 6.0 } else { 0.0 }
    } else if maximum == green {
        (blue - red) / delta + 2.0
    } else {
        (red - green) / delta + 4.0
    } / 6.0;
    (hue, saturation, luminance)
}

fn hsl_to_rgb(hue: f64, saturation: f64, luminance: f64) -> (u8, u8, u8) {
    if saturation == 0.0 {
        let value = (luminance * 255.0).round() as u8;
        return (value, value, value);
    }
    let q = if luminance < 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let p = 2.0 * luminance - q;
    let channel = |mut t: f64| -> u8 {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        let value = if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 1.0 / 2.0 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        };
        (value * 255.0).round() as u8
    };
    (
        channel(hue + 1.0 / 3.0),
        channel(hue),
        channel(hue - 1.0 / 3.0),
    )
}

/// Cell comments (legacy notes) attached to a worksheet, as
/// (cell reference, author, text) tuples.
/// PivotTable output areas on a worksheet (from each pivot part's
/// `<location ref>`). The viewer must protect these cells: editing baked
/// pivot output corrupts the file's pivot semantics.
pub struct PivotPartInfo {
    pub path: String,
    pub cache_path: Option<String>,
    pub output_ref: String,
}

pub fn read_pivot_tables(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<PivotPartInfo>, SidecarError> {
    let relationships = read_relationships(archive, worksheet_path)?;
    let mut infos = Vec::new();
    for relationship in relationships.values() {
        if !relationship.relationship_type.ends_with("/pivotTable") {
            continue;
        }
        let pivot_path = resolve_part_target(worksheet_path, &relationship.target)?;
        let Some(xml) = read_optional_xml(archive, &pivot_path)? else {
            continue;
        };
        let document = parse_document(&xml, &pivot_path)?;
        let Some(output_ref) = document
            .descendants()
            .find(|node| node.has_tag_name("location"))
            .and_then(|node| node.attribute("ref"))
        else {
            continue;
        };
        let cache_path = read_relationships(archive, &pivot_path)?
            .values()
            .find(|part| part.relationship_type.ends_with("/pivotCacheDefinition"))
            .map(|part| resolve_part_target(&pivot_path, &part.target))
            .transpose()?;
        infos.push(PivotPartInfo {
            path: pivot_path,
            cache_path,
            output_ref: output_ref.to_owned(),
        });
    }
    Ok(infos)
}

pub fn read_comments(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<(String, String, String)>, SidecarError> {
    let relationships = read_relationships(archive, worksheet_path)?;
    let Some(comments_relationship) = relationships
        .values()
        .find(|relationship| relationship.relationship_type.ends_with("/comments"))
    else {
        return Ok(Vec::new());
    };
    let comments_path = resolve_part_target(worksheet_path, &comments_relationship.target)?;
    let Some(xml) = read_optional_xml(archive, &comments_path)? else {
        return Ok(Vec::new());
    };
    let document = parse_document(&xml, &comments_path)?;
    let authors = document
        .descendants()
        .find(|node| node.has_tag_name("authors"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("author"))
                .map(|child| child.text().unwrap_or_default().to_owned())
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Ok(document
        .descendants()
        .filter(|node| node.has_tag_name("comment"))
        .filter_map(|comment| {
            let reference = comment.attribute("ref")?.to_owned();
            let author = comment
                .attribute("authorId")
                .and_then(|id| id.parse::<usize>().ok())
                .and_then(|id| authors.get(id))
                .cloned()
                .unwrap_or_default();
            let text = comment
                .descendants()
                .filter(|node| node.has_tag_name("t"))
                .filter_map(|node| node.text())
                .collect::<String>();
            Some((reference, author, text))
        })
        .collect())
}

/// Package paths of the table parts attached to a worksheet.
pub fn table_part_paths(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<String>, SidecarError> {
    let relationships = read_relationships(archive, worksheet_path)?;
    let mut paths = Vec::new();
    for relationship in relationships.values() {
        if relationship.relationship_type.ends_with("/table") {
            paths.push(resolve_part_target(worksheet_path, &relationship.target)?);
        }
    }
    Ok(paths)
}

/// Relationship id → hyperlink target for a worksheet part. Internal
/// (location-only) links carry no relationship and are not included here.
pub fn hyperlink_targets(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<HashMap<String, String>, SidecarError> {
    Ok(read_relationships(archive, worksheet_path)?
        .into_iter()
        .filter(|(_, relationship)| relationship.relationship_type.ends_with("/hyperlink"))
        .map(|(id, relationship)| (id, relationship.target))
        .collect())
}

fn read_relationships(
    archive: &mut ZipArchive<File>,
    source_path: &str,
) -> Result<HashMap<String, Relationship>, SidecarError> {
    let source = Path::new(source_path);
    let file_name = source
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| SidecarError::Workbook("Relationship source path is invalid.".into()))?;
    let parent = source.parent().unwrap_or_else(|| Path::new(""));
    let relationship_path = parent
        .join("_rels")
        .join(format!("{file_name}.rels"))
        .to_string_lossy()
        .replace('\\', "/");
    let Some(xml) = read_optional_xml(archive, &relationship_path)? else {
        return Ok(HashMap::new());
    };
    let document = parse_document(&xml, &relationship_path)?;
    Ok(document
        .descendants()
        .filter(|node| node.has_tag_name("Relationship"))
        .filter_map(|node| {
            Some((
                node.attribute("Id")?.to_owned(),
                Relationship {
                    target: node.attribute("Target")?.to_owned(),
                    relationship_type: node.attribute("Type").unwrap_or_default().to_owned(),
                },
            ))
        })
        .collect())
}

fn resolve_part_target(source_path: &str, target: &str) -> Result<String, SidecarError> {
    let candidate = if target.starts_with('/') {
        PathBuf::from(target.trim_start_matches('/'))
    } else {
        Path::new(source_path)
            .parent()
            .unwrap_or_else(|| Path::new(""))
            .join(target)
    };
    let mut normalized = PathBuf::new();
    for component in candidate.components() {
        match component {
            Component::Normal(value) => normalized.push(value),
            Component::ParentDir => {
                if !normalized.pop() {
                    return Err(SidecarError::Workbook(
                        "OOXML relationship escapes the package.".into(),
                    ));
                }
            }
            Component::CurDir => {}
            _ => {
                return Err(SidecarError::Workbook(
                    "OOXML relationship has an unsafe path.".into(),
                ));
            }
        }
    }
    normalized
        .to_str()
        .map(|value| value.replace('\\', "/"))
        .ok_or_else(|| SidecarError::Workbook("OOXML part path is invalid UTF-8.".into()))
}

fn read_xml(archive: &mut ZipArchive<File>, path: &str) -> Result<String, SidecarError> {
    read_optional_xml(archive, path)?
        .ok_or_else(|| SidecarError::Workbook(format!("Workbook is missing {path}.")))
}

fn read_optional_xml(
    archive: &mut ZipArchive<File>,
    path: &str,
) -> Result<Option<String>, SidecarError> {
    let Ok(mut entry) = archive.by_name(path) else {
        return Ok(None);
    };
    let mut xml = String::new();
    entry.read_to_string(&mut xml)?;
    Ok(Some(xml))
}

fn parse_document<'a>(xml: &'a str, path: &str) -> Result<Document<'a>, SidecarError> {
    Document::parse(xml)
        .map_err(|error| SidecarError::Workbook(format!("Invalid XML in {path}: {error}")))
}

fn direct_child<'a>(node: Node<'a, 'a>, name: &str) -> Option<Node<'a, 'a>> {
    node.children().find(|child| child.has_tag_name(name))
}

fn relationship_id(node: Node<'_, '_>) -> Option<String> {
    node.attributes()
        .find(|attribute| attribute.name() == "id" || attribute.name() == "embed")
        .map(|attribute| attribute.value().to_owned())
}

fn drawing_name(anchor: Node<'_, '_>) -> Option<String> {
    anchor
        .descendants()
        .find(|node| node.has_tag_name("cNvPr"))
        .and_then(|node| node.attribute("name"))
        .map(ToOwned::to_owned)
}

fn cached_values(node: Node<'_, '_>) -> Vec<String> {
    node.descendants()
        .filter(|child| child.has_tag_name("pt"))
        .filter_map(|point| {
            point
                .children()
                .find(|child| child.has_tag_name("v"))
                .and_then(|value| value.text())
                .map(ToOwned::to_owned)
        })
        .collect()
}

fn first_cached_value(node: Node<'_, '_>) -> Option<String> {
    cached_values(node).into_iter().next().or_else(|| {
        node.descendants()
            .find(|child| child.has_tag_name("v"))
            .and_then(|child| child.text())
            .map(ToOwned::to_owned)
    })
}

fn marker_value(marker: Node<'_, '_>, name: &str) -> Option<usize> {
    marker
        .children()
        .find(|child| child.has_tag_name(name))
        .and_then(|child| child.text())
        .and_then(|value| value.parse::<usize>().ok())
}

fn marker_signed_value(marker: Node<'_, '_>, name: &str) -> Option<i64> {
    marker
        .children()
        .find(|child| child.has_tag_name(name))
        .and_then(|child| child.text())
        .and_then(|value| value.parse::<i64>().ok())
}

fn numeric_attribute(node: Node<'_, '_>, name: &str) -> Option<usize> {
    node.attribute(name)?.parse::<usize>().ok()
}

fn media_type_for_path(path: &str) -> Option<&'static str> {
    match Path::new(path)
        .extension()
        .and_then(|value| value.to_str())?
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => Some("image/png"),
        "jpg" | "jpeg" => Some("image/jpeg"),
        "gif" => Some("image/gif"),
        "bmp" => Some("image/bmp"),
        "svg" => Some("image/svg+xml"),
        _ => None,
    }
}

/// Implicit number formats, ECMA-376 §18.8.30. Ids 23-26 are undocumented
/// and stay unresolved; everything else is mapped so that a builtin id never
/// falls back to General (which would surface raw date serials, #numFmt58).
///
/// Locale-reserved ranges carry no formatCode in styles.xml — the reader is
/// expected to know them:
///  - 27-36 / 50-58: East-Asian date/time formats. The same id means a
///    different pattern per locale (zh/ja/ko/zh-TW) and the file does not
///    record which; we resolve with the zh-CN table since every locale's
///    variant is a date/time of the same shape, and zh Excel/WPS files are
///    the ones we actually receive. The zh AM/PM token (U+4E0A/U+4E0B
///    U+5348) is not understood by the renderer's numfmt, so 34/35/55/56
///    render as 24-hour. Escapes: U+5E74 year, U+6708 month, U+65E5 day,
///    U+65F6 hour, U+5206 minute, U+79D2 second.
///  - 41-44: accounting formats; 42/44 use "$" as the symbol is likewise
///    locale-defined and unrecorded.
///  - 59-81: th-TH; numfmt has no Thai digit/era tokens, so these map to
///    Arabic-digit equivalents (Buddhist-era years render as Gregorian).
fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        // ECMA-376 prints 14 as "mm-dd-yy", but Excel actually renders the
        // locale short date — m/d/yyyy under en-US — and users reconcile
        // against Excel, not the spec text (#184).
        14 => Some("m/d/yyyy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        27 | 36 | 50 | 52 | 57 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\""),
        28 | 29 | 51 | 53 | 54 | 58 => Some("m\"\u{6708}\"d\"\u{65e5}\""),
        30 => Some("m-d-yy"),
        31 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\""),
        32 | 34 | 55 => Some("h\"\u{65f6}\"mm\"\u{5206}\""),
        33 | 35 | 56 => Some("h\"\u{65f6}\"mm\"\u{5206}\"ss\"\u{79d2}\""),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        41 => Some(r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#),
        42 => Some(r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#),
        43 => Some(r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#),
        44 => Some(r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        59 => Some("0"),
        60 => Some("0.00"),
        61 => Some("#,##0"),
        62 => Some("#,##0.00"),
        67 => Some("0%"),
        68 => Some("0.00%"),
        69 => Some("# ?/?"),
        70 => Some("# ??/??"),
        71 => Some("d/m/yyyy"),
        72 => Some("d-mmm-yy"),
        73 => Some("d-mmm"),
        74 => Some("mmm-yy"),
        75 => Some("h:mm"),
        76 => Some("h:mm:ss"),
        77 => Some("d/m/yyyy h:mm"),
        78 => Some("mm:ss"),
        79 => Some("[h]:mm:ss"),
        80 => Some("mm:ss.0"),
        81 => Some("d/m/yy"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn metadata_with(body: &str, colors: &ColorContext) -> ChartMetadata {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>{body}</c:chart></c:chartSpace>"#
        );
        chart_metadata(&Document::parse(&xml).unwrap(), colors)
    }

    fn metadata(body: &str) -> ChartMetadata {
        metadata_with(body, &ColorContext::default())
    }

    fn theme_colors() -> ColorContext {
        ColorContext {
            theme: (0..12).map(|slot| (slot as u8, 0x22, 0x33)).collect(),
        }
    }

    #[test]
    fn maps_legend_positions_and_defaults() {
        for (val, expected) in [
            ("r", "right"),
            ("b", "bottom"),
            ("t", "top"),
            ("l", "left"),
            ("tr", "right"),
        ] {
            let body = format!(r#"<c:legend><c:legendPos val="{val}"/></c:legend>"#);
            assert_eq!(metadata(&body).legend, expected, "legendPos {val}");
        }
        assert_eq!(metadata("<c:legend/>").legend, "right");
        assert_eq!(metadata("<c:plotArea><c:barChart/></c:plotArea>").legend, "none");
    }

    #[test]
    fn maps_data_labels_from_plot_or_series() {
        let plot = |labels: &str| {
            format!("<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser>{labels}</c:pieChart></c:plotArea>")
        };
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showVal val=\"1\"/></c:dLbls>")).data_labels,
            "value"
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showPercent val=\"1\"/></c:dLbls>")).data_labels,
            "percent"
        );
        assert_eq!(
            metadata(&plot(
                "<c:dLbls><c:showCatName val=\"1\"/><c:showPercent val=\"1\"/></c:dLbls>"
            ))
            .data_labels,
            "category-percent"
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:delete val=\"1\"/></c:dLbls>")).data_labels,
            "none"
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showVal val=\"0\"/></c:dLbls>")).data_labels,
            "none"
        );
        assert_eq!(metadata(&plot("")).data_labels, "none");
        // Plot-level dLbls missing: fall back to the first series.
        let series_level = "<c:plotArea><c:barChart><c:ser><c:dLbls><c:showVal val=\"1\"/></c:dLbls></c:ser></c:barChart></c:plotArea>";
        assert_eq!(metadata(series_level).data_labels, "value");
    }

    #[test]
    fn maps_data_label_position_and_format() {
        let plot = |labels: &str| {
            format!("<c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/></c:ser>{labels}</c:barChart></c:plotArea>")
        };
        for (val, expected) in [("ctr", "center"), ("inEnd", "inside-end"), ("outEnd", "outside-end")] {
            let body = plot(&format!("<c:dLbls><c:dLblPos val=\"{val}\"/></c:dLbls>"));
            assert_eq!(metadata(&body).data_label_position.as_deref(), Some(expected), "dLblPos {val}");
        }
        let best_fit = plot("<c:dLbls><c:dLblPos val=\"bestFit\"/></c:dLbls>");
        assert!(metadata(&best_fit).data_label_position.is_none());
        assert!(metadata(&plot("<c:dLbls><c:showVal val=\"1\"/></c:dLbls>")).data_label_position.is_none());

        let series_level = "<c:plotArea><c:barChart><c:ser><c:dLbls><c:dLblPos val=\"outEnd\"/><c:numFmt formatCode=\"0.0%\"/></c:dLbls></c:ser></c:barChart></c:plotArea>";
        let chart = metadata(series_level);
        assert_eq!(chart.data_label_position.as_deref(), Some("outside-end"));
        assert_eq!(chart.data_label_format.as_deref(), Some("0.0%"));

        let formatted = plot("<c:dLbls><c:numFmt formatCode=\"#,##0\" sourceLinked=\"0\"/></c:dLbls>");
        assert_eq!(metadata(&formatted).data_label_format.as_deref(), Some("#,##0"));
        assert!(metadata(&plot("<c:dLbls/>")).data_label_format.is_none());
    }

    /// Issue #181: a cell-linked title (<c:tx><c:strRef>) shows the cached
    /// cell text from strCache instead of the "Chart" placeholder.
    #[test]
    fn reads_cell_linked_chart_titles_from_the_str_cache() {
        let linked = metadata(
            r#"<c:title><c:tx><c:strRef><c:f>Charts!$B$58</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales by Salesperson</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title><c:plotArea><c:lineChart/></c:plotArea>"#,
        );
        assert_eq!(linked.title, "Sales by Salesperson");
        // rich-text titles keep winning when both forms are present
        let rich = metadata(
            r#"<c:title><c:tx><c:rich><a:p><a:r><a:t>Static</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:lineChart/></c:plotArea>"#,
        );
        assert_eq!(rich.title, "Static");
    }

    #[test]
    fn collects_axis_titles() {
        let axis_title = |text: &str| {
            format!("<c:title><c:tx><c:rich><a:p><a:r><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx></c:title>")
        };
        let both = format!(
            "<c:plotArea><c:barChart/><c:catAx>{}</c:catAx><c:valAx>{}</c:valAx></c:plotArea>",
            axis_title("Month"),
            axis_title("Sales"),
        );
        let titles = metadata(&both).axis_titles.unwrap();
        assert_eq!(titles.category.as_deref(), Some("Month"));
        assert_eq!(titles.value.as_deref(), Some("Sales"));

        let date_axis = format!(
            "<c:plotArea><c:lineChart/><c:dateAx>{}</c:dateAx><c:valAx/></c:plotArea>",
            axis_title("Quarter"),
        );
        let titles = metadata(&date_axis).axis_titles.unwrap();
        assert_eq!(titles.category.as_deref(), Some("Quarter"));
        assert_eq!(titles.value, None);

        let untitled = "<c:plotArea><c:barChart/><c:catAx/><c:valAx/></c:plotArea>";
        assert!(metadata(untitled).axis_titles.is_none());
    }

    #[test]
    fn reads_grouping_from_first_grouped_plot() {
        for value in ["clustered", "stacked", "percentStacked", "standard"] {
            let body = format!(
                "<c:plotArea><c:barChart><c:grouping val=\"{value}\"/></c:barChart></c:plotArea>"
            );
            assert_eq!(metadata(&body).grouping.as_deref(), Some(value));
        }
        let unknown =
            "<c:plotArea><c:areaChart><c:grouping val=\"weird\"/></c:areaChart></c:plotArea>";
        assert!(metadata(unknown).grouping.is_none());
        assert!(metadata("<c:plotArea><c:pieChart/></c:plotArea>").grouping.is_none());
    }

    #[test]
    fn reads_point_colors_from_srgb_and_scheme_fills() {
        let body = r#"<c:plotArea><c:pieChart><c:ser>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr></c:dPt>
            <c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr></c:dPt>
            <c:dPt><c:idx val="3"/></c:dPt>
        </c:ser></c:pieChart></c:plotArea>"#;
        let chart = metadata_with(body, &theme_colors());
        let points = chart.series[0].point_colors.as_ref().unwrap();
        assert_eq!(points.len(), 2);
        assert_eq!((points[0].index, points[0].color.as_str()), (0, "#FF8800"));
        // accent2 lives at theme slot 5.
        assert_eq!((points[1].index, points[1].color.as_str()), (2, "#052233"));

        let plain = "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser></c:pieChart></c:plotArea>";
        assert!(metadata(plain).series[0].point_colors.is_none());
    }

    #[test]
    fn reads_gridlines_only_when_a_value_axis_exists() {
        let with = "<c:plotArea><c:barChart/><c:catAx/><c:valAx><c:majorGridlines/></c:valAx></c:plotArea>";
        assert_eq!(metadata(with).gridlines, Some(true));
        let without = "<c:plotArea><c:barChart/><c:catAx/><c:valAx/></c:plotArea>";
        assert_eq!(metadata(without).gridlines, Some(false));
        assert!(metadata("<c:plotArea><c:pieChart/></c:plotArea>").gridlines.is_none());
    }

    #[test]
    fn reads_value_axis_bounds() {
        let both = "<c:plotArea><c:barChart/><c:valAx><c:scaling><c:min val=\"-2.5\"/><c:max val=\"100\"/></c:scaling></c:valAx></c:plotArea>";
        let bounds = metadata(both).value_axis.unwrap();
        assert_eq!(bounds.min, Some(-2.5));
        assert_eq!(bounds.max, Some(100.0));

        let max_only = "<c:plotArea><c:barChart/><c:valAx><c:scaling><c:orientation val=\"minMax\"/><c:max val=\"40\"/></c:scaling></c:valAx></c:plotArea>";
        let bounds = metadata(max_only).value_axis.unwrap();
        assert_eq!(bounds.min, None);
        assert_eq!(bounds.max, Some(40.0));

        let auto = "<c:plotArea><c:barChart/><c:valAx><c:scaling><c:orientation val=\"minMax\"/></c:scaling></c:valAx></c:plotArea>";
        assert!(metadata(auto).value_axis.is_none());
        assert!(metadata("<c:plotArea><c:pieChart/></c:plotArea>").value_axis.is_none());
    }

    /// Issue #182: category number formats survive into the metadata so the
    /// renderer can show `Jan-22` instead of the raw serial 44562.
    #[test]
    fn reads_category_formats_from_num_cache_and_axis() {
        let dated = r#"<c:plotArea><c:barChart><c:ser>
            <c:cat><c:numRef><c:f>D!$A$2</c:f><c:numCache><c:formatCode>mmm\-yy</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>44562</c:v></c:pt><c:pt idx="1"><c:v>44593</c:v></c:pt></c:numCache></c:numRef></c:cat>
            <c:val><c:numRef><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart><c:catAx/><c:valAx/></c:plotArea>"#;
        let chart = metadata(dated);
        let series = &chart.series[0];
        assert_eq!(series.category_format.as_deref(), Some("mmm\\-yy"));
        assert_eq!(series.number_format.as_deref(), Some("0.00"));
        assert_eq!(series.categories, vec!["44562", "44593"]);
        assert!(chart.category_axis_format.is_none());

        let axis_level = r#"<c:plotArea><c:barChart/><c:catAx><c:numFmt formatCode="0.0%" sourceLinked="0"/></c:catAx><c:valAx/></c:plotArea>"#;
        assert_eq!(metadata(axis_level).category_axis_format.as_deref(), Some("0.0%"));
        let date_axis = r#"<c:plotArea><c:lineChart/><c:dateAx><c:numFmt formatCode="mmm\-yy" sourceLinked="1"/></c:dateAx><c:valAx/></c:plotArea>"#;
        assert_eq!(metadata(date_axis).category_axis_format.as_deref(), Some("mmm\\-yy"));

        // scatter X data (c:xVal) carries the same field
        let scatter = r#"<c:plotArea><c:scatterChart><c:ser>
            <c:xVal><c:numRef><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>0.15</c:v></c:pt></c:numCache></c:numRef></c:xVal>
            <c:yVal><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>0.4</c:v></c:pt></c:numCache></c:numRef></c:yVal>
        </c:ser></c:scatterChart></c:plotArea>"#;
        assert_eq!(metadata(scatter).series[0].category_format.as_deref(), Some("0%"));

        // string categories carry no format
        let plain = "<c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>a</c:v></c:pt></c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea>";
        assert!(metadata(plain).series[0].category_format.is_none());
    }

    /// Issue #180: a scatter chart's first valAx in document order is the X
    /// axis; gridlines/bounds must come from the left (Y) one.
    #[test]
    fn value_axis_prefers_the_left_axis() {
        let scatter = r#"<c:plotArea><c:scatterChart/>
            <c:valAx><c:axId val="1"/><c:scaling><c:max val="10"/></c:scaling><c:delete val="0"/><c:axPos val="b"/></c:valAx>
            <c:valAx><c:axId val="2"/><c:scaling><c:max val="0.45"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/></c:valAx>
        </c:plotArea>"#;
        let chart = metadata(scatter);
        assert_eq!(chart.value_axis.unwrap().max, Some(0.45));
        assert_eq!(chart.gridlines, Some(true));
        // No axPos="l" (horizontal bar puts the value axis at the bottom):
        // the document-order fallback still finds it.
        let bar = r#"<c:plotArea><c:barChart/><c:catAx/><c:valAx><c:scaling><c:max val="7"/></c:scaling><c:axPos val="b"/><c:majorGridlines/></c:valAx></c:plotArea>"#;
        let chart = metadata(bar);
        assert_eq!(chart.value_axis.unwrap().max, Some(7.0));
        assert_eq!(chart.gridlines, Some(true));
    }

    #[test]
    fn reads_gap_width_and_hole_size() {
        let bar = "<c:plotArea><c:barChart><c:gapWidth val=\"80\"/></c:barChart></c:plotArea>";
        assert_eq!(metadata(bar).gap_width_pct, Some(80));
        // Missing gapWidth stays absent; the default is the consumer's call.
        assert!(metadata("<c:plotArea><c:barChart/></c:plotArea>").gap_width_pct.is_none());

        let doughnut = "<c:plotArea><c:doughnutChart><c:holeSize val=\"65\"/></c:doughnutChart></c:plotArea>";
        assert_eq!(metadata(doughnut).hole_size_pct, Some(65));
        assert!(metadata("<c:plotArea><c:doughnutChart/></c:plotArea>").hole_size_pct.is_none());
    }

    #[test]
    fn reads_series_and_point_explosions() {
        let body = r#"<c:plotArea><c:pieChart><c:ser>
            <c:explosion val="12"/>
            <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr><c:explosion val="25"/></c:dPt>
            <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></c:spPr></c:dPt>
            <c:dPt><c:idx val="2"/><c:explosion val="40"/></c:dPt>
        </c:ser></c:pieChart></c:plotArea>"#;
        let chart = metadata(body);
        let series = &chart.series[0];
        assert_eq!(series.explosion_pct, Some(12));
        let explosions = series.point_explosions.as_ref().unwrap();
        assert_eq!(explosions.len(), 2);
        assert_eq!((explosions[0].index, explosions[0].pct), (0, 25));
        assert_eq!((explosions[1].index, explosions[1].pct), (2, 40));
        // dPt 0 keeps its color even though it also carries an explosion.
        let points = series.point_colors.as_ref().unwrap();
        assert_eq!(points.len(), 2);
        assert_eq!((points[0].index, points[0].color.as_str()), (0, "#FF8800"));

        let plain = "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser></c:pieChart></c:plotArea>";
        assert!(metadata(plain).series[0].explosion_pct.is_none());
        assert!(metadata(plain).series[0].point_explosions.is_none());
    }

    #[test]
    fn serializes_category_formats_with_expected_json_names() {
        let body = r#"<c:plotArea><c:barChart><c:ser>
            <c:cat><c:numRef><c:numCache><c:formatCode>mmm\-yy</c:formatCode><c:pt idx="0"><c:v>44562</c:v></c:pt></c:numCache></c:numRef></c:cat>
        </c:ser></c:barChart><c:catAx><c:numFmt formatCode="d-mmm" sourceLinked="0"/></c:catAx><c:valAx/></c:plotArea>"#;
        let json = serde_json::to_value(metadata(body)).unwrap();
        assert_eq!(json["categoryAxisFormat"], "d-mmm");
        assert_eq!(json["series"][0]["categoryFormat"], "mmm\\-yy");

        let plain = "<c:plotArea><c:pieChart><c:ser/></c:pieChart></c:plotArea>";
        let json = serde_json::to_value(metadata(plain)).unwrap();
        assert!(json.get("categoryAxisFormat").is_none());
        assert!(json["series"][0].get("categoryFormat").is_none());
    }

    #[test]
    fn serializes_new_fields_with_expected_json_names() {
        let body = r#"<c:plotArea><c:barChart><c:grouping val="stacked"/>
            <c:ser><c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr></c:dPt></c:ser>
            <c:dLbls><c:showVal val="1"/><c:dLblPos val="inEnd"/><c:numFmt formatCode="0.00" sourceLinked="0"/></c:dLbls><c:gapWidth val="150"/></c:barChart>
            <c:catAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>
            <c:valAx><c:scaling><c:max val="120.5"/></c:scaling><c:majorGridlines/></c:valAx></c:plotArea>
            <c:legend><c:legendPos val="b"/></c:legend>"#;
        let json = serde_json::to_value(metadata(body)).unwrap();
        assert_eq!(json["legend"], "bottom");
        assert_eq!(json["dataLabels"], "value");
        assert_eq!(json["dataLabelPosition"], "inside-end");
        assert_eq!(json["dataLabelFormat"], "0.00");
        assert_eq!(json["grouping"], "stacked");
        assert_eq!(json["axisTitles"]["category"], "Month");
        assert!(json["axisTitles"].get("value").is_none());
        assert_eq!(
            json["series"][0]["pointColors"],
            serde_json::json!([{ "index": 1, "color": "#00AA00" }])
        );
        assert_eq!(json["gridlines"], true);
        assert_eq!(json["valueAxis"], serde_json::json!({ "max": 120.5 }));
        assert_eq!(json["gapWidthPct"], 150);
        assert!(json.get("holeSizePct").is_none());

        let doughnut = r#"<c:plotArea><c:doughnutChart><c:holeSize val="50"/>
            <c:ser><c:explosion val="10"/>
            <c:dPt><c:idx val="2"/><c:explosion val="30"/></c:dPt></c:ser>
            </c:doughnutChart></c:plotArea>"#;
        let json = serde_json::to_value(metadata(doughnut)).unwrap();
        assert!(json.get("dataLabelPosition").is_none());
        assert!(json.get("dataLabelFormat").is_none());
        assert!(json.get("gridlines").is_none());
        assert!(json.get("valueAxis").is_none());
        assert!(json.get("gapWidthPct").is_none());
        assert_eq!(json["holeSizePct"], 50);
        assert_eq!(json["series"][0]["explosionPct"], 10);
        assert_eq!(
            json["series"][0]["pointExplosions"],
            serde_json::json!([{ "index": 2, "pct": 30 }])
        );
        assert!(json["series"][0].get("pointColors").is_none());
    }
}

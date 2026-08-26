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
    /// alignment/@shrinkToFit — Excel scales the font down to fit the column
    /// instead of clipping. Omitted when false to keep payloads small.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// Theme provenance (slot index + tint) for colors resolved from the
    /// theme palette, so the renderer can re-resolve them when the document
    /// theme changes. Absent for literal rgb / indexed colors.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color_theme: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color_tint: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color_theme: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color_tint: Option<f64>,
    /// font/scheme: "major" or "minor" — the family follows the theme fonts.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_scheme: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    /// OOXML alignment/@textRotation: 1-90 counter-clockwise, 91-180 encodes
    /// clockwise as 90+deg, 255 is vertically stacked.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u32>,
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
    /// Table-style dxf inner grid edges (<horizontal>/<vertical>) — consumed
    /// by the custom table palette only, never serialized per cell.
    #[serde(skip)]
    pub border_inner_horizontal: Option<BorderEdge>,
    #[serde(skip)]
    pub border_inner_vertical: Option<BorderEdge>,
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
            || self.text_rotation != default.text_rotation
            || self.wrap_text != default.wrap_text
            || self.shrink_to_fit != default.shrink_to_fit
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
    /// True when the file carried a real `<xdr:to>` marker. Excel clamps
    /// such an offset at its cell edge; synthesized to markers
    /// (oneCellAnchor ext, absoluteAnchor, group children) encode sizes as
    /// offsets past the edge and must keep walking.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub explicit_to: bool,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: String,
    /// `c:tx/c:strRef/c:f` when the series name is a cell reference whose
    /// cache is missing — the renderer resolves it from the live cells.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name_ref: Option<String>,
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
    /// spPr/a:ln color; "none" for an explicit a:noFill line.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// spPr/a:ln/@w converted from EMU to CSS px (w / 12700 pt · 96/72).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
    /// c:marker/c:symbol — "none" hides scatter/line markers.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker: Option<String>,
    /// First outer level of a multiLvlStrCache category axis; start/end are
    /// positions in the compacted innermost `categories` (end exclusive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category_groups: Option<Vec<CategoryGroup>>,
}

/// One outer-level group label spanning innermost categories [start, end).
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CategoryGroup {
    pub label: String,
    pub start: usize,
    pub end: usize,
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

/// One paragraph of a shape/text-box `xdr:txBody`, with Excel's run styling.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeParagraph {
    /// a:pPr/@algn — l | ctr | r | just; absent means left.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    pub runs: Vec<ShapeRun>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeRun {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "is_false")]
    pub bold: bool,
    #[serde(skip_serializing_if = "is_false")]
    pub italic: bool,
    #[serde(skip_serializing_if = "is_false")]
    pub underline: bool,
    /// Points (a:rPr/@sz / 100).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
}

fn is_false(value: &bool) -> bool {
    !*value
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

/// One plot axis, keyed by its `c:axPos` side rather than element kind, so
/// scatter charts (two valAx) pair titles/bounds with the right axis.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AxisInfo {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    pub major_gridlines: bool,
    /// c:delete — the axis exists for scaling but is not drawn.
    pub hidden: bool,
    /// c:scaling/c:orientation val="maxMin" — categories/values run reversed.
    pub reversed: bool,
}

/// Chart-title font shorthand from c:title/c:txPr//a:defRPr.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartTitleStyle {
    /// Points (defRPr/@sz / 100).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
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
    /// Absent when the part has no dLbls at all (renderer defaults apply);
    /// "none" is an explicit off.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_labels: Option<String>,
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
    /// Bottom/top axis (category, or scatter X), by axPos.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x_axis: Option<AxisInfo>,
    /// Left/right axis (values), by axPos.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y_axis: Option<AxisInfo>,
    /// c:scatterStyle — whether scatter points connect with lines.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    /// Plot-level `c:lineChart/c:marker` flag; per-series symbols refine it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_markers: Option<bool>,
    /// Second left/right value axis (combo charts) — scaling for the line
    /// series and, when not hidden, a drawn right-hand scale.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub secondary_y_axis: Option<AxisInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_style: Option<ChartTitleStyle>,
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
    /// a:blip/a:alphaModFix amt as 0..1; absent when fully opaque.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    /// spPr/a:blipFill on a shape — the image painted clipped to the
    /// preset geometry (the flat fill_color stays the loading fallback).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_media_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_media_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shape_type: Option<String>,
    /// a:custGeom pathLst as one SVG path string in the path coordinate
    /// space (moveTo/lnTo/beziers/close; shapes with arcs stay unsupported).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_path: Option<CustomPath>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    /// xdr:style fillRef resolved against a theme fillStyleLst gradient;
    /// fill_color stays the flat approximation.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_gradient: Option<FillGradient>,
    /// spPr/a:ln solid color, or the xdr:style lnRef theme color; "none"
    /// for an explicit a:noFill outline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// a:ln/@w in points.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_width: Option<f64>,
    /// a:ln/a:prstDash/@val — solid when absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_dash: Option<String>,
    /// a:ln/@cap — rnd | sq | flat.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_cap: Option<String>,
    /// a:xfrm/@flipH, @flipV — mirror the preset geometry.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// xdr:style fontRef theme color — the default run color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_color: Option<String>,
    /// a:bodyPr/@anchor — t | ctr | b.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraphs: Option<Vec<ShapeParagraph>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// Degrees clockwise.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    /// a:xfrm ext in EMU — the true unrotated frame of a rotated shape.
    /// The anchor stores rotated bounds (Excel: quadrant-swapped snap rect,
    /// LibreOffice: the AABB); both keep the anchor center on the shape
    /// center, so the renderer restores ext around it before rotating.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frame_width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frame_height: Option<f64>,
    /// xdr:cNvPr/@id — pairs a drawing fallback shape with its worksheet
    /// <oleObject shapeId=…>. Engine-internal, never serialized.
    #[serde(skip)]
    pub nv_id: Option<u32>,
    /// ZIP entry path of the drawing part this visual lives in, plus its
    /// anchor index within that part — the save-side edit locator.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub drawing_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub drawing_index: Option<usize>,
}

/// A custGeom outline: `d` uses the `<a:path>` coordinate space so the
/// renderer scales it into the anchor frame (degenerate 0 extents become 1).
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CustomPath {
    pub width: f64,
    pub height: f64,
    pub d: String,
    /// True when every subpath is stroke-only (`<a:path fill="none">`).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub stroke_only: bool,
    /// Fillable subpaths only, present when the geometry mixes filled and
    /// stroke-only subpaths — filling `d` would paint the stroke-only ones.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_d: Option<String>,
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
pub(crate) struct Relationship {
    pub(crate) target: String,
    pub(crate) relationship_type: String,
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
    color_theme: Option<usize>,
    color_tint: Option<f64>,
    scheme: Option<String>,
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
    // <vertical>/<horizontal>: inner grid edges, only meaningful in
    // table-style dxfs.
    vertical: Option<BorderEdge>,
    horizontal: Option<BorderEdge>,
}

/// Theme palette in `theme` attribute index order (0↔1 and 2↔3 are swapped
/// versus the clrScheme document order, per the xlsx theme index mapping).
#[derive(Clone, Default)]
pub struct ColorContext {
    theme: Vec<(u8, u8, u8)>,
    /// fmtScheme/fillStyleLst entries (1-based fillRef idx order); None for
    /// non-gradient entries.
    fill_styles: Vec<Option<ThemeGradient>>,
    /// styles.xml colors/indexedColors override of the legacy palette
    /// (hex without '#'); indexes past its end fall back to the builtin.
    indexed: Vec<String>,
}

/// A theme gradient with phClr stops: the placeholder resolves to the
/// fillRef color at use time, then each stop's transforms apply.
#[derive(Clone, Debug)]
pub struct ThemeGradient {
    pub stops: Vec<ThemeGradientStop>,
    /// Degrees clockwise, 0 = left-to-right.
    pub angle: f64,
}

#[derive(Clone, Debug)]
pub struct ThemeGradientStop {
    /// 0..1 along the gradient axis.
    pub position: f64,
    /// (transform tag, val/100000) pairs in document order.
    pub modifiers: Vec<(String, f64)>,
}

/// A shape gradient with fully resolved stop colors, ready to render.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FillGradient {
    /// Degrees clockwise, 0 = left-to-right.
    pub angle: f64,
    pub stops: Vec<FillGradientStop>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FillGradientStop {
    /// 0..1 along the gradient axis.
    pub position: f64,
    pub color: String,
}

impl ColorContext {
    /// The palette as `#RRGGBB` strings in theme index order, or None when
    /// the workbook has no readable theme.
    pub fn palette_hex(&self) -> Option<Vec<String>> {
        if self.theme.is_empty() {
            return None;
        }
        Some(
            self.theme
                .iter()
                .map(|(red, green, blue)| format!("#{red:02X}{green:02X}{blue:02X}"))
                .collect(),
        )
    }
}

/// Major/minor latin typefaces from the theme's fontScheme.
#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFonts {
    pub major: String,
    pub minor: String,
    /// minorFont `<a:ea typeface>` when non-empty: the East-Asian face a CJK
    /// Excel resolves scheme="minor" fonts to (the latin face only covers
    /// Latin text, but column-width MDW follows the Normal font's EA face).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_ea: Option<String>,
}

pub fn read_theme_fonts(
    archive: &mut ZipArchive<File>,
) -> Result<Option<ThemeFonts>, SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/theme/theme1.xml")? else {
        return Ok(None);
    };
    let document = parse_document(&xml, "theme1.xml")?;
    let Some(scheme) = document
        .descendants()
        .find(|node| node.has_tag_name("fontScheme"))
    else {
        return Ok(None);
    };
    let typeface = |name: &str, script: &str| -> Option<String> {
        scheme
            .children()
            .find(|child| child.has_tag_name(name))?
            .children()
            .find(|child| child.has_tag_name(script))?
            .attribute("typeface")
            .filter(|value| !value.is_empty())
            .map(ToOwned::to_owned)
    };
    Ok(
        match (
            typeface("majorFont", "latin"),
            typeface("minorFont", "latin"),
        ) {
            (Some(major), Some(minor)) => Some(ThemeFonts {
                major,
                minor,
                minor_ea: typeface("minorFont", "ea"),
            }),
            _ => None,
        },
    )
}

/// Children with the given local tag, resolving mc:AlternateContent wrappers
/// (ECMA-376 Part 3): a core-only consumer takes mc:Fallback (first mc:Choice
/// when a producer omits the fallback). Hancom exports wrap individual
/// font/xf/dxf entries this way; skipping them shifted every later
/// fontId/fillId/borderId/dxfId.
fn mc_children<'a, 'input>(parent: Node<'a, 'input>, tag: &str) -> Vec<Node<'a, 'input>> {
    let mut nodes = Vec::new();
    collect_mc_children(parent, tag, &mut nodes);
    nodes
}

fn collect_mc_children<'a, 'input>(
    parent: Node<'a, 'input>,
    tag: &str,
    nodes: &mut Vec<Node<'a, 'input>>,
) {
    for child in parent.children() {
        if child.has_tag_name(tag) {
            nodes.push(child);
        } else if child.has_tag_name("AlternateContent") {
            let branch = child
                .children()
                .find(|node| node.has_tag_name("Fallback"))
                .or_else(|| child.children().find(|node| node.has_tag_name("Choice")));
            if let Some(branch) = branch {
                collect_mc_children(branch, tag, nodes);
            }
        }
    }
}

pub fn read_styles(
    archive: &mut ZipArchive<File>,
    colors: &ColorContext,
    theme_fonts: Option<&ThemeFonts>,
    locale: &str,
    short_date_format: Option<&str>,
) -> Result<(Vec<CellStyle>, Vec<CellStyle>, Option<String>), SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/styles.xml")? else {
        return Ok((vec![CellStyle::default()], Vec::new(), None));
    };
    let document = parse_document(&xml, "styles.xml")?;
    // Only the top-level <numFmts> table: dxf-local <numFmt> entries reuse
    // file-local ids that must not shadow builtin ids for cell xfs.
    let custom_formats = document
        .descendants()
        .filter(|node| node.has_tag_name("numFmts"))
        .flat_map(|node| mc_children(node, "numFmt"))
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
            mc_children(node, "font")
                .into_iter()
                .map(|font| parse_font(font, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let fills = document
        .descendants()
        .find(|node| node.has_tag_name("fills"))
        .map(|node| {
            mc_children(node, "fill")
                .into_iter()
                .map(|fill| parse_fill(fill, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let borders = document
        .descendants()
        .find(|node| node.has_tag_name("borders"))
        .map(|node| {
            mc_children(node, "border")
                .into_iter()
                .map(|border| parse_border(border, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let cell_xfs = document
        .descendants()
        .find(|node| node.has_tag_name("cellXfs"));
    // Literal cached <name val> of the Normal (cellXfs[0]) font: for scheme
    // fonts the theme substitution below erases it, but Excel derives the
    // column-width MDW from this face (a ja workbook caches the locale
    // resolution, e.g. MS PGothic, while the theme latin says Calibri).
    let normal_font_name = cell_xfs
        .and_then(|node| mc_children(node, "xf").into_iter().next())
        .and_then(|xf| numeric_attribute(xf, "fontId"))
        .and_then(|index| fonts.get(index))
        .and_then(|font| font.family.clone());
    let styles = cell_xfs
        .map(|node| {
            mc_children(node, "xf")
                .into_iter()
                .map(|xf| {
                    let font = numeric_attribute(xf, "fontId")
                        .and_then(|index| fonts.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let fill = numeric_attribute(xf, "fillId")
                        .and_then(|index| fills.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let border = numeric_attribute(xf, "borderId")
                        .and_then(|index| borders.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let number_format = numeric_attribute(xf, "numFmtId").and_then(|id| {
                        custom_formats
                            .get(&(id as u32))
                            .cloned()
                            .or_else(|| short_date_number_format(id as u32, short_date_format))
                            .or_else(|| {
                                builtin_number_format(id as u32, locale).map(ToOwned::to_owned)
                            })
                    });
                    let alignment = xf.children().find(|child| child.has_tag_name("alignment"));
                    // Excel resolves scheme fonts against the theme; the
                    // literal <name val> is only a cached copy.
                    let font_family = match (font.scheme.as_deref(), theme_fonts) {
                        (Some("major"), Some(fonts)) => Some(fonts.major.clone()),
                        (Some("minor"), Some(fonts)) => Some(fonts.minor.clone()),
                        _ => font.family,
                    };
                    CellStyle {
                        font_family,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        underline: font.underline,
                        strikethrough: font.strikethrough,
                        wrap_text: alignment
                            .and_then(|node| node.attribute("wrapText"))
                            .is_some_and(|value| value == "1" || value == "true"),
                        shrink_to_fit: alignment
                            .and_then(|node| node.attribute("shrinkToFit"))
                            .is_some_and(|value| value == "1" || value == "true"),
                        font_color: font.color,
                        fill_color: fill.color,
                        font_color_theme: font.color_theme,
                        font_color_tint: font.color_tint,
                        fill_color_theme: fill.theme,
                        fill_color_tint: fill.tint,
                        font_scheme: font.scheme,
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
                        text_rotation: alignment
                            .and_then(|node| node.attribute("textRotation"))
                            .and_then(|value| value.parse::<u32>().ok())
                            .filter(|degrees| (1..=180).contains(degrees) || *degrees == 255),
                        number_format,
                        border_top: border.top,
                        border_bottom: border.bottom,
                        border_left: border.left,
                        border_right: border.right,
                        border_diagonal: border.diagonal,
                        diagonal_up: border.diagonal_up,
                        diagonal_down: border.diagonal_down,
                        border_inner_horizontal: None,
                        border_inner_vertical: None,
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let dxfs = document
        .descendants()
        .find(|node| node.has_tag_name("dxfs"))
        .map(|node| {
            mc_children(node, "dxf")
                .into_iter()
                .map(|dxf| parse_dxf(dxf, colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let styles = if styles.is_empty() {
        vec![CellStyle::default()]
    } else {
        styles
    };
    Ok((styles, dxfs, normal_font_name))
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
                .or_else(|| pattern.children().find(|node| node.has_tag_name("fgColor")))
                .and_then(|node| parse_color(node, colors))
        });
    let border = dxf
        .children()
        .find(|node| node.has_tag_name("border"))
        .map(|node| parse_border(node, colors))
        .unwrap_or_default();
    // Unlike cell xfs, a dxf carries its format code inline.
    let number_format = dxf
        .children()
        .find(|node| node.has_tag_name("numFmt"))
        .and_then(|node| node.attribute("formatCode"))
        .filter(|code| !code.is_empty() && *code != "General")
        .map(ToOwned::to_owned);
    CellStyle {
        font_family: font.family,
        font_size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        wrap_text: false,
        shrink_to_fit: false,
        font_color: font.color,
        fill_color,
        font_color_theme: None,
        font_color_tint: None,
        fill_color_theme: None,
        fill_color_tint: None,
        font_scheme: None,
        horizontal_alignment: None,
        vertical_alignment: None,
        indent: None,
        text_rotation: None,
        number_format,
        border_top: border.top,
        border_bottom: border.bottom,
        border_left: border.left,
        border_right: border.right,
        border_diagonal: border.diagonal,
        diagonal_up: border.diagonal_up,
        diagonal_down: border.diagonal_down,
        border_inner_horizontal: border.horizontal,
        border_inner_vertical: border.vertical,
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
        let drawing_path =
            resolve_part_target(&sheet.worksheet_path, &drawing_relationship.target)?;
        let prog_ids = read_ole_prog_ids(archive, &sheet.worksheet_path)?;
        let start = visuals.len();
        visuals.extend(read_drawing(
            archive,
            &drawing_path,
            &sheet.sheet_id,
            visuals.len(),
            colors,
            &prog_ids,
        )?);
        // OLE embeds keep a hidden drawing fallback shape; give it the
        // object's progId as caption text so the placeholder reads as an
        // embedded object instead of an empty rectangle.
        if !prog_ids.is_empty() {
            for visual in &mut visuals[start..] {
                if visual.kind != "shape" || visual.text.is_some() {
                    continue;
                }
                let Some(prog_id) = visual.nv_id.and_then(|id| prog_ids.get(&id)) else {
                    continue;
                };
                visual.text = Some(prog_id.clone());
            }
        }
    }
    Ok(visuals)
}

/// Worksheet `<oleObject shapeId=… progId=…>` pairs (deduplicated — the
/// mc:AlternateContent choice/fallback repeat the same object).
fn read_ole_prog_ids(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<HashMap<u32, String>, SidecarError> {
    let xml = read_xml(archive, worksheet_path)?;
    let mut prog_ids = HashMap::new();
    if !xml.contains("oleObject") {
        return Ok(prog_ids);
    }
    let document = parse_document(&xml, worksheet_path)?;
    for node in document
        .descendants()
        .filter(|node| node.has_tag_name("oleObject"))
    {
        let Some(shape_id) = node
            .attribute("shapeId")
            .and_then(|value| value.parse::<u32>().ok())
        else {
            continue;
        };
        let Some(prog_id) = node.attribute("progId") else {
            continue;
        };
        prog_ids
            .entry(shape_id)
            .or_insert_with(|| prog_id.to_owned());
    }
    Ok(prog_ids)
}

/// DrawingML solid fill: srgbClr, or schemeClr resolved via the theme palette.
fn drawing_fill_color(node: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    // An a:ln below `node` carries the outline color, not the fill.
    let outside_outline = |child: &Node<'_, '_>| {
        !child
            .ancestors()
            .take_while(|ancestor| *ancestor != node)
            .any(|ancestor| ancestor.has_tag_name("ln"))
    };
    let fill = node
        .descendants()
        .find(|child| child.has_tag_name("solidFill") && outside_outline(child))
        .or_else(|| {
            // Gradient fills approximate to their first stop color.
            node.descendants()
                .find(|child| child.has_tag_name("gradFill") && outside_outline(child))
                .and_then(|grad| grad.descendants().find(|child| child.has_tag_name("gs")))
        })?;
    if let Some(srgb) = fill
        .descendants()
        .find(|child| child.has_tag_name("srgbClr"))
    {
        return srgb.attribute("val").map(|value| format!("#{value}"));
    }
    let scheme = fill
        .descendants()
        .find(|child| child.has_tag_name("schemeClr"))?;
    let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
    Some(apply_color_modifiers(scheme, base))
}

/// DrawingML child color transforms (a:lumMod/a:lumOff/a:tint/a:shade), the
/// usual RGB approximations — enough for Excel's accent "40% lighter" dPt
/// and legend variants.
fn apply_color_modifiers(color_node: Node<'_, '_>, base: (u8, u8, u8)) -> String {
    let modifiers = color_node
        .children()
        .filter(|child| child.is_element())
        .filter_map(|child| {
            Some((
                child.tag_name().name().to_owned(),
                child.attribute("val")?.parse::<f64>().ok()? / 100_000.0,
            ))
        })
        .collect::<Vec<_>>();
    apply_modifier_values(base, &modifiers)
}

/// xdr:style fillRef pointing at a fmtScheme gradient entry: each phClr
/// stop resolves against the reference's scheme color.
fn style_fill_gradient(
    style_node: Option<Node<'_, '_>>,
    colors: &ColorContext,
) -> Option<FillGradient> {
    let reference = style_node?
        .children()
        .find(|node| node.has_tag_name("fillRef"))?;
    let idx = reference.attribute("idx")?.parse::<usize>().ok()?;
    let gradient = colors.fill_styles.get(idx.checked_sub(1)?)?.as_ref()?;
    let scheme = reference
        .children()
        .find(|node| node.has_tag_name("schemeClr"))?;
    let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
    let base = parse_hex_rgb(apply_color_modifiers(scheme, base).trim_start_matches('#'))?;
    let stops = gradient
        .stops
        .iter()
        .map(|stop| FillGradientStop {
            position: stop.position,
            color: apply_modifier_values(base, &stop.modifiers),
        })
        .collect();
    Some(FillGradient {
        angle: gradient.angle,
        stops,
    })
}

/// Data-driven twin of apply_color_modifiers, for theme gradient stops whose
/// source document is gone by the time the fillRef color is known.
fn apply_modifier_values(base: (u8, u8, u8), modifiers: &[(String, f64)]) -> String {
    let mut channels = [f64::from(base.0), f64::from(base.1), f64::from(base.2)];
    for (name, value) in modifiers {
        let value = *value;
        match name.as_str() {
            "lumMod" | "shade" => {
                for channel in &mut channels {
                    *channel *= value;
                }
            }
            "lumOff" => {
                for channel in &mut channels {
                    *channel += 255.0 * value;
                }
            }
            "tint" => {
                for channel in &mut channels {
                    *channel = *channel * value + 255.0 * (1.0 - value);
                }
            }
            "satMod" => {
                let clamp = |v: f64| v.round().clamp(0.0, 255.0) as u8;
                let (hue, saturation, luminance) =
                    rgb_to_hsl((clamp(channels[0]), clamp(channels[1]), clamp(channels[2])));
                let (red, green, blue) = hsl_to_rgb(hue, (saturation * value).min(1.0), luminance);
                channels = [f64::from(red), f64::from(green), f64::from(blue)];
            }
            _ => {}
        }
    }
    let clamp = |value: f64| value.round().clamp(0.0, 255.0) as u8;
    format!(
        "#{:02X}{:02X}{:02X}",
        clamp(channels[0]),
        clamp(channels[1]),
        clamp(channels[2])
    )
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
    let mut entry = crate::zip_entry(archive, media_path)?;
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

/// xdr:txBody paragraphs with per-run styling; a:br becomes a newline run.
fn parse_text_paragraphs(body: Node<'_, '_>, colors: &ColorContext) -> Vec<ShapeParagraph> {
    body.children()
        .filter(|node| node.has_tag_name("p"))
        .map(|paragraph| {
            let align = direct_child(paragraph, "pPr")
                .and_then(|node| node.attribute("algn"))
                .map(ToOwned::to_owned);
            let mut runs = Vec::new();
            for child in paragraph.children() {
                if child.has_tag_name("br") {
                    runs.push(ShapeRun {
                        text: "\n".into(),
                        color: None,
                        bold: false,
                        italic: false,
                        underline: false,
                        size: None,
                    });
                    continue;
                }
                if !child.has_tag_name("r") {
                    continue;
                }
                let text = direct_child(child, "t")
                    .and_then(|node| node.text())
                    .unwrap_or_default()
                    .to_owned();
                let properties = direct_child(child, "rPr");
                let flag = |name: &str| {
                    properties
                        .and_then(|rpr| rpr.attribute(name))
                        .is_some_and(|value| value == "1" || value == "true")
                };
                runs.push(ShapeRun {
                    text,
                    color: properties.and_then(|rpr| drawing_fill_color(rpr, colors)),
                    bold: flag("b"),
                    italic: flag("i"),
                    underline: properties
                        .and_then(|rpr| rpr.attribute("u"))
                        .is_some_and(|value| value != "none"),
                    size: properties
                        .and_then(|rpr| rpr.attribute("sz"))
                        .and_then(|value| value.parse::<f64>().ok())
                        .map(|value| value / 100.0),
                });
            }
            ShapeParagraph { align, runs }
        })
        .collect()
}

fn read_drawing(
    archive: &mut ZipArchive<File>,
    drawing_path: &str,
    sheet_id: &str,
    id_offset: usize,
    colors: &ColorContext,
    ole_prog_ids: &HashMap<u32, String>,
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
        // Excel does not render objects flagged hidden (cNvPr hidden="1").
        // OLE placeholder shapes are the exception: their hidden fallback
        // shape is the visible render.
        let nv = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"));
        let hidden = nv.is_some_and(hidden_attribute);
        let is_ole_placeholder = nv
            .and_then(|node| node.attribute("id"))
            .and_then(|value| value.parse::<u32>().ok())
            .is_some_and(|id| ole_prog_ids.contains_key(&id));
        if hidden && !is_ole_placeholder {
            continue;
        }
        let visual_id = format!("visual-{}", id_offset + index + 1);
        if let Some(group) = anchor_node
            .children()
            .find(|node| node.has_tag_name("grpSp"))
        {
            // The group box spans the whole anchor; its EMU size comes from
            // the group's own xfrm ext (Excel keeps the two equal).
            let xfrm = group_xfrm(group);
            let width = xfrm_value(xfrm, "ext", "cx").unwrap_or(0.0);
            let height = xfrm_value(xfrm, "ext", "cy").unwrap_or(0.0);
            if width > 0.0 && height > 0.0 {
                let mut counter = 0;
                expand_group(
                    group,
                    &anchor,
                    (0.0, 0.0, width, height),
                    &visual_id,
                    &mut counter,
                    sheet_id,
                    colors,
                    drawing_path,
                    &relationships,
                    &mut visuals,
                )?;
            }
            continue;
        }
        if let Some(chart_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("chart"))
        {
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
                opacity: None,
                fill_media_path: None,
                fill_media_type: None,
                name: drawing_name(anchor_node),
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        // Only an xdr:pic is a picture — an xdr:sp with a:blipFill keeps its
        // geometry and outline and falls through to the shape branch.
        if let Some(pic_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("pic"))
        {
            let Some(blip_node) = pic_node
                .descendants()
                .find(|node| node.has_tag_name("blip"))
            else {
                continue;
            };
            let Some(id) = relationship_id(blip_node) else {
                continue;
            };
            let Some(relationship) = relationships.get(&id) else {
                continue;
            };
            // Excel draws nothing for a picture whose anchor collapses to a
            // point (from == to with equal offsets, or a 0x0 ext).
            if anchor_is_zero_extent(&anchor) {
                continue;
            }
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
                opacity: blip_opacity(blip_node),
                fill_media_path: None,
                fill_media_type: None,
                name: drawing_name(anchor_node),
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        if let Some(shape_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("sp") || node.has_tag_name("cxnSp"))
        {
            visuals.push(shape_visual(
                shape_node,
                anchor,
                visual_id,
                sheet_id,
                drawing_name(anchor_node),
                colors,
                drawing_path,
                &relationships,
                Some(index),
            ));
        }
    }
    // Chart children of expanded groups carry only their part path.
    for visual in &mut visuals {
        if visual.kind == "chart" && visual.chart.is_none() {
            if let Some(chart_path) = visual.chart_path.clone() {
                visual.chart = Some(read_chart(archive, &chart_path, colors)?);
            }
        }
    }
    Ok(visuals)
}

/// One `sp`/`cxnSp` node → a shape visual placed at `anchor`. Shared by the
/// direct anchor branch and grpSp expansion (children pass their own name
/// and no drawing_index — anchor edits would rewrite the whole group).
#[allow(clippy::too_many_arguments)]
fn shape_visual(
    shape_node: Node<'_, '_>,
    anchor: DrawingAnchor,
    visual_id: String,
    sheet_id: &str,
    name: Option<String>,
    colors: &ColorContext,
    drawing_path: &str,
    relationships: &HashMap<String, Relationship>,
    drawing_index: Option<usize>,
) -> VisualObject {
    {
        {
            let shape_type = shape_node
                .descendants()
                .find(|node| node.has_tag_name("prstGeom"))
                .and_then(|node| node.attribute("prst"))
                .map(ToOwned::to_owned);
            let custom_path = if shape_type.is_none() {
                parse_custom_geometry(shape_node)
            } else {
                None
            };
            let shape_sppr = shape_node.children().find(|node| node.has_tag_name("spPr"));
            // An explicit <a:noFill/> directly under spPr means transparent —
            // it must not fall through to the xdr:style fillRef theme color.
            let has_no_fill = shape_sppr
                .is_some_and(|sppr| sppr.children().any(|node| node.has_tag_name("noFill")));
            let fill_color = if has_no_fill {
                Some("none".into())
            } else {
                shape_sppr.and_then(|sppr| drawing_fill_color(sppr, colors))
            };
            let fill_media_path = shape_sppr
                .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("blipFill")))
                .and_then(|fill| fill.descendants().find(|node| node.has_tag_name("blip")))
                .and_then(relationship_id)
                .and_then(|id| relationships.get(&id))
                .and_then(|relationship| {
                    resolve_part_target(drawing_path, &relationship.target).ok()
                });
            let fill_media_type = fill_media_path
                .as_deref()
                .and_then(media_type_for_path)
                .map(ToOwned::to_owned);
            let xfrm = shape_node
                .descendants()
                .find(|node| node.has_tag_name("xfrm"));
            let rotation = xfrm
                .and_then(|node| node.attribute("rot"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 60_000.0);
            // Only rotated shapes need the true frame — unrotated anchors
            // already are the frame.
            let frame_extent = |attribute: &str| {
                if rotation.is_none() {
                    return None;
                }
                xfrm_value(xfrm, "ext", attribute).filter(|value| *value > 0.0)
            };
            let frame_width = frame_extent("cx");
            let frame_height = frame_extent("cy");
            let flipped = |attribute: &str| {
                xfrm.and_then(|node| node.attribute(attribute))
                    .is_some_and(|value| value == "1" || value == "true")
            };
            let body = shape_node
                .descendants()
                .find(|node| node.has_tag_name("txBody"));
            let paragraphs = body
                .map(|node| parse_text_paragraphs(node, colors))
                .filter(|list| !list.is_empty());
            let text = paragraphs.as_ref().map(|list| {
                list.iter()
                    .map(|paragraph| {
                        paragraph
                            .runs
                            .iter()
                            .map(|run| run.text.as_str())
                            .collect::<String>()
                    })
                    .collect::<Vec<_>>()
                    .join("\n")
            });
            let text_anchor = body
                .and_then(|node| direct_child(node, "bodyPr"))
                .and_then(|node| node.attribute("anchor"))
                .map(ToOwned::to_owned);
            // xdr:style theme references are the fallback when spPr carries
            // no explicit fill/line (Excel's default for inserted shapes).
            let style_node = shape_node
                .children()
                .find(|node| node.has_tag_name("style"));
            let style_color = |name: &str| {
                let reference = style_node?
                    .children()
                    .find(|node| node.has_tag_name(name))?;
                let scheme = reference
                    .children()
                    .find(|node| node.has_tag_name("schemeClr"))?;
                let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
                Some(apply_color_modifiers(scheme, base))
            };
            let fill_gradient = if fill_color.is_none() {
                style_fill_gradient(style_node, colors)
            } else {
                None
            };
            let fill_color = fill_color.or_else(|| style_color("fillRef"));
            let line_node = shape_node
                .children()
                .find(|node| node.has_tag_name("spPr"))
                .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("ln")));
            let line_color = line_node
                .and_then(|ln| {
                    if ln.children().any(|node| node.has_tag_name("noFill")) {
                        return Some("none".into());
                    }
                    drawing_fill_color(ln, colors)
                })
                .or_else(|| style_color("lnRef"));
            let line_width = line_node
                .and_then(|ln| ln.attribute("w"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|emu| emu / 12_700.0);
            let line_dash = line_node
                .and_then(|ln| ln.children().find(|node| node.has_tag_name("prstDash")))
                .and_then(|node| node.attribute("val"))
                .map(ToOwned::to_owned);
            let line_cap = line_node
                .and_then(|ln| ln.attribute("cap"))
                .map(ToOwned::to_owned);
            let text_color = style_color("fontRef");
            let nv_id = shape_node
                .descendants()
                .find(|node| node.has_tag_name("cNvPr"))
                .and_then(|node| node.attribute("id"))
                .and_then(|value| value.parse::<u32>().ok());
            VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "shape".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_path: None,
                media_type: None,
                opacity: None,
                fill_media_path,
                fill_media_type,
                name,
                shape_type,
                custom_path,
                fill_color,
                fill_gradient,
                line_color,
                line_width,
                line_dash,
                line_cap,
                flip_h: flipped("flipH"),
                flip_v: flipped("flipV"),
                text_color,
                text_anchor,
                paragraphs,
                text,
                rotation,
                frame_width,
                frame_height,
                nv_id,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index,
            }
        }
    }
}

fn format_path_number(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value:.2}")
    }
}

/// Parse a shape's a:custGeom pathLst into one SVG path string. Multiple
/// `<a:path>` entries scale into the first path's coordinate space. Returns
/// None when a command has no SVG mapping (arcTo) — the caller falls back
/// to the placeholder frame.
fn parse_custom_geometry(shape_node: Node<'_, '_>) -> Option<CustomPath> {
    let geometry = shape_node
        .descendants()
        .find(|node| node.has_tag_name("custGeom"))?;
    let path_list = geometry
        .children()
        .find(|node| node.has_tag_name("pathLst"))?;
    let mut base_width = 0.0_f64;
    let mut base_height = 0.0_f64;
    let mut d = String::new();
    let mut fill_d = String::new();
    let mut stroke_only = true;
    for path in path_list
        .children()
        .filter(|node| node.has_tag_name("path"))
    {
        let dimension = |attribute: &str| {
            path.attribute(attribute)
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| *value > 0.0)
                .unwrap_or(0.0)
        };
        let width = dimension("w");
        let height = dimension("h");
        if d.is_empty() {
            base_width = width;
            base_height = height;
        }
        let path_fills = path.attribute("fill") != Some("none");
        if path_fills {
            stroke_only = false;
        }
        let segment_start = d.len();
        let scale_x = if width > 0.0 && base_width > 0.0 {
            base_width / width
        } else {
            1.0
        };
        let scale_y = if height > 0.0 && base_height > 0.0 {
            base_height / height
        } else {
            1.0
        };
        for command in path.children().filter(|node| node.is_element()) {
            let points: Vec<(f64, f64)> = command
                .children()
                .filter(|node| node.has_tag_name("pt"))
                .filter_map(|point| {
                    Some((
                        point.attribute("x")?.parse::<f64>().ok()? * scale_x,
                        point.attribute("y")?.parse::<f64>().ok()? * scale_y,
                    ))
                })
                .collect();
            let (letter, expected) = match command.tag_name().name() {
                "moveTo" => ("M", 1),
                "lnTo" => ("L", 1),
                "cubicBezTo" => ("C", 3),
                "quadBezTo" => ("Q", 2),
                "close" => ("Z", 0),
                _ => return None,
            };
            if points.len() < expected {
                return None;
            }
            if !d.is_empty() {
                d.push(' ');
            }
            d.push_str(letter);
            for (x, y) in points.iter().take(expected) {
                d.push(' ');
                d.push_str(&format_path_number(*x));
                d.push(' ');
                d.push_str(&format_path_number(*y));
            }
        }
        if path_fills && d.len() > segment_start {
            if !fill_d.is_empty() {
                fill_d.push(' ');
            }
            fill_d.push_str(d[segment_start..].trim_start());
        }
    }
    if d.is_empty() {
        return None;
    }
    let fill_d = if stroke_only || fill_d == d {
        None
    } else {
        Some(fill_d)
    };
    Some(CustomPath {
        width: base_width.max(1.0),
        height: base_height.max(1.0),
        d,
        stroke_only,
        fill_d,
    })
}

fn hidden_attribute(node: Node<'_, '_>) -> bool {
    node.attribute("hidden")
        .is_some_and(|value| value == "1" || value == "true")
}

fn group_xfrm<'a>(group: Node<'a, 'a>) -> Option<Node<'a, 'a>> {
    group
        .children()
        .find(|node| node.has_tag_name("grpSpPr"))?
        .children()
        .find(|node| node.has_tag_name("xfrm"))
}

fn xfrm_value(xfrm: Option<Node<'_, '_>>, tag: &str, attribute: &str) -> Option<f64> {
    xfrm?
        .children()
        .find(|node| node.has_tag_name(tag))?
        .attribute(attribute)?
        .parse::<f64>()
        .ok()
}

/// Flatten a grpSp into per-child visuals. `group_box` is the group's frame
/// as (x, y, width, height) in EMU relative to the anchor's `from` marker;
/// children map from the group's chOff/chExt space onto that box. Child
/// anchors encode the box as offsets within the from cell, which the
/// renderer's marker walk resolves across real row/column sizes.
#[allow(clippy::too_many_arguments)]
fn expand_group(
    group: Node<'_, '_>,
    anchor: &DrawingAnchor,
    group_box: (f64, f64, f64, f64),
    visual_id: &str,
    counter: &mut usize,
    sheet_id: &str,
    colors: &ColorContext,
    drawing_path: &str,
    relationships: &HashMap<String, Relationship>,
    visuals: &mut Vec<VisualObject>,
) -> Result<(), SidecarError> {
    let (box_x, box_y, box_width, box_height) = group_box;
    let xfrm = group_xfrm(group);
    let ch_off_x = xfrm_value(xfrm, "chOff", "x").unwrap_or(0.0);
    let ch_off_y = xfrm_value(xfrm, "chOff", "y").unwrap_or(0.0);
    let ch_ext_x = xfrm_value(xfrm, "chExt", "cx").filter(|value| *value > 0.0);
    let ch_ext_y = xfrm_value(xfrm, "chExt", "cy").filter(|value| *value > 0.0);
    let scale_x = box_width / ch_ext_x.unwrap_or(box_width);
    let scale_y = box_height / ch_ext_y.unwrap_or(box_height);
    for child in group.children() {
        let is_group = child.has_tag_name("grpSp");
        let is_shape = child.has_tag_name("sp") || child.has_tag_name("cxnSp");
        let is_picture = child.has_tag_name("pic");
        let is_frame = child.has_tag_name("graphicFrame");
        if !is_group && !is_shape && !is_picture && !is_frame {
            continue;
        }
        if child
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"))
            .is_some_and(hidden_attribute)
        {
            continue;
        }
        // The first xfrm under the child is its own (spPr or grpSpPr).
        let child_xfrm = child.descendants().find(|node| node.has_tag_name("xfrm"));
        let (Some(off_x), Some(off_y), Some(ext_x), Some(ext_y)) = (
            xfrm_value(child_xfrm, "off", "x"),
            xfrm_value(child_xfrm, "off", "y"),
            xfrm_value(child_xfrm, "ext", "cx"),
            xfrm_value(child_xfrm, "ext", "cy"),
        ) else {
            continue;
        };
        let child_box = (
            box_x + (off_x - ch_off_x) * scale_x,
            box_y + (off_y - ch_off_y) * scale_y,
            ext_x * scale_x,
            ext_y * scale_y,
        );
        if is_group {
            expand_group(
                child,
                anchor,
                child_box,
                visual_id,
                counter,
                sheet_id,
                colors,
                drawing_path,
                relationships,
                visuals,
            )?;
            continue;
        }
        let child_anchor = DrawingAnchor {
            from_row: anchor.from_row,
            from_column: anchor.from_column,
            from_row_offset: anchor.from_row_offset + child_box.1.round() as i64,
            from_column_offset: anchor.from_column_offset + child_box.0.round() as i64,
            to_row: anchor.from_row,
            to_column: anchor.from_column,
            to_row_offset: anchor.from_row_offset + (child_box.1 + child_box.3).round() as i64,
            to_column_offset: anchor.from_column_offset
                + (child_box.0 + child_box.2).round() as i64,
            explicit_to: false,
        };
        *counter += 1;
        let child_id = format!("{visual_id}-{counter}");
        let child_name = child
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"))
            .and_then(|node| node.attribute("name"))
            .map(ToOwned::to_owned);
        if is_shape {
            visuals.push(shape_visual(
                child,
                child_anchor,
                child_id,
                sheet_id,
                child_name,
                colors,
                drawing_path,
                relationships,
                None,
            ));
            continue;
        }
        if is_frame {
            // Chart data is backfilled by read_drawing (reading the part
            // needs the archive, which this expansion deliberately avoids).
            let Some(chart_path) = child
                .descendants()
                .find(|node| node.has_tag_name("chart"))
                .and_then(relationship_id)
                .and_then(|id| relationships.get(&id))
                .map(|relationship| resolve_part_target(drawing_path, &relationship.target))
                .transpose()?
            else {
                continue;
            };
            visuals.push(VisualObject {
                id: child_id,
                sheet_id: sheet_id.to_owned(),
                kind: "chart".into(),
                anchor: child_anchor,
                chart: None,
                chart_path: Some(chart_path),
                media_path: None,
                media_type: None,
                opacity: None,
                fill_media_path: None,
                fill_media_type: None,
                name: child_name,
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: None,
            });
            continue;
        }
        let Some(blip_node) = child.descendants().find(|node| node.has_tag_name("blip")) else {
            continue;
        };
        let Some(id) = relationship_id(blip_node) else {
            continue;
        };
        let Some(relationship) = relationships.get(&id) else {
            continue;
        };
        if anchor_is_zero_extent(&child_anchor) {
            continue;
        }
        let media_path = resolve_part_target(drawing_path, &relationship.target)?;
        visuals.push(VisualObject {
            id: child_id,
            sheet_id: sheet_id.to_owned(),
            kind: "image".into(),
            anchor: child_anchor,
            chart: None,
            chart_path: None,
            media_type: media_type_for_path(&media_path).map(ToOwned::to_owned),
            media_path: Some(media_path),
            opacity: blip_opacity(blip_node),
            fill_media_path: None,
            fill_media_type: None,
            name: child_name,
            shape_type: None,
            custom_path: None,
            fill_color: None,
            fill_gradient: None,
            line_color: None,
            line_width: None,
            line_dash: None,
            line_cap: None,
            flip_h: false,
            flip_v: false,
            text_color: None,
            text_anchor: None,
            paragraphs: None,
            text: None,
            rotation: None,
            frame_width: None,
            frame_height: None,
            nv_id: None,
            drawing_path: Some(drawing_path.to_owned()),
            drawing_index: None,
        });
    }
    Ok(())
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

/// 3D plot elements fold onto their flat pipeline at parse time (flat 2D
/// projection of the right chart type); the wire only carries the 2D name.
/// stockChart/surfaceChart stay unmapped.
const CHART_TYPE_3D_FOLDS: [(&str, &str); 4] = [
    ("bar3DChart", "barChart"),
    ("line3DChart", "lineChart"),
    ("pie3DChart", "pieChart"),
    ("area3DChart", "areaChart"),
];

/// The 2D wire name of a plot element, folding 3D variants.
fn flat_plot_name(node: Node<'_, '_>) -> Option<&'static str> {
    let name = node.tag_name().name();
    CHART_TYPE_NAMES
        .iter()
        .copied()
        .find(|flat| *flat == name)
        .or_else(|| {
            CHART_TYPE_3D_FOLDS
                .iter()
                .find(|(three_d, _)| *three_d == name)
                .map(|(_, flat)| *flat)
        })
}

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
        .filter(|name| {
            document
                .descendants()
                .any(|node| flat_plot_name(node) == Some(**name))
        })
        .map(|name| (*name).to_owned())
        .collect::<Vec<_>>();
    // Only the chart-level title — axes carry their own c:title deeper down.
    let chart_node = document
        .descendants()
        .find(|node| node.has_tag_name("chart"));
    let title_node = chart_node.and_then(|chart| direct_child(chart, "title"));
    let explicit_title = title_node
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
        .filter(|value| !value.is_empty());
    let auto_title_deleted = chart_node
        .and_then(|chart| direct_child(chart, "autoTitleDeleted"))
        .and_then(|node| node.attribute("val"))
        .is_some_and(|value| value == "1" || value == "true");
    let bar_direction = document
        .descendants()
        .find(|node| node.has_tag_name("barDir"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let series_nodes = document
        .descendants()
        .filter(|node| node.has_tag_name("ser"))
        .collect::<Vec<_>>();
    let sole_series_named = matches!(&series_nodes[..],
        [only] if direct_child(*only, "tx").and_then(first_cached_value).is_some());
    let series = series_nodes
        .iter()
        .enumerate()
        .map(|(index, node)| parse_chart_series(*node, index, colors))
        .collect::<Vec<_>>();
    // Excel's title rules: explicit text wins; a deleted auto title (or no
    // <c:title> at all) means no title; a present-but-empty <c:title> shows
    // the auto title — the sole series' name, else the "Chart Title"
    // placeholder. An empty string means "no title" on the wire.
    let title = match explicit_title {
        Some(text) => text,
        None if auto_title_deleted || title_node.is_none() => String::new(),
        None => match &series[..] {
            [only] if sole_series_named => only.name.clone(),
            _ => "Chart Title".into(),
        },
    };
    let (x_axis, y_axis, secondary_y_axis) = axis_infos(document);
    let title_style = title_node
        .and_then(|node| direct_child(node, "txPr"))
        .and_then(|txpr| {
            txpr.descendants()
                .find(|child| child.has_tag_name("defRPr"))
        })
        .map(|def| ChartTitleStyle {
            size: def
                .attribute("sz")
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 100.0),
            bold: def
                .attribute("b")
                .map(|value| value == "1" || value == "true"),
            color: drawing_fill_color(def, colors),
        });
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
        x_axis,
        y_axis,
        scatter_style: document
            .descendants()
            .find(|node| node.has_tag_name("scatterStyle"))
            .and_then(|node| node.attribute("val"))
            .map(ToOwned::to_owned),
        // Direct child only: c:ser/c:marker (a symbol container) must not
        // register as the plot flag. CT_Boolean: a bare <c:marker/> is true.
        line_markers: document
            .descendants()
            .find(|node| node.has_tag_name("lineChart") || node.has_tag_name("line3DChart"))
            .and_then(|plot| direct_child(plot, "marker"))
            .map(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        secondary_y_axis,
        title_style,
        series,
    }
}

/// All plot axes keyed by side: axPos b/t → X, l/r → Y. Falls back to the
/// element kind (catAx → X, valAx → Y) when axPos is missing.
fn axis_infos(document: &Document<'_>) -> (Option<AxisInfo>, Option<AxisInfo>, Option<AxisInfo>) {
    let mut x_axis = None;
    let mut left_axes: Vec<AxisInfo> = Vec::new();
    // (info, is value axis) — only value axes qualify as the secondary scale.
    let mut right_axes: Vec<(AxisInfo, bool)> = Vec::new();
    for axis in document.descendants().filter(|node| {
        ["catAx", "dateAx", "valAx"]
            .iter()
            .any(|name| node.has_tag_name(*name))
    }) {
        let scaling = direct_child(axis, "scaling");
        let bound = |name: &str| {
            scaling
                .and_then(|node| direct_child(node, name))
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<f64>().ok())
        };
        let info = AxisInfo {
            // Rich text, else the cell-linked strCache <c:v>; a truly empty
            // <c:title> is Excel's auto axis title — the "Axis Title"
            // placeholder.
            title: direct_child(axis, "title").map(|node| {
                let rich = node
                    .descendants()
                    .filter(|child| child.has_tag_name("t"))
                    .filter_map(|child| child.text())
                    .collect::<String>();
                if !rich.is_empty() {
                    return rich;
                }
                let cached = node
                    .descendants()
                    .filter(|child| child.has_tag_name("v"))
                    .filter_map(|child| child.text())
                    .collect::<String>();
                if cached.is_empty() {
                    "Axis Title".into()
                } else {
                    cached
                }
            }),
            min: bound("min"),
            max: bound("max"),
            major_unit: direct_child(axis, "majorUnit")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<f64>().ok()),
            num_fmt: direct_child(axis, "numFmt")
                .and_then(|node| node.attribute("formatCode"))
                .filter(|code| !code.is_empty() && *code != "General")
                .map(ToOwned::to_owned),
            major_gridlines: direct_child(axis, "majorGridlines").is_some(),
            // CT_Boolean: a bare <c:delete/> means true.
            hidden: direct_child(axis, "delete")
                .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
            reversed: scaling
                .and_then(|node| direct_child(node, "orientation"))
                .and_then(|node| node.attribute("val"))
                == Some("maxMin"),
        };
        let position = direct_child(axis, "axPos").and_then(|node| node.attribute("val"));
        let is_x = match position {
            Some("b") | Some("t") => true,
            Some("l") | Some("r") => false,
            _ => axis.has_tag_name("catAx") || axis.has_tag_name("dateAx"),
        };
        if is_x {
            if x_axis.is_none() {
                x_axis = Some(info);
            }
        } else if position == Some("l") {
            left_axes.push(info);
        } else {
            right_axes.push((info, axis.has_tag_name("valAx")));
        }
    }
    // The left axis is the primary scale regardless of document order; the
    // secondary scale is the first remaining right VALUE axis.
    let mut right_values = right_axes
        .into_iter()
        .filter(|(_, is_value)| *is_value)
        .map(|(info, _)| info);
    let (y_axis, secondary_y_axis) = match left_axes.into_iter().next() {
        Some(left) => (Some(left), right_values.next()),
        None => (right_values.next(), right_values.next()),
    };
    (x_axis, y_axis, secondary_y_axis)
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
    let plot = document
        .descendants()
        .find(|node| node.has_tag_name(plot))?;
    direct_child(plot, name)
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok())
}

fn legend_position(document: &Document<'_>) -> String {
    let Some(legend) = document
        .descendants()
        .find(|node| node.has_tag_name("legend"))
    else {
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

/// The dLbls node the single-mode metadata reads. Plot-level wins, unless it
/// resolves to no labels while the first series' own dLbls shows some —
/// per-series dLbls override the plot default in Excel, so an all-zero plot
/// element must not hide labels a series switched on.
fn data_labels_node<'a>(document: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    let plot_labels = document
        .descendants()
        .find(|node| flat_plot_name(*node).is_some())
        .and_then(|plot| direct_child(plot, "dLbls"));
    let series_labels = document
        .descendants()
        .find(|node| node.has_tag_name("ser"))
        .and_then(|series| direct_child(series, "dLbls"));
    match (plot_labels, series_labels) {
        (Some(plot), Some(series))
            if data_labels_mode(plot) == "none" && data_labels_mode(series) != "none" =>
        {
            Some(series)
        }
        (Some(plot), _) => Some(plot),
        (None, series) => series,
    }
}

fn data_labels_mode(labels: Node<'_, '_>) -> &'static str {
    let shown = |name: &str| {
        direct_child(labels, name)
            .and_then(|node| node.attribute("val"))
            .is_some_and(|value| value == "1" || value == "true")
    };
    if shown("delete") {
        return "none";
    }
    match (shown("showCatName"), shown("showVal"), shown("showPercent")) {
        (true, true, true) => "category-value-percent",
        (true, _, true) => "category-percent",
        (_, _, true) => "percent",
        (_, true, _) => "value",
        _ => "none",
    }
}

fn data_labels(document: &Document<'_>) -> Option<String> {
    Some(data_labels_mode(data_labels_node(document)?).into())
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
        matches!(
            flat_plot_name(*node),
            Some("barChart") | Some("areaChart") | Some("lineChart")
        )
    })?;
    direct_child(plot, "grouping")
        .and_then(|node| node.attribute("val"))
        .filter(|value| {
            matches!(
                *value,
                "clustered" | "stacked" | "percentStacked" | "standard"
            )
        })
        .map(ToOwned::to_owned)
}

fn parse_chart_series(series: Node<'_, '_>, index: usize, colors: &ColorContext) -> ChartSeries {
    // Unnamed series get Excel's global Series1..N numbering. A cell-linked
    // name without a strCache keeps its reference for renderer-side lookup.
    let tx = direct_child(series, "tx");
    let cached_name = tx.and_then(first_cached_value);
    let name_ref = match &cached_name {
        Some(_) => None,
        None => tx.and_then(formula_ref),
    };
    let name = cached_name.unwrap_or_else(|| format!("Series{}", index + 1));
    // Explicit series fill/line color, else the theme accent cycle Excel uses
    // for automatic chart colors — keyed by c:idx, not document position.
    let accent_index = direct_child(series, "idx")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(index);
    let color = direct_child(series, "spPr")
        .and_then(|sppr| drawing_fill_color(sppr, colors))
        .or_else(|| theme_accent(colors, accent_index % 6 + 1).map(|base| tint_to_hex(base, 0.0)));
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
    let line = direct_child(series, "spPr")
        .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("ln")));
    let line_color = line.and_then(|ln| {
        if ln.children().any(|node| node.has_tag_name("noFill")) {
            return Some("none".into());
        }
        drawing_fill_color(ln, colors)
    });
    let line_width = line
        .and_then(|ln| ln.attribute("w"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|emu| emu / 12700.0 * (96.0 / 72.0));
    let smooth = direct_child(series, "smooth")
        .and_then(|node| node.attribute("val"))
        .map(|value| value == "1" || value == "true");
    let marker = direct_child(series, "marker")
        .and_then(|node| direct_child(node, "symbol"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    ChartSeries {
        name,
        name_ref,
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
        line_color,
        line_width,
        smooth,
        marker,
        category_groups: category_node.and_then(category_groups),
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

/// Both markers coincide — a zero span on both axes. (The no-ext fallbacks
/// in parse_anchor always produce a real span, so they never match.)
fn anchor_is_zero_extent(anchor: &DrawingAnchor) -> bool {
    anchor.from_row == anchor.to_row
        && anchor.from_row_offset == anchor.to_row_offset
        && anchor.from_column == anchor.to_column
        && anchor.from_column_offset == anchor.to_column_offset
}

/// a:blip/a:alphaModFix amt (per-100000) → 0..1; None when absent or opaque.
fn blip_opacity(blip_node: Node<'_, '_>) -> Option<f64> {
    let amt = blip_node
        .children()
        .find(|node| node.has_tag_name("alphaModFix"))?
        .attribute("amt")?
        .parse::<f64>()
        .ok()?;
    let opacity = (amt / 100_000.0).clamp(0.0, 1.0);
    (opacity < 1.0).then_some(opacity)
}

fn parse_anchor(anchor: Node<'_, '_>) -> Option<DrawingAnchor> {
    let Some(from) = direct_child(anchor, "from") else {
        // absoluteAnchor: xdr:pos + xdr:ext in sheet EMU. Encode both corners
        // as offsets from cell (0,0) — the renderer's marker walk carries
        // offsets across real row/column sizes.
        let pos = direct_child(anchor, "pos")?;
        let ext = direct_child(anchor, "ext")?;
        let coordinate =
            |node: Node<'_, '_>, attribute: &str| node.attribute(attribute)?.parse::<i64>().ok();
        let x = coordinate(pos, "x")?;
        let y = coordinate(pos, "y")?;
        let cx = coordinate(ext, "cx")?;
        let cy = coordinate(ext, "cy")?;
        return Some(DrawingAnchor {
            from_row: 0,
            from_column: 0,
            from_row_offset: y,
            from_column_offset: x,
            to_row: 0,
            to_column: 0,
            to_row_offset: y + cy,
            to_column_offset: x + cx,
            explicit_to: false,
        });
    };
    let from_row = marker_value(from, "row")?;
    let from_column = marker_value(from, "col")?;
    let from_row_offset = marker_signed_value(from, "rowOff").unwrap_or(0);
    let from_column_offset = marker_signed_value(from, "colOff").unwrap_or(0);
    if let Some(to) = direct_child(anchor, "to") {
        return Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: marker_value(to, "row").unwrap_or(from_row + 20),
            to_column: marker_value(to, "col").unwrap_or(from_column + 8),
            to_row_offset: marker_signed_value(to, "rowOff").unwrap_or(0),
            to_column_offset: marker_signed_value(to, "colOff").unwrap_or(0),
            explicit_to: true,
        });
    }
    // oneCellAnchor: the size lives in xdr:ext (EMU). Encode it as offsets
    // within the from cell — the renderer's marker walk handles offsets past
    // the cell edge, so no new anchor fields are needed. (Previously `to`
    // fell back to `from` itself: a zero-size box.)
    let ext = direct_child(anchor, "ext");
    let extent = |attribute: &str| {
        ext.and_then(|node| node.attribute(attribute))
            .and_then(|value| value.parse::<i64>().ok())
    };
    match (extent("cx"), extent("cy")) {
        (Some(cx), Some(cy)) => Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: from_row,
            to_column: from_column,
            to_row_offset: from_row_offset + cy,
            to_column_offset: from_column_offset + cx,
            explicit_to: false,
        }),
        _ => Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: from_row + 20,
            to_column: from_column + 8,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        }),
    }
}

fn parse_font(font: Node<'_, '_>, colors: &ColorContext) -> FontStyle {
    let color_node = font.children().find(|node| node.has_tag_name("color"));
    let color = color_node.and_then(|node| parse_color(node, colors));
    let (color_theme, color_tint) = color_node
        .filter(|_| color.is_some())
        .map(|node| theme_provenance(node, colors))
        .unwrap_or((None, None));
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
        bold: font
            .children()
            .find(|node| node.has_tag_name("b"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        italic: font
            .children()
            .find(|node| node.has_tag_name("i"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        underline: font
            .children()
            .find(|node| node.has_tag_name("u"))
            .is_some_and(|node| node.attribute("val") != Some("none")),
        strikethrough: font
            .children()
            .find(|node| node.has_tag_name("strike"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        color,
        color_theme,
        color_tint,
        scheme: font
            .children()
            .find(|node| node.has_tag_name("scheme"))
            .and_then(|node| node.attribute("val"))
            .filter(|value| *value == "major" || *value == "minor")
            .map(ToOwned::to_owned),
    }
}

/// Theme slot + tint of a color node, when the color resolves through the
/// theme palette (mirrors resolve_color: a resolvable theme slot wins over
/// rgb/indexed).
fn theme_provenance(node: Node<'_, '_>, colors: &ColorContext) -> (Option<usize>, Option<f64>) {
    let theme = node
        .attribute("theme")
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|index| colors.theme.get(*index).is_some());
    if theme.is_none() {
        return (None, None);
    }
    // Clamped to the wire schema's [-1, 1]; apply_tint saturates the same
    // way, so a clamped tint reproduces the baked color exactly.
    let tint = node
        .attribute("tint")
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .map(|value| value.clamp(-1.0, 1.0))
        .filter(|value| *value != 0.0);
    (theme, tint)
}

#[derive(Clone, Default)]
struct FillInfo {
    color: Option<String>,
    theme: Option<usize>,
    tint: Option<f64>,
}

fn parse_fill(fill: Node<'_, '_>, colors: &ColorContext) -> FillInfo {
    let Some(pattern) = fill
        .children()
        .find(|node| node.has_tag_name("patternFill"))
    else {
        return parse_gradient_fill(fill, colors);
    };
    let pattern_type = pattern.attribute("patternType");
    if pattern_type == Some("none") {
        return FillInfo::default();
    }
    let fg_node = pattern.children().find(|node| node.has_tag_name("fgColor"));
    let bg_node = pattern.children().find(|node| node.has_tag_name("bgColor"));
    let foreground = fg_node.and_then(|node| parse_color(node, colors));
    let background = bg_node.and_then(|node| parse_color(node, colors));
    // Textured patterns (gray125, stripes, …) render as the per-channel blend
    // of both colors — the closest flat-color approximation of the texture.
    // Blends carry no single theme provenance.
    if pattern_type.is_some_and(|value| value != "solid") {
        if let (Some(fg), Some(bg)) = (foreground.as_deref(), background.as_deref()) {
            if let Some(mixed) = mix_hex(fg, bg) {
                return FillInfo {
                    color: Some(mixed),
                    ..FillInfo::default()
                };
            }
        }
    }
    let (color, node) = if foreground.is_some() {
        (foreground, fg_node)
    } else {
        (background, bg_node)
    };
    let (theme, tint) = node
        .filter(|_| color.is_some())
        .map(|node| theme_provenance(node, colors))
        .unwrap_or((None, None));
    FillInfo { color, theme, tint }
}

/// The cell style model carries one flat color, so a gradientFill is
/// approximated by the mid-gradient blend of its outermost stops. The blend
/// carries no single theme provenance.
fn parse_gradient_fill(fill: Node<'_, '_>, colors: &ColorContext) -> FillInfo {
    let Some(gradient) = fill
        .children()
        .find(|node| node.has_tag_name("gradientFill"))
    else {
        return FillInfo::default();
    };
    let mut stops = gradient
        .children()
        .filter(|node| node.has_tag_name("stop"))
        .filter_map(|stop| {
            let position = stop.attribute("position")?.parse::<f64>().ok()?;
            let color = stop
                .children()
                .find(|node| node.has_tag_name("color"))
                .and_then(|node| parse_color(node, colors))?;
            Some((position, color))
        })
        .collect::<Vec<_>>();
    // Stops are not guaranteed document-ordered.
    stops.sort_by(|a, b| a.0.total_cmp(&b.0));
    let (Some((_, first)), Some((_, last))) = (stops.first(), stops.last()) else {
        return FillInfo::default();
    };
    FillInfo {
        color: mix_hex(first, last),
        ..FillInfo::default()
    }
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
        vertical: edge("vertical"),
        horizontal: edge("horizontal"),
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
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
    "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
    "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
    "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
    "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
    "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
    "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333", "000000", "FFFFFF",
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
    // A resolvable theme slot wins over rgb: Excel treats rgb as a cached
    // copy of the theme color, and some producers bake a wrong cache
    // (tdf113271 writes theme="1" rgb="FFFFFF" for black dk1 text).
    if let Some(base) = theme
        .and_then(|value| value.parse::<usize>().ok())
        .and_then(|index| colors.theme.get(index).copied())
    {
        let tint = tint
            .and_then(|value| value.parse::<f64>().ok())
            .unwrap_or(0.0);
        let (red, green, blue) = apply_tint(base, tint);
        return Some(format!("#{red:02X}{green:02X}{blue:02X}"));
    }
    if let Some(rgb) = rgb {
        let value = if rgb.len() == 8 { &rgb[2..] } else { rgb };
        return Some(format!("#{value}"));
    }
    let index = indexed?.parse::<usize>().ok()?;
    // 64/65 are the fixed system window text/background slots — producers
    // that override the palette still expect the system colors there.
    if index < 64 {
        if let Some(value) = colors.indexed.get(index) {
            return Some(format!("#{value}"));
        }
    }
    INDEXED_COLORS.get(index).map(|value| format!("#{value}"))
}

/// styles.xml `<colors><indexedColors>` — a legacy-palette override written
/// by workbooks converted from .xls. Entries are ARGB ("00RRGGBB").
pub fn read_indexed_palette(
    archive: &mut ZipArchive<File>,
    colors: &mut ColorContext,
) -> Result<(), SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/styles.xml")? else {
        return Ok(());
    };
    let document = parse_document(&xml, "styles.xml")?;
    let Some(list) = document
        .descendants()
        .find(|node| node.has_tag_name("indexedColors"))
    else {
        return Ok(());
    };
    colors.indexed = list
        .children()
        .filter(|node| node.has_tag_name("rgbColor"))
        .filter_map(|node| node.attribute("rgb"))
        .map(|value| {
            let hex = if value.len() == 8 { &value[2..] } else { value };
            hex.to_owned()
        })
        .collect();
    Ok(())
}

/// Theme accent color (1-6) as rgb, if the palette was loaded.
pub fn theme_accent(colors: &ColorContext, accent: usize) -> Option<(u8, u8, u8)> {
    // Effective palette order: [lt1, dk1, lt2, dk2, accent1-6, ...]
    colors.theme.get(3 + accent).copied()
}

/// Theme dk1 (neutral text) color as rgb, if the palette was loaded.
pub fn theme_dark1(colors: &ColorContext) -> Option<(u8, u8, u8)> {
    colors.theme.get(1).copied()
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
    let fill_styles = document
        .descendants()
        .find(|node| node.has_tag_name("fillStyleLst"))
        .map(|list| {
            list.children()
                .filter(|node| node.is_element())
                .map(|node| parse_theme_gradient(node))
                .collect()
        })
        .unwrap_or_default();
    Ok(ColorContext {
        theme,
        fill_styles,
        indexed: Vec::new(),
    })
}

/// A fillStyleLst gradFill entry as data (the theme document does not
/// outlive the palette); non-gradient entries map to None.
fn parse_theme_gradient(node: Node<'_, '_>) -> Option<ThemeGradient> {
    if !node.has_tag_name("gradFill") {
        return None;
    }
    let stops = node
        .descendants()
        .filter(|child| child.has_tag_name("gs"))
        .filter_map(|gs| {
            let position = gs.attribute("pos")?.parse::<f64>().ok()? / 100_000.0;
            let color = gs
                .children()
                .find(|child| child.has_tag_name("schemeClr") || child.has_tag_name("srgbClr"))?;
            let modifiers = color
                .children()
                .filter(|child| child.is_element())
                .filter_map(|child| {
                    Some((
                        child.tag_name().name().to_owned(),
                        child.attribute("val")?.parse::<f64>().ok()? / 100_000.0,
                    ))
                })
                .collect();
            Some(ThemeGradientStop {
                position,
                modifiers,
            })
        })
        .collect::<Vec<_>>();
    if stops.len() < 2 {
        return None;
    }
    let angle = node
        .descendants()
        .find(|child| child.has_tag_name("lin"))
        .and_then(|lin| lin.attribute("ang"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| value / 60_000.0)
        .unwrap_or(90.0);
    Some(ThemeGradient { stops, angle })
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
    pub style_name: Option<String>,
    pub first_data_row: usize,
    pub row_grand_totals: bool,
    pub show_row_stripes: bool,
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
        let Some(location) = document
            .descendants()
            .find(|node| node.has_tag_name("location"))
        else {
            continue;
        };
        let Some(output_ref) = location.attribute("ref") else {
            continue;
        };
        let first_data_row = location
            .attribute("firstDataRow")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(1);
        let root = document.root_element();
        let row_grand_totals = root
            .attribute("rowGrandTotals")
            .map(|value| value == "1" || value == "true")
            .unwrap_or(true);
        let style_info = document
            .descendants()
            .find(|node| node.has_tag_name("pivotTableStyleInfo"));
        let style_name = style_info
            .and_then(|node| node.attribute("name"))
            .map(str::to_owned);
        let show_row_stripes = style_info
            .and_then(|node| node.attribute("showRowStripes"))
            .is_some_and(|value| value == "1" || value == "true");
        let cache_path = read_relationships(archive, &pivot_path)?
            .values()
            .find(|part| part.relationship_type.ends_with("/pivotCacheDefinition"))
            .map(|part| resolve_part_target(&pivot_path, &part.target))
            .transpose()?;
        infos.push(PivotPartInfo {
            path: pivot_path,
            cache_path,
            output_ref: output_ref.to_owned(),
            style_name,
            first_data_row,
            row_grand_totals,
            show_row_stripes,
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

pub(crate) fn read_relationships(
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

pub(crate) fn resolve_part_target(source_path: &str, target: &str) -> Result<String, SidecarError> {
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

pub(crate) fn read_optional_xml(
    archive: &mut ZipArchive<File>,
    path: &str,
) -> Result<Option<String>, SidecarError> {
    let Ok(mut entry) = crate::zip_entry(archive, path) else {
        return Ok(None);
    };
    let mut xml = String::new();
    entry.read_to_string(&mut xml)?;
    Ok(Some(xml))
}

pub(crate) fn parse_document<'a>(xml: &'a str, path: &str) -> Result<Document<'a>, SidecarError> {
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
    // multiLvlStrCache: sweeping every c:pt would flatten L levels into one
    // L×N array. The first c:lvl is the innermost (detail) level per OOXML.
    let scope = node
        .descendants()
        .find(|child| child.has_tag_name("lvl"))
        .unwrap_or(node);
    // Document order, compacted: the numeric value path drops blanks the same
    // way, so padding sparse c:pt/@idx gaps here would misalign categories
    // against values. Aligning both caches by idx union needs a wire change.
    scope
        .descendants()
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

/// First outer level of a multiLvlStrCache: each pt idx marks a group start
/// in cache-idx space; the span runs until the next pt (or ptCount). Spans
/// are remapped onto the document-order positions `cached_values` emits, so
/// sparse caches cannot misalign groups against categories. Ambiguity —
/// an outer pt without an idx, or a group whose members are not contiguous
/// in the emitted order — skips the groups entirely.
fn category_groups(node: Node<'_, '_>) -> Option<Vec<CategoryGroup>> {
    let cache = node
        .descendants()
        .find(|child| child.has_tag_name("multiLvlStrCache"))?;
    let mut levels = cache.children().filter(|child| child.has_tag_name("lvl"));
    let inner = levels.next()?;
    let outer = levels.next()?;
    let inner_idx: Vec<usize> = inner
        .children()
        .filter(|child| child.has_tag_name("pt"))
        .enumerate()
        .map(|(position, point)| numeric_attribute(point, "idx").unwrap_or(position))
        .collect();
    if inner_idx.is_empty() {
        return None;
    }
    let outer_points: Vec<Node<'_, '_>> = outer
        .children()
        .filter(|child| child.has_tag_name("pt"))
        .collect();
    let mut starts: Vec<(usize, String)> = outer_points
        .iter()
        .filter_map(|point| {
            Some((
                numeric_attribute(*point, "idx")?,
                direct_child(*point, "v")?.text()?.to_owned(),
            ))
        })
        .collect();
    if starts.is_empty() || starts.len() != outer_points.len() {
        return None;
    }
    starts.sort_by_key(|(idx, _)| *idx);
    let point_count = direct_child(cache, "ptCount")
        .and_then(|count| numeric_attribute(count, "val"))
        .unwrap_or_else(|| inner_idx.iter().max().map_or(0, |max| max + 1));
    // A group's members are the emitted positions whose idx falls in
    // [start, next); document order need not be idx order, so require the
    // member positions to be contiguous before expressing them as a span.
    let mut groups = Vec::new();
    for (position, (idx, label)) in starts.iter().enumerate() {
        let next = starts
            .get(position + 1)
            .map_or_else(|| point_count.max(*idx), |(next_idx, _)| *next_idx);
        let members: Vec<usize> = inner_idx
            .iter()
            .enumerate()
            .filter(|(_, inner)| (*idx..next).contains(*inner))
            .map(|(emitted, _)| emitted)
            .collect();
        let (Some(&start), Some(&last)) = (members.first(), members.last()) else {
            continue;
        };
        let end = last + 1;
        if end - start != members.len() {
            return None;
        }
        groups.push(CategoryGroup {
            label: label.clone(),
            start,
            end,
        });
    }
    (!groups.is_empty()).then_some(groups)
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
        "webp" => Some("image/webp"),
        "svg" => Some("image/svg+xml"),
        // GDI metafiles: the renderer rasterizes these to PNG before display.
        "emf" => Some("image/x-emf"),
        "wmf" => Some("image/x-wmf"),
        "emz" => Some("image/x-emz"),
        "wmz" => Some("image/x-wmz"),
        _ => None,
    }
}

/// Implicit number formats, ECMA-376 §18.8.30. Ids 23-26 are undocumented
/// and stay unresolved; everything else is mapped so that a builtin id never
/// falls back to General (which would surface raw date serials, #numFmt58).
///
/// Locale-reserved ranges carry no formatCode in styles.xml — the reader is
/// expected to resolve them for its current locale:
///  - 27-36 / 50-58: locale-dependent date/time formats. The same id means a
///    different pattern per locale and the file does not record which. CJK
///    locales use the zh-CN-compatible table below; other locales use their
///    local full short-date pattern so a CJK month/day format cannot leak into
///    a European workbook and discard its year. The zh AM/PM token (U+4E0A/U+4E0B
///    U+5348) is not understood by the renderer's numfmt, so 34/35 render as
///    24-hour. Ids 55/56 are dates: Excel renders them as the OS short date
///    (verified against a ja-authored workbook where the zh time mapping
///    turned a month header into "0\u{65f6}00\u{5206}"), so they follow the
///    host short-date pattern like 14/22 and fall back to a plain date.
///    Escapes: U+5E74 year, U+6708 month, U+65E5 day, U+65F6 hour, U+5206
///    minute, U+79D2 second.
///  - 41-44: accounting formats; 42/44 use "$" as the symbol is likewise
///    locale-defined and unrecorded.
///  - 59-81: th-TH; numfmt has no Thai digit/era tokens, so these map to
///    Arabic-digit equivalents (Buddhist-era years render as Gregorian).
/// Ids 14/22 are the locale-reactive short-date builtins; when the host
/// supplies the OS short-date pattern they follow it (explicit formatCode
/// entries still win at the call site).
fn short_date_number_format(id: u32, short_date: Option<&str>) -> Option<String> {
    let short_date = short_date?;
    match id {
        14 | 55 | 56 => Some(short_date.to_owned()),
        22 => Some(format!("{short_date} h:mm")),
        _ => None,
    }
}

fn builtin_number_format(id: u32, locale: &str) -> Option<&'static str> {
    if matches!(
        id,
        27 | 28 | 29 | 30 | 31 | 36 | 50 | 51 | 52 | 53 | 54 | 55 | 56 | 57 | 58
    ) && !matches!(locale, "zh" | "zh-TW" | "ja" | "ko")
    {
        return Some(locale_short_date_format(locale));
    }
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some(r##""$"#,##0_);("$"#,##0)"##),
        6 => Some(r##""$"#,##0_);[Red]("$"#,##0)"##),
        7 => Some(r##""$"#,##0.00_);("$"#,##0.00)"##),
        8 => Some(r##""$"#,##0.00_);[Red]("$"#,##0.00)"##),
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
        // ECMA-376 prints 20/21 as h:mm(:ss), but Excel renders these
        // builtins with a leading zero on the hour (09:30, matching
        // LibreOffice's HH:MM mapping) — verified against Excel output.
        20 => Some("hh:mm"),
        21 => Some("hh:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        27 | 36 | 50 | 52 | 57 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\""),
        28 | 29 | 51 | 53 | 54 | 58 => Some("m\"\u{6708}\"d\"\u{65e5}\""),
        30 => Some("m-d-yy"),
        31 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\""),
        32 | 34 => Some("h\"\u{65f6}\"mm\"\u{5206}\""),
        33 | 35 => Some("h\"\u{65f6}\"mm\"\u{5206}\"ss\"\u{79d2}\""),
        // CJK fallback when the host supplied no OS short-date pattern.
        55 | 56 => Some("yyyy/m/d"),
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

fn locale_short_date_format(locale: &str) -> &'static str {
    match locale {
        "en" => "m/d/yyyy",
        "de" | "pl" | "ru" => "d.m.yyyy",
        "nl" => "d-m-yyyy",
        _ => "d/m/yyyy",
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

    fn custom_geometry(paths: &str) -> Option<CustomPath> {
        let xml = format!(
            r#"<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:spPr><a:custGeom><a:avLst/><a:pathLst>{paths}</a:pathLst></a:custGeom></xdr:spPr></xdr:sp>"#
        );
        let document = Document::parse(&xml).unwrap();
        parse_custom_geometry(document.root_element())
    }

    #[test]
    fn custom_geometry_maps_svg_commands() {
        let path = custom_geometry(
            r#"<a:path w="715645" h="5080"><a:moveTo><a:pt x="715060" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="4572"/></a:lnTo><a:close/></a:path>"#,
        )
        .unwrap();
        assert_eq!(path.d, "M 715060 0 L 0 0 L 0 4572 Z");
        assert_eq!(path.width, 715645.0);
        assert_eq!(path.height, 5080.0);
        assert!(!path.stroke_only);
    }

    #[test]
    fn custom_geometry_mixed_fills_expose_fillable_subpaths_only() {
        let path = custom_geometry(
            r#"<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path><a:path w="100" h="100" fill="none"><a:moveTo><a:pt x="10" y="10"/></a:moveTo><a:lnTo><a:pt x="90" y="10"/></a:lnTo></a:path>"#,
        )
        .unwrap();
        assert!(!path.stroke_only);
        assert_eq!(path.d, "M 0 0 L 100 100 Z M 10 10 L 90 10");
        assert_eq!(path.fill_d.as_deref(), Some("M 0 0 L 100 100 Z"));
    }

    #[test]
    fn custom_geometry_uniform_fills_have_no_fill_d() {
        let filled = custom_geometry(
            r#"<a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:close/></a:path>"#,
        )
        .unwrap();
        assert_eq!(filled.fill_d, None);
        let stroked = custom_geometry(
            r#"<a:path w="10" h="10" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo></a:path>"#,
        )
        .unwrap();
        assert!(stroked.stroke_only);
        assert_eq!(stroked.fill_d, None);
    }

    #[test]
    fn custom_geometry_stroke_only_open_path_with_degenerate_height() {
        let path = custom_geometry(
            r#"<a:path w="233679" h="0" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="233171" y="0"/></a:lnTo></a:path>"#,
        )
        .unwrap();
        assert_eq!(path.d, "M 0 0 L 233171 0");
        assert_eq!(path.height, 1.0);
        assert!(path.stroke_only);
    }

    #[test]
    fn custom_geometry_scales_secondary_paths_into_the_first() {
        let path = custom_geometry(
            r#"<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path><a:path w="200" h="50"><a:lnTo><a:pt x="200" y="50"/></a:lnTo></a:path>"#,
        )
        .unwrap();
        assert_eq!(path.d, "M 0 0 L 100 100");
    }

    #[test]
    fn custom_geometry_rejects_arcs() {
        assert!(
            custom_geometry(
                r#"<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="10" hR="10" stAng="0" swAng="5400000"/></a:path>"#,
            )
            .is_none()
        );
    }

    fn fill_info(body: &str, colors: &ColorContext) -> FillInfo {
        let xml = format!("<fill>{body}</fill>");
        let document = Document::parse(&xml).unwrap();
        parse_fill(document.root_element(), colors)
    }

    #[test]
    fn gradient_fill_blends_outermost_stops() {
        let colors = ColorContext {
            theme: vec![
                (0xFF, 0xFF, 0xFF),
                (0x00, 0x00, 0x00),
                (0x00, 0x00, 0x00),
                (0x00, 0x00, 0x00),
                (0x44, 0x72, 0xC4),
            ],
            ..ColorContext::default()
        };
        // Stops deliberately out of document order; theme stops resolve
        // through the palette before blending.
        let fill = fill_info(
            r#"<gradientFill degree="270"><stop position="1"><color theme="4"/></stop><stop position="0.5"><color rgb="FFFF0000"/></stop><stop position="0"><color theme="0"/></stop></gradientFill>"#,
            &colors,
        );
        assert_eq!(fill.color.as_deref(), Some("#A1B8E1"));
        assert_eq!(fill.theme, None);
        assert_eq!(fill.tint, None);
    }

    #[test]
    fn gradient_fill_without_resolvable_stops_stays_unfilled() {
        let fill = fill_info(
            r#"<gradientFill><stop position="0"/></gradientFill>"#,
            &ColorContext::default(),
        );
        assert_eq!(fill.color, None);
    }

    #[test]
    fn indexed_palette_override_wins_below_system_slots() {
        let colors = ColorContext {
            indexed: vec!["112233".into(); 20],
            ..ColorContext::default()
        };
        assert_eq!(
            resolve_color(None, Some("8"), None, None, &colors),
            Some("#112233".into())
        );
        // Past the override table: builtin legacy palette.
        assert_eq!(
            resolve_color(None, Some("22"), None, None, &colors),
            Some("#C0C0C0".into())
        );
        // 64/65 stay the fixed system slots even when overridden.
        let colors = ColorContext {
            indexed: vec!["112233".into(); 66],
            ..ColorContext::default()
        };
        assert_eq!(
            resolve_color(None, Some("64"), None, None, &colors),
            Some("#000000".into())
        );
    }

    #[test]
    fn media_types_cover_gdi_metafiles() {
        assert_eq!(
            media_type_for_path("xl/media/image1.emf"),
            Some("image/x-emf")
        );
        assert_eq!(
            media_type_for_path("xl/media/image1.WMF"),
            Some("image/x-wmf")
        );
        assert_eq!(
            media_type_for_path("xl/media/image1.emz"),
            Some("image/x-emz")
        );
        assert_eq!(
            media_type_for_path("xl/media/image1.wmz"),
            Some("image/x-wmz")
        );
        assert_eq!(media_type_for_path("xl/media/object1.bin"), None);
    }

    #[test]
    fn media_types_cover_webp() {
        assert_eq!(
            media_type_for_path("xl/media/image.webp"),
            Some("image/webp")
        );
        assert_eq!(
            media_type_for_path("xl/media/IMAGE2.WEBP"),
            Some("image/webp")
        );
    }

    fn parsed_anchor(body: &str) -> DrawingAnchor {
        let xml = format!(
            r#"<xdr:anchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">{body}</xdr:anchor>"#
        );
        let document = Document::parse(&xml).unwrap();
        parse_anchor(document.root_element()).unwrap()
    }

    #[test]
    fn zero_extent_anchor_detection() {
        let marker = |row: usize, col: usize| {
            format!(
                "<xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff>"
            )
        };
        // from == to with zero offsets: Excel shows no picture.
        let degenerate = parsed_anchor(&format!(
            "<xdr:from>{m}</xdr:from><xdr:to>{m}</xdr:to>",
            m = marker(3, 2)
        ));
        assert!(anchor_is_zero_extent(&degenerate));
        let spanning = parsed_anchor(&format!(
            "<xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>",
            marker(3, 2),
            marker(5, 4)
        ));
        assert!(!anchor_is_zero_extent(&spanning));
        // oneCellAnchor with a real ext keeps its span.
        let one_cell = parsed_anchor(&format!(
            r#"<xdr:from>{}</xdr:from><xdr:ext cx="914400" cy="914400"/>"#,
            marker(0, 0)
        ));
        assert!(!anchor_is_zero_extent(&one_cell));
        // oneCellAnchor with a 0x0 ext collapses.
        let one_cell_zero = parsed_anchor(&format!(
            r#"<xdr:from>{}</xdr:from><xdr:ext cx="0" cy="0"/>"#,
            marker(0, 0)
        ));
        assert!(anchor_is_zero_extent(&one_cell_zero));
        // No ext at all: the 20x8-cell fallback frame stays visible.
        let fallback = parsed_anchor(&format!("<xdr:from>{}</xdr:from>", marker(0, 0)));
        assert!(!anchor_is_zero_extent(&fallback));
        // absoluteAnchor with a real ext keeps its span.
        let absolute =
            parsed_anchor(r#"<xdr:pos x="100" y="100"/><xdr:ext cx="914400" cy="914400"/>"#);
        assert!(!anchor_is_zero_extent(&absolute));
        // Only real <xdr:to> markers are flagged for cell-edge clamping;
        // synthesized to markers must keep the walk-past-the-edge encoding.
        assert!(degenerate.explicit_to);
        assert!(spanning.explicit_to);
        assert!(!one_cell.explicit_to);
        assert!(!fallback.explicit_to);
        assert!(!absolute.explicit_to);
    }

    fn opacity_of(blip_body: &str) -> Option<f64> {
        let xml = format!(
            r#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1">{blip_body}</a:blip>"#
        );
        let document = Document::parse(&xml).unwrap();
        blip_opacity(document.root_element())
    }

    #[test]
    fn alpha_mod_fix_becomes_opacity() {
        assert_eq!(opacity_of(r#"<a:alphaModFix amt="20000"/>"#), Some(0.2));
        assert_eq!(opacity_of(r#"<a:alphaModFix amt="100000"/>"#), None);
        assert_eq!(opacity_of(r#"<a:alphaModFix amt="250000"/>"#), None);
        assert_eq!(opacity_of(r#"<a:alphaModFix/>"#), None);
        assert_eq!(opacity_of(""), None);
    }

    #[test]
    fn shape_with_blip_fill_keeps_geometry_and_carries_fill_media() {
        let xml = r#"<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:spPr><a:prstGeom prst="heart"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></xdr:spPr></xdr:sp>"#;
        let document = Document::parse(xml).unwrap();
        let relationships = HashMap::from([(
            "rId1".to_owned(),
            Relationship {
                target: "../media/image1.png".into(),
                relationship_type: String::new(),
            },
        )]);
        let anchor = DrawingAnchor {
            from_row: 0,
            from_column: 0,
            from_row_offset: 0,
            from_column_offset: 0,
            to_row: 10,
            to_column: 5,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        };
        let visual = shape_visual(
            document.root_element(),
            anchor,
            "visual-1".into(),
            "sheet-1",
            None,
            &ColorContext::default(),
            "xl/drawings/drawing1.xml",
            &relationships,
            Some(0),
        );
        assert_eq!(visual.kind, "shape");
        assert_eq!(visual.shape_type.as_deref(), Some("heart"));
        assert_eq!(
            visual.fill_media_path.as_deref(),
            Some("xl/media/image1.png")
        );
        assert_eq!(visual.fill_media_type.as_deref(), Some("image/png"));
    }

    #[test]
    fn numbers_unnamed_series_like_excel() {
        let chart = metadata(
            r#"<c:title/><c:plotArea><c:barChart><c:ser/><c:ser/></c:barChart></c:plotArea>"#,
        );
        assert_eq!(chart.series[0].name, "Series1");
        assert_eq!(chart.series[1].name, "Series2");
        assert_eq!(chart.title, "Chart Title");
    }

    #[test]
    fn sole_named_series_still_becomes_auto_title() {
        let chart = metadata(
            r#"<c:title/><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx></c:ser></c:barChart></c:plotArea>"#,
        );
        assert_eq!(chart.series[0].name, "Revenue");
        assert_eq!(chart.title, "Revenue");
    }

    #[test]
    fn sole_unnamed_series_keeps_placeholder_title() {
        let chart =
            metadata(r#"<c:title/><c:plotArea><c:barChart><c:ser/></c:barChart></c:plotArea>"#);
        assert_eq!(chart.series[0].name, "Series1");
        assert_eq!(chart.title, "Chart Title");
    }

    fn theme_colors() -> ColorContext {
        ColorContext {
            theme: (0..12).map(|slot| (slot as u8, 0x22, 0x33)).collect(),
            fill_styles: Vec::new(),
            indexed: Vec::new(),
        }
    }

    const XDR: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    #[test]
    fn parses_absolute_anchor_as_offsets_from_origin() {
        let xml = format!(
            r#"<xdr:wsDr {XDR}><xdr:absoluteAnchor><xdr:pos x="1440000" y="1080000"/><xdr:ext cx="2880000" cy="720000"/><xdr:sp/></xdr:absoluteAnchor></xdr:wsDr>"#
        );
        let document = Document::parse(&xml).unwrap();
        let anchor_node = document
            .descendants()
            .find(|node| node.has_tag_name("absoluteAnchor"))
            .unwrap();
        let anchor = parse_anchor(anchor_node).unwrap();
        assert_eq!(anchor.from_row, 0);
        assert_eq!(anchor.from_column, 0);
        assert_eq!(anchor.from_column_offset, 1_440_000);
        assert_eq!(anchor.from_row_offset, 1_080_000);
        assert_eq!(anchor.to_column_offset, 4_320_000);
        assert_eq!(anchor.to_row_offset, 1_800_000);
    }

    #[test]
    fn expands_group_children_through_child_space() {
        // Group box 200x100 (EMU), child space 100x50 offset at (10, 20):
        // scale is 2x on both axes.
        let xml = format!(
            r#"<xdr:wsDr {XDR}><xdr:grpSp>
              <xdr:nvGrpSpPr><xdr:cNvPr id="1" name="g"/></xdr:nvGrpSpPr>
              <xdr:grpSpPr><a:xfrm>
                <a:off x="0" y="0"/><a:ext cx="200" cy="100"/>
                <a:chOff x="10" y="20"/><a:chExt cx="100" cy="50"/>
              </a:xfrm></xdr:grpSpPr>
              <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="child"/></xdr:nvSpPr>
                <xdr:spPr><a:xfrm><a:off x="20" y="30"/><a:ext cx="40" cy="10"/></a:xfrm>
                <a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>
              <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="hidden" hidden="1"/></xdr:nvSpPr>
                <xdr:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="5" cy="5"/></a:xfrm>
                <a:prstGeom prst="rect"/></xdr:spPr></xdr:sp>
            </xdr:grpSp></xdr:wsDr>"#
        );
        let document = Document::parse(&xml).unwrap();
        let group = document
            .descendants()
            .find(|node| node.has_tag_name("grpSp"))
            .unwrap();
        let anchor = DrawingAnchor {
            from_row: 3,
            from_column: 2,
            from_row_offset: 1000,
            from_column_offset: 500,
            to_row: 9,
            to_column: 8,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        };
        let mut visuals = Vec::new();
        let mut counter = 0;
        expand_group(
            group,
            &anchor,
            (0.0, 0.0, 200.0, 100.0),
            "visual-1",
            &mut counter,
            "sheet1",
            &ColorContext::default(),
            "xl/drawings/drawing1.xml",
            &HashMap::new(),
            &mut visuals,
        )
        .unwrap();
        assert_eq!(visuals.len(), 1, "hidden child must be skipped");
        let child = &visuals[0];
        assert_eq!(child.id, "visual-1-1");
        assert_eq!(child.name.as_deref(), Some("child"));
        // (20-10)*2 = 20 within the box, plus the group's own from offset.
        assert_eq!(child.anchor.from_column_offset, 500 + 20);
        assert_eq!(child.anchor.from_row_offset, 1000 + 20);
        assert_eq!(child.anchor.to_column_offset, 500 + 20 + 80);
        assert_eq!(child.anchor.to_row_offset, 1000 + 20 + 20);
        assert_eq!(child.anchor.from_row, 3);
        assert_eq!(child.anchor.to_row, 3);
        assert!(
            child.drawing_index.is_none(),
            "group children are read-only"
        );
    }

    #[test]
    fn expands_group_chart_children_with_their_part_path() {
        let xml = format!(
            r#"<xdr:wsDr {XDR} xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:grpSp>
              <xdr:nvGrpSpPr><xdr:cNvPr id="1" name="g"/></xdr:nvGrpSpPr>
              <xdr:grpSpPr><a:xfrm>
                <a:off x="0" y="0"/><a:ext cx="200" cy="100"/>
                <a:chOff x="0" y="0"/><a:chExt cx="200" cy="100"/>
              </a:xfrm></xdr:grpSpPr>
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="chart child"/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="20" y="30"/><a:ext cx="40" cy="10"/></xdr:xfrm>
                <a:graphic><a:graphicData><c:chart r:id="rId7"/></a:graphicData></a:graphic>
              </xdr:graphicFrame>
            </xdr:grpSp></xdr:wsDr>"#
        );
        let document = Document::parse(&xml).unwrap();
        let group = document
            .descendants()
            .find(|node| node.has_tag_name("grpSp"))
            .unwrap();
        let anchor = DrawingAnchor {
            from_row: 0,
            from_column: 0,
            from_row_offset: 0,
            from_column_offset: 0,
            to_row: 5,
            to_column: 5,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        };
        let relationships = HashMap::from([(
            "rId7".to_owned(),
            Relationship {
                target: "../charts/chart1.xml".to_owned(),
                relationship_type: String::new(),
            },
        )]);
        let mut visuals = Vec::new();
        let mut counter = 0;
        expand_group(
            group,
            &anchor,
            (0.0, 0.0, 200.0, 100.0),
            "visual-1",
            &mut counter,
            "sheet1",
            &ColorContext::default(),
            "xl/drawings/drawing1.xml",
            &relationships,
            &mut visuals,
        )
        .unwrap();
        assert_eq!(visuals.len(), 1);
        let child = &visuals[0];
        assert_eq!(child.kind, "chart");
        assert_eq!(child.chart_path.as_deref(), Some("xl/charts/chart1.xml"));
        assert!(child.chart.is_none(), "data is backfilled by read_drawing");
        assert_eq!(child.name.as_deref(), Some("chart child"));
        assert_eq!(child.anchor.from_column_offset, 20);
        assert_eq!(child.anchor.from_row_offset, 30);
        assert_eq!(child.anchor.to_column_offset, 60);
        assert_eq!(child.anchor.to_row_offset, 40);
        assert!(
            child.drawing_index.is_none(),
            "group children are read-only"
        );
    }

    #[test]
    fn rotated_shape_reports_true_frame_extent() {
        let xml = format!(
            r#"<xdr:wsDr {XDR}><xdr:sp>
              <xdr:nvSpPr><xdr:cNvPr id="2" name="r"/></xdr:nvSpPr>
              <xdr:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="3686174" cy="419100"/></a:xfrm>
              <a:prstGeom prst="rect"/></xdr:spPr></xdr:sp></xdr:wsDr>"#
        );
        let document = Document::parse(&xml).unwrap();
        let shape = document
            .descendants()
            .find(|node| node.has_tag_name("sp"))
            .unwrap();
        let anchor = DrawingAnchor {
            from_row: 0,
            from_column: 0,
            from_row_offset: 0,
            from_column_offset: 0,
            to_row: 1,
            to_column: 1,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        };
        let visual = shape_visual(
            shape,
            anchor,
            "visual-1".into(),
            "sheet1",
            None,
            &ColorContext::default(),
            "xl/drawings/drawing1.xml",
            &HashMap::new(),
            Some(0),
        );
        assert_eq!(visual.rotation, Some(45.0));
        assert_eq!(visual.frame_width, Some(3_686_174.0));
        assert_eq!(visual.frame_height, Some(419_100.0));
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
        assert_eq!(
            metadata("<c:plotArea><c:barChart/></c:plotArea>").legend,
            "none"
        );
    }

    #[test]
    fn maps_data_labels_from_plot_or_series() {
        let plot = |labels: &str| {
            format!(
                "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser>{labels}</c:pieChart></c:plotArea>"
            )
        };
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showVal val=\"1\"/></c:dLbls>"))
                .data_labels
                .as_deref(),
            Some("value")
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showPercent val=\"1\"/></c:dLbls>"))
                .data_labels
                .as_deref(),
            Some("percent")
        );
        assert_eq!(
            metadata(&plot(
                "<c:dLbls><c:showCatName val=\"1\"/><c:showPercent val=\"1\"/></c:dLbls>"
            ))
            .data_labels
            .as_deref(),
            Some("category-percent")
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:delete val=\"1\"/></c:dLbls>"))
                .data_labels
                .as_deref(),
            Some("none")
        );
        assert_eq!(
            metadata(&plot("<c:dLbls><c:showVal val=\"0\"/></c:dLbls>"))
                .data_labels
                .as_deref(),
            Some("none")
        );
        assert_eq!(
            metadata(&plot(
                "<c:dLbls><c:showCatName val=\"1\"/><c:showVal val=\"1\"/><c:showPercent val=\"1\"/></c:dLbls>"
            ))
            .data_labels
            .as_deref(),
            Some("category-value-percent")
        );
        // No dLbls anywhere: absent, so renderer defaults may apply.
        assert_eq!(metadata(&plot("")).data_labels, None);
        // Plot-level dLbls missing: fall back to the first series.
        let series_level = "<c:plotArea><c:barChart><c:ser><c:dLbls><c:showVal val=\"1\"/></c:dLbls></c:ser></c:barChart></c:plotArea>";
        assert_eq!(metadata(series_level).data_labels.as_deref(), Some("value"));
    }

    #[test]
    fn series_labels_override_an_all_zero_plot_element() {
        // Excel: per-series dLbls win over the plot-level default, so a plot
        // element with every show* off must not hide the series' labels.
        let both = "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/><c:dLbls>\
            <c:showVal val=\"1\"/><c:showCatName val=\"1\"/><c:showPercent val=\"1\"/>\
            </c:dLbls></c:ser><c:dLbls><c:showVal val=\"0\"/><c:showCatName val=\"0\"/>\
            <c:showPercent val=\"0\"/></c:dLbls></c:pieChart></c:plotArea>";
        assert_eq!(
            metadata(both).data_labels.as_deref(),
            Some("category-value-percent")
        );
        // A plot element that shows labels still wins over the series.
        let plot_wins = "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/><c:dLbls>\
            <c:showPercent val=\"1\"/></c:dLbls></c:ser><c:dLbls><c:showVal val=\"1\"/>\
            </c:dLbls></c:pieChart></c:plotArea>";
        assert_eq!(metadata(plot_wins).data_labels.as_deref(), Some("value"));
    }

    #[test]
    fn uncached_series_name_reference_is_kept_for_lookup() {
        let chart = metadata(
            r#"<c:title/><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>Dashboard!$C$13</c:f></c:strRef></c:tx></c:ser></c:barChart></c:plotArea>"#,
        );
        assert_eq!(chart.series[0].name, "Series1");
        assert_eq!(chart.series[0].name_ref.as_deref(), Some("Dashboard!$C$13"));
        // A cached name needs no reference lookup.
        let cached = metadata(
            r#"<c:title/><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>Dashboard!$C$13</c:f><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx></c:ser></c:barChart></c:plotArea>"#,
        );
        assert_eq!(cached.series[0].name, "Revenue");
        assert_eq!(cached.series[0].name_ref, None);
    }

    #[test]
    fn maps_data_label_position_and_format() {
        let plot = |labels: &str| {
            format!(
                "<c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/></c:ser>{labels}</c:barChart></c:plotArea>"
            )
        };
        for (val, expected) in [
            ("ctr", "center"),
            ("inEnd", "inside-end"),
            ("outEnd", "outside-end"),
        ] {
            let body = plot(&format!("<c:dLbls><c:dLblPos val=\"{val}\"/></c:dLbls>"));
            assert_eq!(
                metadata(&body).data_label_position.as_deref(),
                Some(expected),
                "dLblPos {val}"
            );
        }
        let best_fit = plot("<c:dLbls><c:dLblPos val=\"bestFit\"/></c:dLbls>");
        assert!(metadata(&best_fit).data_label_position.is_none());
        assert!(
            metadata(&plot("<c:dLbls><c:showVal val=\"1\"/></c:dLbls>"))
                .data_label_position
                .is_none()
        );

        let series_level = "<c:plotArea><c:barChart><c:ser><c:dLbls><c:dLblPos val=\"outEnd\"/><c:numFmt formatCode=\"0.0%\"/></c:dLbls></c:ser></c:barChart></c:plotArea>";
        let chart = metadata(series_level);
        assert_eq!(chart.data_label_position.as_deref(), Some("outside-end"));
        assert_eq!(chart.data_label_format.as_deref(), Some("0.0%"));

        let formatted =
            plot("<c:dLbls><c:numFmt formatCode=\"#,##0\" sourceLinked=\"0\"/></c:dLbls>");
        assert_eq!(
            metadata(&formatted).data_label_format.as_deref(),
            Some("#,##0")
        );
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
            format!(
                "<c:title><c:tx><c:rich><a:p><a:r><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx></c:title>"
            )
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
        assert!(
            metadata("<c:plotArea><c:pieChart/></c:plotArea>")
                .grouping
                .is_none()
        );
    }

    #[test]
    fn folds_3d_plot_types_onto_flat_pipelines() {
        let pie = metadata("<c:plotArea><c:pie3DChart><c:ser/></c:pie3DChart></c:plotArea>");
        assert_eq!(pie.chart_types, vec!["pieChart"]);
        let bar = metadata(
            r#"<c:plotArea><c:bar3DChart><c:barDir val="col"/><c:grouping val="percentStacked"/><c:ser/></c:bar3DChart></c:plotArea>"#,
        );
        assert_eq!(bar.chart_types, vec!["barChart"]);
        assert_eq!(bar.grouping.as_deref(), Some("percentStacked"));
        assert_eq!(bar.bar_direction.as_deref(), Some("col"));
        let line = metadata("<c:plotArea><c:line3DChart/><c:area3DChart/></c:plotArea>");
        assert_eq!(line.chart_types, vec!["lineChart", "areaChart"]);
        // Unmapped 3D plots keep the old behavior: no chart type.
        assert!(
            metadata("<c:plotArea><c:surface3DChart/></c:plotArea>")
                .chart_types
                .is_empty()
        );
    }

    #[test]
    fn default_series_accent_follows_ser_idx() {
        let body = r#"<c:plotArea><c:barChart>
            <c:ser><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr></c:ser>
            <c:ser><c:idx val="0"/></c:ser>
            <c:ser><c:idx val="1"/></c:ser>
        </c:barChart></c:plotArea>"#;
        let chart = metadata_with(body, &theme_colors());
        // Explicit fill wins regardless of idx.
        assert_eq!(chart.series[0].color.as_deref(), Some("#FF8800"));
        // idx 0 → accent1 (theme slot 4), idx 1 → accent2 (slot 5).
        assert_eq!(chart.series[1].color.as_deref(), Some("#042233"));
        assert_eq!(chart.series[2].color.as_deref(), Some("#052233"));
        // Without c:idx the document position keeps driving the cycle.
        let plain = "<c:plotArea><c:barChart><c:ser/><c:ser/></c:barChart></c:plotArea>";
        let chart = metadata_with(plain, &theme_colors());
        assert_eq!(chart.series[0].color.as_deref(), Some("#042233"));
        assert_eq!(chart.series[1].color.as_deref(), Some("#052233"));
    }

    #[test]
    fn reads_plot_level_line_marker_flag() {
        let on = r#"<c:plotArea><c:lineChart><c:ser><c:marker><c:symbol val="none"/></c:marker></c:ser><c:marker val="1"/></c:lineChart></c:plotArea>"#;
        let chart = metadata(on);
        assert_eq!(chart.line_markers, Some(true));
        // The per-series marker container stays a symbol, not the plot flag.
        assert_eq!(chart.series[0].marker.as_deref(), Some("none"));
        let off = r#"<c:plotArea><c:lineChart><c:marker val="0"/></c:lineChart></c:plotArea>"#;
        assert_eq!(metadata(off).line_markers, Some(false));
        // Bare <c:marker/> is CT_Boolean true; absent stays absent.
        let bare = r#"<c:plotArea><c:lineChart><c:marker/></c:lineChart></c:plotArea>"#;
        assert_eq!(metadata(bare).line_markers, Some(true));
        let series_only = r#"<c:plotArea><c:lineChart><c:ser><c:marker><c:symbol val="diamond"/></c:marker></c:ser></c:lineChart></c:plotArea>"#;
        assert_eq!(metadata(series_only).line_markers, None);
    }

    #[test]
    fn multi_level_categories_use_the_innermost_level_only() {
        let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:f>D!$B$2:$E$3</c:f><c:multiLvlStrCache><c:ptCount val="4"/>
            <c:lvl><c:pt idx="0"><c:v>Qtr 1</c:v></c:pt><c:pt idx="1"><c:v>Qtr 2</c:v></c:pt><c:pt idx="2"><c:v>Qtr 3</c:v></c:pt><c:pt idx="3"><c:v>Qtr 4</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        assert_eq!(
            metadata(body).series[0].categories,
            vec!["Qtr 1", "Qtr 2", "Qtr 3", "Qtr 4"]
        );
    }

    #[test]
    fn reads_series_line_width_in_px() {
        let body = r#"<c:plotArea><c:lineChart><c:ser><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:ser><c:ser/></c:lineChart></c:plotArea>"#;
        let chart = metadata(body);
        // 12700 EMU = 1pt = 4/3 px.
        let width = chart.series[0].line_width.unwrap();
        assert!((width - 4.0 / 3.0).abs() < 1e-9);
        assert!(chart.series[1].line_width.is_none());
        let json = serde_json::to_value(&chart).unwrap();
        assert!(json["series"][0].get("lineWidth").is_some());
        assert!(json["series"][1].get("lineWidth").is_none());
    }

    #[test]
    fn outer_category_levels_become_groups_with_spans() {
        let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="8"/>
            <c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt>
                <c:pt idx="4"><c:v>Q1</c:v></c:pt><c:pt idx="5"><c:v>Q2</c:v></c:pt><c:pt idx="6"><c:v>Q3</c:v></c:pt><c:pt idx="7"><c:v>Q4</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="4"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        let groups = metadata(body).series[0].category_groups.clone().unwrap();
        let spans: Vec<(&str, usize, usize)> = groups
            .iter()
            .map(|group| (group.label.as_str(), group.start, group.end))
            .collect();
        assert_eq!(spans, vec![("2008", 0, 4), ("2009", 4, 8)]);
        let json = serde_json::to_value(metadata(body)).unwrap();
        assert_eq!(
            json["series"][0]["categoryGroups"][0],
            serde_json::json!({ "label": "2008", "start": 0, "end": 4 })
        );
        // Single-level caches carry no groups.
        let flat = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="2"/>
            <c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt>
        </c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        assert!(metadata(flat).series[0].category_groups.is_none());
    }

    #[test]
    fn category_group_spans_follow_the_compacted_innermost_positions() {
        // Innermost idx 2 is missing: the compacted categories are [Q1, Q2,
        // Q4, Q1] so the 2008 group must span 3 positions, not 4.
        let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="5"/>
            <c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt><c:pt idx="4"><c:v>Q1</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="4"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        let chart = metadata(body);
        assert_eq!(chart.series[0].categories, vec!["Q1", "Q2", "Q4", "Q1"]);
        let groups = chart.series[0].category_groups.clone().unwrap();
        let spans: Vec<(&str, usize, usize)> = groups
            .iter()
            .map(|group| (group.label.as_str(), group.start, group.end))
            .collect();
        assert_eq!(spans, vec![("2008", 0, 3), ("2009", 3, 4)]);
        // An outer pt without idx cannot be placed: no groups at all.
        let no_idx = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="2"/>
            <c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:lvl>
            <c:lvl><c:pt><c:v>2008</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        assert!(metadata(no_idx).series[0].category_groups.is_none());
    }

    #[test]
    fn category_group_spans_follow_document_order_not_idx_order() {
        // Innermost pts arrive out of idx order: categories emit as
        // [Q1'09, Q1'08, Q2'08], so the 2008 group sits at positions 1-2
        // and the 2009 group at position 0.
        let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="3"/>
            <c:lvl><c:pt idx="2"><c:v>Q1 09</c:v></c:pt><c:pt idx="0"><c:v>Q1 08</c:v></c:pt><c:pt idx="1"><c:v>Q2 08</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="2"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        let chart = metadata(body);
        assert_eq!(chart.series[0].categories, vec!["Q1 09", "Q1 08", "Q2 08"]);
        let groups = chart.series[0].category_groups.clone().unwrap();
        let spans: Vec<(&str, usize, usize)> = groups
            .iter()
            .map(|group| (group.label.as_str(), group.start, group.end))
            .collect();
        assert_eq!(spans, vec![("2008", 1, 3), ("2009", 0, 1)]);
        // Interleaved emission leaves a group non-contiguous: a span cannot
        // express it, so no groups at all.
        let interleaved = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="3"/>
            <c:lvl><c:pt idx="0"><c:v>Q1 08</c:v></c:pt><c:pt idx="2"><c:v>Q1 09</c:v></c:pt><c:pt idx="1"><c:v>Q2 08</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="2"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        assert!(metadata(interleaved).series[0].category_groups.is_none());
    }

    #[test]
    fn sparse_caches_stay_compacted_in_document_order() {
        // Both cache kinds compact sparse points the same way, keeping
        // categories aligned with values (idx-union alignment is a leftover).
        let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="3"/>
            <c:pt idx="2"><c:v>c</c:v></c:pt><c:pt idx="0"><c:v>a</c:v></c:pt>
        </c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
        assert_eq!(metadata(body).series[0].categories, vec!["c", "a"]);
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

        let plain =
            "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser></c:pieChart></c:plotArea>";
        assert!(metadata(plain).series[0].point_colors.is_none());
    }

    #[test]
    fn reads_gridlines_only_when_a_value_axis_exists() {
        let with = "<c:plotArea><c:barChart/><c:catAx/><c:valAx><c:majorGridlines/></c:valAx></c:plotArea>";
        assert_eq!(metadata(with).gridlines, Some(true));
        let without = "<c:plotArea><c:barChart/><c:catAx/><c:valAx/></c:plotArea>";
        assert_eq!(metadata(without).gridlines, Some(false));
        assert!(
            metadata("<c:plotArea><c:pieChart/></c:plotArea>")
                .gridlines
                .is_none()
        );
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
        assert!(
            metadata("<c:plotArea><c:pieChart/></c:plotArea>")
                .value_axis
                .is_none()
        );
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
        assert_eq!(
            metadata(axis_level).category_axis_format.as_deref(),
            Some("0.0%")
        );
        let date_axis = r#"<c:plotArea><c:lineChart/><c:dateAx><c:numFmt formatCode="mmm\-yy" sourceLinked="1"/></c:dateAx><c:valAx/></c:plotArea>"#;
        assert_eq!(
            metadata(date_axis).category_axis_format.as_deref(),
            Some("mmm\\-yy")
        );

        // scatter X data (c:xVal) carries the same field
        let scatter = r#"<c:plotArea><c:scatterChart><c:ser>
            <c:xVal><c:numRef><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>0.15</c:v></c:pt></c:numCache></c:numRef></c:xVal>
            <c:yVal><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>0.4</c:v></c:pt></c:numCache></c:numRef></c:yVal>
        </c:ser></c:scatterChart></c:plotArea>"#;
        assert_eq!(
            metadata(scatter).series[0].category_format.as_deref(),
            Some("0%")
        );

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
        assert!(
            metadata("<c:plotArea><c:barChart/></c:plotArea>")
                .gap_width_pct
                .is_none()
        );

        let doughnut =
            "<c:plotArea><c:doughnutChart><c:holeSize val=\"65\"/></c:doughnutChart></c:plotArea>";
        assert_eq!(metadata(doughnut).hole_size_pct, Some(65));
        assert!(
            metadata("<c:plotArea><c:doughnutChart/></c:plotArea>")
                .hole_size_pct
                .is_none()
        );
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

        let plain =
            "<c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/></c:ser></c:pieChart></c:plotArea>";
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

    #[test]
    fn outline_solid_fill_is_not_the_shape_fill() {
        let sppr = r#"<xdr:spPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
            </xdr:spPr>"#;
        let document = Document::parse(sppr).unwrap();
        let root = document.root_element();
        assert_eq!(drawing_fill_color(root, &ColorContext::default()), None);
        // Called on the a:ln itself, its solidFill is the answer.
        let line = root
            .children()
            .find(|node| node.has_tag_name("ln"))
            .unwrap();
        assert_eq!(
            drawing_fill_color(line, &ColorContext::default()).as_deref(),
            Some("#FF0000")
        );

        let sppr = r#"<xdr:spPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
            <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
            </xdr:spPr>"#;
        let document = Document::parse(sppr).unwrap();
        assert_eq!(
            drawing_fill_color(document.root_element(), &ColorContext::default()).as_deref(),
            Some("#00FF00")
        );
    }

    /// Producers may bake a wrong rgb cache next to a theme reference
    /// (tdf113271: theme="1" rgb="FFFFFF" on black dk1 text); the theme
    /// slot wins, and rgb is the fallback when the slot cannot resolve.
    #[test]
    fn theme_attribute_wins_over_rgb_cache() {
        let colors = theme_colors();
        assert_eq!(
            resolve_color(Some("FFFFFF"), None, Some("1"), None, &colors).as_deref(),
            Some("#012233")
        );
        // Slot outside the palette (or no palette at all) → rgb cache.
        assert_eq!(
            resolve_color(Some("FF00FF"), None, Some("99"), None, &colors).as_deref(),
            Some("#FF00FF")
        );
        assert_eq!(
            resolve_color(
                Some("FF00FF"),
                None,
                Some("1"),
                None,
                &ColorContext::default()
            )
            .as_deref(),
            Some("#FF00FF")
        );
    }

    #[test]
    fn resolves_style_fill_ref_theme_gradient() {
        let theme = r#"<a:fillStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:gradFill rotWithShape="1"><a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>
            </a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>
            </a:fillStyleLst>"#;
        let document = Document::parse(theme).unwrap();
        let fill_styles: Vec<Option<ThemeGradient>> = document
            .root_element()
            .children()
            .filter(|node| node.is_element())
            .map(parse_theme_gradient)
            .collect();
        assert!(fill_styles[0].is_none());
        let colors = ColorContext {
            fill_styles,
            ..theme_colors()
        };

        let style = r#"<xdr:style xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
            </xdr:style>"#;
        let document = Document::parse(style).unwrap();
        let gradient =
            style_fill_gradient(Some(document.root_element()), &colors).expect("gradient");
        assert_eq!(gradient.angle, 270.0);
        // accent1 = theme slot 4 = (0x04, 0x22, 0x33); second stop is a 50% tint.
        assert_eq!(gradient.stops[0].position, 0.0);
        assert_eq!(gradient.stops[0].color, "#042233");
        assert_eq!(gradient.stops[1].position, 1.0);
        assert_eq!(gradient.stops[1].color, "#829199");

        // idx 1 is the solid entry; idx 1001+ (bgFillStyleLst) is out of range.
        for idx in ["1", "1001", "0"] {
            let style = format!(
                r#"<xdr:style xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fillRef idx="{idx}"><a:schemeClr val="accent1"/></a:fillRef></xdr:style>"#
            );
            let document = Document::parse(&style).unwrap();
            assert!(style_fill_gradient(Some(document.root_element()), &colors).is_none());
        }
    }
}

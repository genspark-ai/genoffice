use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Read;
use std::path::{Component, Path, PathBuf};

use base64::Engine;
use roxmltree::{Document, Node};
use serde::Serialize;
use zip::ZipArchive;

use crate::SidecarError;

mod charts;
mod colors;
mod drawing;
mod styles;
#[cfg(test)]
mod tests;

pub(crate) use charts::*;
pub(crate) use colors::*;
pub(crate) use drawing::*;
pub(crate) use styles::*;

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
    /// Indices whose value cache holds no point (blank cells). `values`
    /// carries 0 there; the renderer applies `c:dispBlanksAs`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blanks: Option<Vec<usize>>,
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
    /// `c:dispBlanksAs` — gap/zero/span (OOXML defaults to zero).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub disp_blanks_as: Option<String>,
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
    /// a:srcRect as 0..1 fractions cut from each source edge; absent when
    /// the picture is uncropped.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop: Option<CropRect>,
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
    /// Worksheet `<oleObject progId=…>` of an embedded object (kind "ole");
    /// the renderer maps it to a caption when no cached preview exists.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
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

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CropRect {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
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
pub(crate) struct FontStyle {
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
pub(crate) struct BorderSet {
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
    /// A context with just a theme palette (`[lt1, dk1, lt2, dk2, accent1-6,
    /// hlink, folHlink]` order), for palette calibration tests.
    #[cfg(test)]
    pub(crate) fn with_theme(theme: Vec<(u8, u8, u8)>) -> Self {
        ColorContext {
            theme,
            fill_styles: Vec::new(),
            indexed: Vec::new(),
        }
    }

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

pub fn read_visual_objects(
    archive: &mut ZipArchive<File>,
    sheets: &[SheetVisualSource],
    colors: &ColorContext,
) -> Result<Vec<VisualObject>, SidecarError> {
    let mut visuals = Vec::new();
    // Workbook-wide serial for `ole-N` ids: the list position is not usable
    // because an OLE visual may take a fallback shape's slot mid-list.
    let mut ole_serial = 0usize;
    for sheet in sheets {
        let sheet_relationships = read_relationships(archive, &sheet.worksheet_path)?;
        let ole_objects = read_ole_objects(archive, &sheet.worksheet_path, &sheet_relationships)?;
        let ole_shape_ids: HashSet<u32> = ole_objects.iter().map(|ole| ole.shape_id).collect();
        let drawing_relationship = sheet_relationships
            .values()
            .find(|relationship| relationship.relationship_type.ends_with("/drawing"));
        // A sheet can carry OLE objects with no drawing part at all (the
        // legacy VML shape is their only anchor), so do not skip such sheets.
        if drawing_relationship.is_none() && ole_objects.is_empty() {
            continue;
        }
        let start = visuals.len();
        if let Some(drawing_relationship) = drawing_relationship {
            let drawing_path =
                resolve_part_target(&sheet.worksheet_path, &drawing_relationship.target)?;
            visuals.extend(read_drawing(
                archive,
                &drawing_path,
                &sheet.sheet_id,
                visuals.len(),
                colors,
                &ole_shape_ids,
            )?);
        }
        if ole_objects.is_empty() {
            continue;
        }
        // Excel 2010+ also writes a hidden `xdr:sp` compat fallback per OLE
        // object (same anchor, cNvPr id == shapeId). The OLE visual below is
        // the rendered form and takes the fallback's slot in the drawing so
        // it keeps Excel's z-order among the other shapes; the fallback only
        // lends its anchor when neither objectPr nor the VML shape carries
        // one. Objects without a fallback (no drawing part) go on top.
        let is_fallback = |visual: &VisualObject| {
            visual.kind == "shape" && visual.nv_id.is_some_and(|id| ole_shape_ids.contains(&id))
        };
        let mut fallback_slots: HashMap<u32, usize> = HashMap::new();
        for (index, visual) in visuals.iter().enumerate().skip(start) {
            if is_fallback(visual) {
                if let Some(id) = visual.nv_id {
                    fallback_slots.entry(id).or_insert(index);
                }
            }
        }
        for ole in ole_objects {
            let slot = fallback_slots.get(&ole.shape_id).copied();
            let Some(anchor) = ole
                .anchor
                .or_else(|| slot.map(|at| visuals[at].anchor.clone()))
            else {
                continue;
            };
            if anchor_is_zero_extent(&anchor) {
                continue;
            }
            ole_serial += 1;
            let visual = VisualObject {
                id: format!("ole-{ole_serial}"),
                sheet_id: sheet.sheet_id.clone(),
                kind: "ole".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_type: ole
                    .preview_path
                    .as_deref()
                    .and_then(media_type_for_path)
                    .map(ToOwned::to_owned),
                media_path: ole.preview_path,
                opacity: None,
                crop: None,
                fill_media_path: None,
                fill_media_type: None,
                name: None,
                shape_type: None,
                custom_path: None,
                fill_color: ole.fill_color,
                fill_gradient: None,
                line_color: ole.frame_color,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                prog_id: Some(ole.prog_id),
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                // No edit locator on purpose: an OLE object lives in the
                // worksheet's <oleObjects>, its embedding part and the legacy
                // VML shape — none of which the drawing edit pipeline
                // rewrites — so it stays read-only and round-trips untouched.
                drawing_path: None,
                drawing_index: None,
            };
            match slot {
                Some(at) => visuals[at] = visual,
                None => visuals.push(visual),
            }
        }
        // Fallbacks whose object was skipped (zero extent, no anchor) must
        // not surface as shapes either. Ids were assigned by read_drawing,
        // so compacting the list here does not disturb them.
        visuals.retain(|visual| !is_fallback(visual));
    }
    Ok(visuals)
}

/// One worksheet `<oleObject>`: Excel's embedded (or linked) object record.
#[derive(Clone, Debug)]
pub(crate) struct OleObjectRecord {
    /// `@shapeId` — pairs the record with its legacy VML shape
    /// (`_x0000_s<shapeId>`) and the hidden drawing fallback shape.
    pub(crate) shape_id: u32,
    pub(crate) prog_id: String,
    /// Excel's cached preview picture (`objectPr/@r:id`, else the VML
    /// shape's `v:imagedata`), usually an EMF under xl/media.
    pub(crate) preview_path: Option<String>,
    /// `objectPr/anchor` (x14 form), else the VML `x:Anchor`.
    pub(crate) anchor: Option<DrawingAnchor>,
    /// VML `stroked`/`strokecolor`: Excel's hairline frame around the object.
    pub(crate) frame_color: Option<String>,
    /// VML `filled`/`fillcolor`: the opaque backdrop hiding the grid.
    pub(crate) fill_color: Option<String>,
}

/// EMU per CSS pixel at 96 DPI — VML anchors store offsets in pixels.
const EMU_PER_PIXEL: i64 = 9525;

/// Worksheet `<oleObjects>`, deduplicated by shapeId: Excel repeats every
/// object inside `mc:AlternateContent` (the x14 `mc:Choice` carries
/// objectPr with the anchor and preview, the `mc:Fallback` copy is bare) and
/// older writers emit the bare form only, in which case the legacy VML
/// drawing supplies the anchor and preview picture.
pub(crate) fn read_ole_objects(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    sheet_relationships: &HashMap<String, Relationship>,
) -> Result<Vec<OleObjectRecord>, SidecarError> {
    let xml = read_xml(archive, worksheet_path)?;
    if !xml.contains("oleObject") {
        return Ok(Vec::new());
    }
    let document = parse_document(&xml, worksheet_path)?;
    let mut records: Vec<OleObjectRecord> = Vec::new();
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
        let object_pr = direct_child(node, "objectPr");
        let preview_path = object_pr
            .and_then(relationship_id)
            .and_then(|id| sheet_relationships.get(&id))
            .map(|relationship| resolve_part_target(worksheet_path, &relationship.target))
            .transpose()?;
        let anchor = object_pr
            .and_then(|node| direct_child(node, "anchor"))
            .and_then(parse_anchor);
        match records
            .iter_mut()
            .find(|record| record.shape_id == shape_id)
        {
            Some(existing) => {
                // Choice/Fallback duplicates: keep the richer fields.
                if existing.preview_path.is_none() {
                    existing.preview_path = preview_path;
                }
                if existing.anchor.is_none() {
                    existing.anchor = anchor;
                }
            }
            None => records.push(OleObjectRecord {
                shape_id,
                prog_id: prog_id.to_owned(),
                preview_path,
                anchor,
                frame_color: None,
                fill_color: None,
            }),
        }
    }
    // The <legacyDrawing> VML shape always decides the object frame and
    // backdrop; for the bare form it also holds the anchor and picture.
    let legacy_path = document
        .descendants()
        .find(|node| node.has_tag_name("legacyDrawing"))
        .and_then(relationship_id)
        .and_then(|id| sheet_relationships.get(&id))
        .map(|relationship| resolve_part_target(worksheet_path, &relationship.target))
        .transpose()?;
    if let Some(legacy_path) = legacy_path {
        let vml_shapes = read_vml_object_shapes(archive, &legacy_path)?;
        for record in &mut records {
            let Some(shape) = vml_shapes.get(&record.shape_id) else {
                continue;
            };
            if record.anchor.is_none() {
                record.anchor = shape.anchor.clone();
            }
            if record.preview_path.is_none() {
                record.preview_path = shape.image_path.clone();
            }
            record.frame_color = shape.frame_color.clone();
            record.fill_color = shape.fill_color.clone();
        }
    }
    Ok(records)
}

#[derive(Clone, Debug, Default)]
pub(crate) struct VmlObjectShape {
    pub(crate) anchor: Option<DrawingAnchor>,
    pub(crate) image_path: Option<String>,
    /// `stroked="t"` → Excel draws a hairline frame around the object.
    pub(crate) frame_color: Option<String>,
    /// `filled="t"` → the object hides the grid behind it.
    pub(crate) fill_color: Option<String>,
}

/// VML system colours (`windowText [64]`, `window [65]`) and named colours
/// resolve to Excel's defaults; `#rrggbb` passes through.
fn vml_color(value: Option<&str>, default: &str) -> String {
    let value = value.map(str::trim).unwrap_or("");
    match value.split_once(' ').map_or(value, |(head, _)| head) {
        hex if hex.starts_with('#') && hex.len() == 7 => hex.to_ascii_uppercase(),
        _ => default.to_owned(),
    }
}

/// VML boolean attributes: "t"/"true"/"1" are true, "f"/"false"/"0" false.
fn vml_bool(value: &str) -> Option<bool> {
    match value.trim() {
        "t" | "true" | "1" => Some(true),
        "f" | "false" | "0" => Some(false),
        _ => None,
    }
}

/// `v:shape` entries of a legacy VML drawing keyed by their numeric shape
/// id (`id="_x0000_s1025"` → 1025): the `x:Anchor` cell box, the
/// `v:imagedata` picture and the frame/fill flags (a shape inherits
/// `filled`/`stroked` from its `v:shapetype`; VML's own default is true).
/// Malformed VML (some third-party writers) is treated as absent rather
/// than failing the whole workbook.
pub(crate) fn read_vml_object_shapes(
    archive: &mut ZipArchive<File>,
    vml_path: &str,
) -> Result<HashMap<u32, VmlObjectShape>, SidecarError> {
    let Some(xml) = read_optional_xml(archive, vml_path)? else {
        return Ok(HashMap::new());
    };
    let Ok(document) = Document::parse(&xml) else {
        return Ok(HashMap::new());
    };
    let relationships = read_relationships(archive, vml_path)?;
    let shape_types: HashMap<String, (Option<bool>, Option<bool>)> = document
        .descendants()
        .filter(|node| node.has_tag_name("shapetype"))
        .filter_map(|node| {
            let id = node.attribute("id")?;
            Some((
                format!("#{id}"),
                (
                    node.attribute("filled").and_then(vml_bool),
                    node.attribute("stroked").and_then(vml_bool),
                ),
            ))
        })
        .collect();
    let mut shapes = HashMap::new();
    for node in document
        .descendants()
        .filter(|node| node.has_tag_name("shape"))
    {
        let Some(shape_id) = node
            .attribute("id")
            .and_then(|id| id.rsplit_once("_s"))
            .and_then(|(_, digits)| digits.parse::<u32>().ok())
        else {
            continue;
        };
        let inherited = node
            .attribute("type")
            .and_then(|kind| shape_types.get(kind))
            .copied()
            .unwrap_or((None, None));
        let filled = node
            .attribute("filled")
            .and_then(vml_bool)
            .or(inherited.0)
            .unwrap_or(true);
        let stroked = node
            .attribute("stroked")
            .and_then(vml_bool)
            .or(inherited.1)
            .unwrap_or(true);
        let fill_color = filled.then(|| vml_color(node.attribute("fillcolor"), "#FFFFFF"));
        let frame_color = stroked.then(|| vml_color(node.attribute("strokecolor"), "#000000"));
        let anchor = node
            .descendants()
            .find(|child| child.has_tag_name("Anchor"))
            .and_then(|child| child.text())
            .and_then(parse_vml_anchor);
        let image_path = node
            .descendants()
            .find(|child| child.has_tag_name("imagedata"))
            .and_then(|child| {
                child
                    .attributes()
                    .find(|attribute| attribute.name() == "relid" || attribute.name() == "id")
                    .map(|attribute| attribute.value().to_owned())
            })
            .and_then(|id| relationships.get(&id))
            .map(|relationship| resolve_part_target(vml_path, &relationship.target))
            .transpose()?;
        shapes.insert(
            shape_id,
            VmlObjectShape {
                anchor,
                image_path,
                frame_color,
                fill_color,
            },
        );
    }
    Ok(shapes)
}

/// `x:Anchor` text: "fromCol, fromColOffPx, fromRow, fromRowOffPx, toCol,
/// toColOffPx, toRow, toRowOffPx" — offsets are pixels, not EMU.
pub(crate) fn parse_vml_anchor(text: &str) -> Option<DrawingAnchor> {
    let values: Vec<i64> = text
        .split(',')
        .map(|value| value.trim().parse::<i64>())
        .collect::<Result<_, _>>()
        .ok()?;
    let [
        from_column,
        from_column_px,
        from_row,
        from_row_px,
        to_column,
        to_column_px,
        to_row,
        to_row_px,
    ] = values[..]
    else {
        return None;
    };
    if from_column < 0 || from_row < 0 || to_column < 0 || to_row < 0 {
        return None;
    }
    Some(DrawingAnchor {
        from_row: from_row as usize,
        from_column: from_column as usize,
        from_row_offset: from_row_px * EMU_PER_PIXEL,
        from_column_offset: from_column_px * EMU_PER_PIXEL,
        to_row: to_row as usize,
        to_column: to_column as usize,
        to_row_offset: to_row_px * EMU_PER_PIXEL,
        to_column_offset: to_column_px * EMU_PER_PIXEL,
        explicit_to: true,
    })
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
    /// location/@firstDataCol — columns left of it are row-label columns.
    pub first_data_col: usize,
    pub row_grand_totals: bool,
    pub show_row_stripes: bool,
    pub show_col_stripes: bool,
    /// One char per `<rowItems><i>` (rows from firstDataRow down): `d` data,
    /// `s` level-1 row subheading, `S` deeper subheading, `t` subtotal,
    /// `g` grand total, `b` blank spacer. Empty when the part has none.
    pub row_kinds: String,
}

/// Classify the pivot's row items so the renderer can paint Excel's
/// subheading / subtotal / grand-total bands. A data item shallower than
/// the last row field is an outer item on its own row (compact and outline
/// forms; tabular fields — `outline="0"` — put the inner item on the same
/// row, so those stay body rows).
fn pivot_row_kinds(document: &Document) -> String {
    let root = document.root_element();
    let pivot_fields: Vec<Node> = root
        .children()
        .find(|node| node.has_tag_name("pivotFields"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("pivotField"))
                .collect()
        })
        .unwrap_or_default();
    let row_field_outline: Vec<bool> = root
        .children()
        .find(|node| node.has_tag_name("rowFields"))
        .map(|node| {
            node.children()
                .filter(|child| child.has_tag_name("field"))
                .map(|field| {
                    field
                        .attribute("x")
                        .and_then(|value| value.parse::<usize>().ok())
                        .and_then(|index| pivot_fields.get(index))
                        .map(|pivot_field| pivot_field.attribute("outline") != Some("0"))
                        .unwrap_or(true)
                })
                .collect()
        })
        .unwrap_or_default();
    let row_field_count = row_field_outline.len();
    let Some(row_items) = root.children().find(|node| node.has_tag_name("rowItems")) else {
        return String::new();
    };
    row_items
        .children()
        .filter(|child| child.has_tag_name("i"))
        .map(|item| {
            let depth = item
                .attribute("r")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(0);
            match item.attribute("t").unwrap_or("data") {
                "grand" => 'g',
                "blank" => 'b',
                "data" => {
                    let outer = depth + 1 < row_field_count
                        && row_field_outline.get(depth).copied().unwrap_or(true);
                    if !outer {
                        'd'
                    } else if depth == 0 {
                        's'
                    } else {
                        'S'
                    }
                }
                // default / sum / countA / avg / max / min / product / count /
                // stdDev / stdDevP / var / varP: a subtotal row.
                _ => 't',
            }
        })
        .collect()
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
        let first_data_col = location
            .attribute("firstDataCol")
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
        let show_col_stripes = style_info
            .and_then(|node| node.attribute("showColStripes"))
            .is_some_and(|value| value == "1" || value == "true");
        let row_kinds = pivot_row_kinds(&document);
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
            first_data_col,
            row_grand_totals,
            show_row_stripes,
            show_col_stripes,
            row_kinds,
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

/// Header/footer pictures larger than this are skipped: they travel to the
/// renderer as data URLs inside the print templates.
const MAX_HEADER_FOOTER_PICTURE_BYTES: u64 = 2 * 1024 * 1024;
/// Six slots × three page variants.
const MAX_HEADER_FOOTER_PICTURES: usize = 18;

/// `&G` pictures behind a worksheet's `<legacyDrawingHF r:id>`: the VML
/// part's `<v:shape id="LH|CH|RH|LF|CF|RF[EVEN|FIRST]">` names the slot,
/// its style carries the printed size, and `<v:imagedata o:relid>` points
/// at the media part through the VML part's own relationships. Legacy VML
/// is not always well-formed XML — an unparseable part loses the pictures,
/// not the sheet.
pub fn read_header_footer_pictures(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    relationship_id: &str,
    sheet_index: usize,
) -> Result<Vec<crate::HeaderFooterPictureInfo>, SidecarError> {
    let relationships = read_relationships(archive, worksheet_path)?;
    let Some(relationship) = relationships.get(relationship_id) else {
        return Ok(Vec::new());
    };
    let vml_path = resolve_part_target(worksheet_path, &relationship.target)?;
    let Some(xml) = read_optional_xml(archive, &vml_path)? else {
        return Ok(Vec::new());
    };
    let Ok(document) = Document::parse(&xml) else {
        return Ok(Vec::new());
    };
    let media_relationships = read_relationships(archive, &vml_path)?;
    let mut pictures: Vec<crate::HeaderFooterPictureInfo> = Vec::new();
    for shape in document
        .descendants()
        .filter(|node| node.has_tag_name("shape"))
    {
        let Some(position) = shape.attribute("id").and_then(header_footer_slot) else {
            continue;
        };
        if pictures.iter().any(|picture| picture.position == position) {
            continue;
        }
        let Some(image_data) = direct_child(shape, "imagedata") else {
            continue;
        };
        // Excel writes o:relid; some producers use r:id instead.
        let Some(relationship_id) = image_data
            .attributes()
            .find(|attribute| attribute.name() == "relid" || attribute.name() == "id")
            .map(|attribute| attribute.value())
        else {
            continue;
        };
        let Some(media) = media_relationships.get(relationship_id) else {
            continue;
        };
        let media_path = resolve_part_target(&vml_path, &media.target)?;
        let Some(media_type) = media_type_for_path(&media_path) else {
            continue;
        };
        let Some((width_pt, height_pt)) = shape.attribute("style").and_then(vml_size_pt) else {
            continue;
        };
        let Ok(entry) = crate::zip_entry(archive, &media_path) else {
            continue;
        };
        let size = entry.size();
        drop(entry);
        if size == 0 || size > MAX_HEADER_FOOTER_PICTURE_BYTES {
            continue;
        }
        pictures.push(crate::HeaderFooterPictureInfo {
            id: format!("hf-picture-{sheet_index}-{}", position.to_ascii_lowercase()),
            position,
            width_pt,
            height_pt,
            media_type: media_type.to_owned(),
            media_path,
        });
        if pictures.len() >= MAX_HEADER_FOOTER_PICTURES {
            break;
        }
    }
    Ok(pictures)
}

/// `LH`/`cfFirst`/`RFEVEN` → the normalized slot name; anything else (comment
/// shapes, form controls) is not a header/footer picture.
fn header_footer_slot(shape_id: &str) -> Option<String> {
    let upper = shape_id.trim().to_ascii_uppercase();
    let bytes = upper.as_bytes();
    if bytes.len() < 2
        || !matches!(bytes[0], b'L' | b'C' | b'R')
        || !matches!(bytes[1], b'H' | b'F')
    {
        return None;
    }
    matches!(&upper[2..], "" | "EVEN" | "FIRST").then_some(upper)
}

/// `width:442.5pt;height:43.5pt` (VML style) → (width, height) in points.
fn vml_size_pt(style: &str) -> Option<(f64, f64)> {
    let mut width = None;
    let mut height = None;
    for declaration in style.split(';') {
        let Some((name, value)) = declaration.split_once(':') else {
            continue;
        };
        match name.trim().to_ascii_lowercase().as_str() {
            "width" => width = css_length_pt(value.trim()),
            "height" => height = css_length_pt(value.trim()),
            _ => {}
        }
    }
    Some((width?, height?))
}

/// A CSS length with pt/in/cm/mm/px/pc units (unitless = pt) → points;
/// None for anything non-positive, non-finite or implausibly large.
fn css_length_pt(value: &str) -> Option<f64> {
    let split = value
        .find(|character: char| character.is_ascii_alphabetic() || character == '%')
        .unwrap_or(value.len());
    let number: f64 = value[..split].trim().parse().ok()?;
    let factor = match value[split..].trim().to_ascii_lowercase().as_str() {
        "" | "pt" => 1.0,
        "in" => 72.0,
        "cm" => 72.0 / 2.54,
        "mm" => 72.0 / 25.4,
        "px" => 0.75,
        "pc" => 12.0,
        _ => return None,
    };
    let points = number * factor;
    (points.is_finite() && points > 0.0 && points <= 2_000.0).then_some(points)
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
///  - 27-36 / 50-58: locale-dependent date/time formats, resolved per viewer
///    locale in locale_reserved_number_format below.
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
        22 => Some(format!("{short_date} hh:mm")),
        _ => None,
    }
}

/// Ids 27-36/50-58 mean a different pattern per viewing locale
/// ([MS-OI29500]); the file records nothing. Era-year patterns (ja ge/ggge)
/// render Gregorian and the zh AM/PM token (U+4E0A/U+4E0B U+5348) renders
/// 24-hour — the renderer's numfmt supports neither. Ids 55/56 follow the
/// host OS short date at the call site first (a ja workbook pinned them to
/// a plain date, #049); the values here are the no-host fallback. Non-CJK
/// Excel renders the long-date id 31 with slash separators in y/m/d order
/// (prod refs: 2025/3/29), while the remaining ids keep the viewer's short
/// date so a CJK month/day format cannot discard the year. Escapes: U+5E74
/// year, U+6708 month, U+65E5 day, U+65F6/U+6642 hour, U+5206 minute,
/// U+79D2 second; ko U+B144 year, U+C6D4 month, U+C77C day, U+C2DC hour,
/// U+BD84 minute, U+CD08 second.
fn locale_reserved_number_format(id: u32, locale: &str) -> Option<&'static str> {
    if !matches!(id, 27..=36 | 50..=58) {
        return None;
    }
    Some(match locale {
        "ja" => match id {
            27 | 36 | 50 | 57 => "yyyy/m/d",
            28 | 29 | 31 | 51 | 54 | 58 => "yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\"",
            30 => "m/d/yy",
            32 => "h\"\u{6642}\"mm\"\u{5206}\"",
            33 => "h\"\u{6642}\"mm\"\u{5206}\"ss\"\u{79d2}\"",
            34 | 52 => "yyyy\"\u{5e74}\"m\"\u{6708}\"",
            35 | 53 => "m\"\u{6708}\"d\"\u{65e5}\"",
            _ => "yyyy/m/d",
        },
        "ko" => match id {
            27 | 36 | 50 | 57 => "yyyy\"\u{5e74}\" mm\"\u{6708}\" dd\"\u{65e5}\"",
            28 | 29 | 51 | 54 | 58 => "mm-dd",
            30 => "mm-dd-yy",
            31 => "yyyy\"\u{b144}\" mm\"\u{c6d4}\" dd\"\u{c77c}\"",
            32 => "h\"\u{c2dc}\" mm\"\u{bd84}\"",
            33 => "h\"\u{c2dc}\" mm\"\u{bd84}\" ss\"\u{cd08}\"",
            34 | 35 | 52 | 53 => "yyyy-mm-dd",
            _ => "yyyy/m/d",
        },
        "zh" | "zh-TW" => match id {
            27 | 36 | 50 | 52 | 57 => "yyyy\"\u{5e74}\"m\"\u{6708}\"",
            28 | 29 | 51 | 53 | 54 | 58 => "m\"\u{6708}\"d\"\u{65e5}\"",
            30 => "m-d-yy",
            31 => "yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\"",
            32 | 34 => "h\"\u{65f6}\"mm\"\u{5206}\"",
            33 | 35 => "h\"\u{65f6}\"mm\"\u{5206}\"ss\"\u{79d2}\"",
            _ => "yyyy/m/d",
        },
        _ => match id {
            31 => "yyyy/m/d",
            _ => locale_short_date_format(locale),
        },
    })
}

fn builtin_number_format(id: u32, locale: &str) -> Option<&'static str> {
    if let Some(reserved) = locale_reserved_number_format(id, locale) {
        return Some(reserved);
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
        // ECMA-376 prints 20/21/22 with h:mm(:ss), but Excel renders these
        // builtins with a leading zero on the hour (09:30, matching
        // LibreOffice's HH:MM mapping) — verified against Excel output,
        // 22 against the prod refs (02:18).
        20 => Some("hh:mm"),
        21 => Some("hh:mm:ss"),
        22 => Some("m/d/yy hh:mm"),
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
        // ECMA-376 prints 47 as "mmss.0", but Excel renders a colon between
        // the minutes and seconds (18:02.0 in the prod refs).
        47 => Some("mm:ss.0"),
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

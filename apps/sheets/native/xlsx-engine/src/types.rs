//! Serialized protocol types: the workbook / sheet metadata, cell, range,
//! print, protection and conditional-formatting records the sidecar writes
//! to the JSON wire. Field names are the camelCase protocol contract.

use super::*;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookMetadata {
    pub session_id: String,
    pub name: String,
    pub entry_count: usize,
    pub sheets: Vec<SheetMetadata>,
    /// workbookView/@activeTab (sheet index in workbook order); 0 when absent.
    pub active_tab: usize,
    pub styles: Vec<CellStyle>,
    pub dxf_styles: Vec<CellStyle>,
    pub visuals: Vec<VisualObject>,
    pub defined_names: Vec<DefinedName>,
    /// Theme palette as `#RRGGBB` in theme index order; None without a theme.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_colors: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_fonts: Option<ThemeFonts>,
    /// Literal cached <name val> of the Normal (cellXfs[0]) font, before any
    /// theme-scheme substitution; the renderer derives column-width MDW from
    /// it the way Excel does.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub normal_font_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_protection: Option<WorkbookProtectionInfo>,
    /// workbookPr/@date1904: serial dates count from 1904-01-01.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub date1904: bool,
    /// System short-date pattern applied to builtin numFmtIds 14/22, echoed
    /// so the renderer's own short-date affordances match.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub short_date_format: Option<String>,
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookProtectionInfo {
    pub lock_structure: bool,
    /// workbookPassword= (legacy) or workbookHashValue= (modern) present.
    pub has_password: bool,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
    /// localSheetId attribute: position in workbook sheet order.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet_index: Option<usize>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMetadata {
    pub id: String,
    pub name: String,
    pub row_count: usize,
    pub column_count: usize,
    /// Uncompressed worksheet XML size. The host uses this to avoid
    /// background recovery saves whose string-based patch path would require
    /// several copies of a very large entry in the Electron main process.
    pub source_xml_bytes: u64,
    pub column_widths: Vec<ColumnWidth>,
    pub default_row_height: Option<f64>,
    /// sheetFormatPr/@customHeight: the default row height is user-fixed, so
    /// Excel keeps wrap rows without their own ht at the default (clipped)
    /// instead of auto-fitting them on open.
    pub default_row_height_fixed: bool,
    pub default_column_width: Option<f64>,
    /// sheetFormatPr/@baseColWidth — Excel derives its built-in default
    /// column width from this (default 8) when defaultColWidth is absent.
    pub base_column_width: Option<f64>,
    pub freeze: Option<FreezePane>,
    pub hidden: bool,
    pub tab_color: Option<String>,
    pub show_grid_lines: bool,
    /// sheetView/@showFormulas: the sheet opens in formula view (#188).
    pub show_formulas: bool,
    /// sheetView/@showRowColHeaders: row/column heading strips.
    pub show_row_col_headers: bool,
    /// sheetView/@rightToLeft: the grid is mirrored (column A at the right).
    pub right_to_left: bool,
    /// Saved normal-view zoom percent (10-400); omitted at the 100% default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u16>,
    pub tables: Vec<TableInfo>,
    pub comments: Vec<CommentInfo>,
    /// PivotTable output areas — protected from edits by the renderer.
    pub pivot_ranges: Vec<MergedRange>,
    /// One entry per pivot table part: enough for the host to read and
    /// parse the definition on demand.
    pub pivot_tables: Vec<PivotTableInfo>,
    /// x14 sparkline groups from the worksheet extLst.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub sparklines: Vec<SparklineGroupInfo>,
    /// Saved `_xlnm.Print_Area` formula for this sheet (workbook.xml).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_area: Option<String>,
    /// Saved `_xlnm.Print_Titles` formula for this sheet.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_titles: Option<String>,
    /// Any localSheetId-scoped definedName targets this sheet (including
    /// hidden and _xlnm.* built-ins): the save refuses to clone such a
    /// sheet, so the host gates duplication up front on the same predicate.
    pub has_scoped_defined_names: bool,
    /// In-cell rich-value pictures ("place picture in cell").
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub cell_images: Vec<CellImageInfo>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellImageInfo {
    /// Media lookup key for the read_media command, like a visual id.
    pub id: String,
    pub row: usize,
    pub column: usize,
    #[serde(skip_serializing)]
    pub media_path: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroupInfo {
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub negative_color: Option<String>,
    pub cells: Vec<SparklineCellInfo>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklineCellInfo {
    pub cell: String,
    pub source_ref: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableInfo {
    pub path: String,
    pub cache_path: Option<String>,
    pub output_ref: String,
    /// Render-time pivot style bands (Excel keeps pivot styling out of cell
    /// xfs entirely), resolved from pivotTableStyleInfo by
    /// `pivot_style_palette`. Each band carries fill / font color / bold;
    /// absent = inherit the band below it (precedence, lowest first:
    /// wholeTable, row/column stripes, firstColumn, subheading, subtotal,
    /// header, totalRow).
    #[serde(flatten)]
    pub palette: PivotStylePalette,
    /// pivotTableStyleInfo carries a style name: band rows stay bold even
    /// when the resolved palette has no fills (Light 1-7, stripes off).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub styled: bool,
    /// location/@firstDataRow — rows above it inside the output ref are
    /// header rows.
    #[serde(skip_serializing_if = "is_zero")]
    pub first_data_row: usize,
    /// location/@firstDataCol — columns left of it are the row-label
    /// columns (Excel's "Row Headers" band). Always serialized: 0 (no row
    /// fields) must not fall back to the renderer's default of 1.
    pub first_data_col: usize,
    /// Always serialized: false must reach the renderer (a missing field
    /// falls back to Excel's default of true).
    pub row_grand_totals: bool,
    /// Row band per output row from firstDataRow down (see
    /// `visuals::PivotPartInfo::row_kinds`).
    #[serde(skip_serializing_if = "String::is_empty")]
    pub row_kinds: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnWidth {
    pub start_column: usize,
    pub end_column: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    pub hidden: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub collapsed: bool,
    /// <col style=>: the default xf for cells in this span that carry no own style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_index: Option<usize>,
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FreezePane {
    pub frozen_columns: usize,
    pub frozen_rows: usize,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RowProperty {
    pub row: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    /// customHeight="1": the user fixed this height; without it ht is just
    /// Excel's last auto-fit result and the row should keep auto-sizing.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub custom_height: bool,
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub collapsed: bool,
    /// <row s= customFormat="1">: the default xf for cells in this row that
    /// carry no own style. Rows beat columns, cells beat rows (OOXML order).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<usize>,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MergedRange {
    pub start_row: usize,
    pub start_column: usize,
    pub end_row: usize,
    pub end_column: usize,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableInfo {
    pub range: MergedRange,
    pub header_row_count: usize,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
    /// The table's autoFilter carries live criteria: Excel then re-ranks row
    /// stripes by visible order instead of physical row parity.
    #[serde(skip_serializing_if = "is_false")]
    pub filter_active: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub second_row_stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub second_column_stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_column_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_font_color: Option<String>,
    /// Rule Excel draws across the top of the totals band (thin/medium/double).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_border_style: Option<String>,
    /// Data-band text color (Dark families paint white on the solid body).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header_cell_font_color: Option<String>,
    /// table/@totalsRowCount — rows at the bottom styled as the totals band.
    #[serde(skip_serializing_if = "is_zero")]
    pub totals_row_count: usize,
    /// Style frame color (outline + header rule) for border-drawn families.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
    /// wholeTable borders (custom-style dxfs and the gridded builtin Light
    /// 15-21 block): outline + inner grid + header rule.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_border_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_horizontal_border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_horizontal_border_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_vertical_border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inner_vertical_border_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_bottom_border_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_bottom_border_style: Option<String>,
}

pub(crate) fn is_zero(value: &usize) -> bool {
    *value == 0
}

pub(crate) fn is_false(value: &bool) -> bool {
    !*value
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CommentInfo {
    pub row: usize,
    pub column: usize,
    pub author: String,
    pub text: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidationRule {
    pub ranges: Vec<MergedRange>,
    pub rule_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    pub formulas: Vec<String>,
    pub allow_blank: bool,
    /// Raw OOXML flag: "1" SUPPRESSES the in-cell dropdown (inverted name).
    pub suppress_dropdown: bool,
    pub show_input_message: bool,
    pub show_error_message: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    pub start_row: usize,
    pub end_row: usize,
    pub start_column: usize,
    pub end_column: usize,
}

impl CellRange {
    pub(crate) fn validate(&self, sheet: &SheetMetadata) -> Result<(), SidecarError> {
        if self.start_row > self.end_row || self.start_column > self.end_column {
            return Err(SidecarError::InvalidRequest(
                "Range boundaries are reversed.".into(),
            ));
        }
        if self.end_row >= sheet.row_count || self.end_column >= sheet.column_count {
            return Err(SidecarError::InvalidRequest(
                "Range is outside the worksheet.".into(),
            ));
        }
        let row_count = self.end_row - self.start_row + 1;
        let column_count = self.end_column - self.start_column + 1;
        if row_count.saturating_mul(column_count) > MAX_RANGE_CELLS {
            return Err(SidecarError::InvalidRequest(format!(
                "Range exceeds the {MAX_RANGE_CELLS} cell response limit."
            )));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(untagged)]
pub enum CellValue {
    String(String),
    Number(f64),
    Boolean(bool),
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRecord {
    pub row: usize,
    pub column: usize,
    pub value: Option<CellValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// `<f t="array" ref>`: the legacy CSE range this master formula fills.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub array_ref: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rich: Option<Vec<RichRun>>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RichRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    /// "subscript" | "superscript" (rPr <vertAlign val>).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

#[derive(Debug)]
pub struct SharedString {
    pub(crate) text: String,
    pub(crate) runs: Option<Vec<RichRun>>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RangeResult {
    pub cells: Vec<CellRecord>,
    pub rows: Vec<RowProperty>,
    pub merges: Vec<MergedRange>,
    pub hyperlinks: Vec<HyperlinkRecord>,
    /// Sheet-wide rules; empty until worksheet indexing completes.
    pub conditional_rules: Vec<ConditionalRule>,
    /// Also sheet-wide, complete-only.
    pub auto_filter: Option<MergedRange>,
    pub data_validations: Vec<DataValidationRule>,
    pub sheet_protection: Option<SheetProtectionInfo>,
    /// Manual page breaks (0-based index of the row/column after the break);
    /// sheet-wide, complete-only.
    pub row_breaks: Vec<usize>,
    pub col_breaks: Vec<usize>,
    /// protectedRanges entries; sheet-wide, complete-only.
    pub protected_ranges: Vec<ProtectedRangeInfo>,
    /// Saved print settings; sheet-wide, complete-only, None when the sheet
    /// declares none.
    pub page_setup: Option<PagePrintInfo>,
    pub indexed_through_row: Option<usize>,
    pub indexing_complete: bool,
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetProtectionInfo {
    pub protected: bool,
    /// password= (legacy) or algorithmName/hashValue (modern) present.
    pub has_password: bool,
}

/// Inches, from `<pageMargins>`.
#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageMarginsInfo {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

/// The sheet's saved print settings (printOptions / pageMargins / pageSetup /
/// sheetPr/pageSetUpPr / headerFooter), sanitized to Excel's valid ranges.
#[derive(Clone, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PagePrintInfo {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_to_width: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_to_height: Option<u32>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub fit_to_page: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margins: Option<PageMarginsInfo>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub print_gridlines: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub print_headings: bool,
    /// Excel-encoded odd header/footer text (&L/&C/&R sections, field codes).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<String>,
    /// headerFooter/@differentOddEven: even pages print the even_* texts.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub different_odd_even: bool,
    /// headerFooter/@differentFirst: page 1 prints the first_* texts.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub different_first: bool,
    /// headerFooter/@scaleWithDoc="0": the header/footer keeps its size
    /// instead of following the print scale (Excel's default is to follow).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub header_footer_fixed_size: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_header: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_footer: Option<String>,
    /// `&G` pictures declared by `<legacyDrawingHF>` (VML shapes LH/CH/RH/
    /// LF/CF/RF, with an EVEN/FIRST suffix for the page variants).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub header_footer_pictures: Vec<HeaderFooterPictureInfo>,
}

/// One header/footer picture slot; the bytes are fetched on demand through
/// read_media with `id`, like a visual.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterPictureInfo {
    /// Media lookup key for the read_media command, like a visual id.
    pub id: String,
    /// VML shape id, upper-cased: L/C/R × H/F plus an optional EVEN or
    /// FIRST suffix (`CH`, `LHFIRST`, `RFEVEN`).
    pub position: String,
    /// Declared picture size in points (the VML shape's style width/height).
    pub width_pt: f64,
    pub height_pt: f64,
    pub media_type: String,
    #[serde(skip_serializing)]
    pub media_path: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProtectedRangeInfo {
    pub name: String,
    pub sqref: String,
    pub has_password: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaCellsResult {
    pub cells: Vec<CellRecord>,
    pub indexing_complete: bool,
    /// True when the sheet has more formula cells than the response cap —
    /// the caller must treat the list as unusable for closure analysis.
    pub truncated: bool,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HyperlinkRecord {
    pub row: usize,
    pub column: usize,
    pub target: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalRule {
    pub ranges: Vec<MergedRange>,
    pub rule_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    pub formulas: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_index: Option<usize>,
    pub priority: i64,
    /// cfRule/@stopIfTrue: a matched rule stops lower-priority rules.
    pub stop_if_true: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,
    pub percent: bool,
    pub bottom: bool,
    pub cfvos: Vec<Cfvo>,
    pub colors: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub icon_set_name: Option<String>,
    pub icon_reverse: bool,
    pub show_value: bool,
    /// x14:dataBar negativeFillColor, merged from the worksheet extLst.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub negative_color: Option<String>,
    /// x14:dataBar/@negativeBarColorSameAsPositive (spec default true when
    /// the x14 twin exists; absent means no x14 twin).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub negative_same_as_positive: Option<bool>,
    /// x14:dataBar/@gradient; absent means the ECMA default (gradient fill).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<bool>,
    /// x14:dataBar/@axisPosition (automatic | middle | none); absent means
    /// no x14 twin or the spec default (automatic).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_position: Option<String>,
    /// x14:dataBar/axisColor resolved to #RRGGBB.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_color: Option<String>,
    /// Effective bar extents as a percentage of the cell width. Set on every
    /// dataBar rule: the 2006 element's minLength/maxLength (schema defaults
    /// 10/90), overridden by the x14 twin's attributes (defaults 0/100)
    /// whenever a twin exists.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub min_length: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max_length: Option<u32>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Cfvo {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gte: Option<bool>,
}

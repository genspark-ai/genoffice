use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{BufReader, BufWriter, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::Duration;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use serde::{Deserialize, Serialize};
use uuid::Uuid;
use zip::ZipArchive;

pub mod archive;
pub mod convert;
pub mod recalc;
mod shared_formulas;
mod visuals;

pub use visuals::{CellStyle, MediaResult, ThemeFonts, VisualObject};
use visuals::{ColorContext, SheetVisualSource};

const CHUNK_ROW_COUNT: usize = 256;
const MAX_RANGE_CELLS: usize = 20_000;
const RANGE_WAIT: Duration = Duration::from_millis(750);
/// Reads this far past the indexed row would only burn the full RANGE_WAIT;
/// answer immediately instead and let the caller poll.
const RANGE_WAIT_MAX_LAG_ROWS: usize = 16 * CHUNK_ROW_COUNT;
const MAX_FORMULA_CELLS: usize = 100_000;
const MAX_ENTRY_COUNT: usize = 10_000;

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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_protection: Option<WorkbookProtectionInfo>,
    /// workbookPr/@date1904: serial dates count from 1904-01-01.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub date1904: bool,
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
    pub column_widths: Vec<ColumnWidth>,
    pub default_row_height: Option<f64>,
    pub default_column_width: Option<f64>,
    pub freeze: Option<FreezePane>,
    pub hidden: bool,
    pub tab_color: Option<String>,
    pub show_grid_lines: bool,
    /// sheetView/@showFormulas: the sheet opens in formula view (#188).
    pub show_formulas: bool,
    /// sheetView/@showRowColHeaders: row/column heading strips.
    pub show_row_col_headers: bool,
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
    /// Render-time pivot style band fill (header + grand-total rows); Excel
    /// keeps pivot styling out of cell xfs entirely.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_fill: Option<String>,
    /// location/@firstDataRow — rows above it inside the output ref are
    /// header rows.
    #[serde(skip_serializing_if = "is_zero")]
    pub first_data_row: usize,
    /// Always serialized: false must reach the renderer (a missing field
    /// falls back to Excel's default of true).
    pub row_grand_totals: bool,
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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header_cell_font_color: Option<String>,
    /// table/@totalsRowCount — rows at the bottom styled as the totals band.
    #[serde(skip_serializing_if = "is_zero")]
    pub totals_row_count: usize,
    /// Style frame color (outline + header rule) for border-drawn families.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
}

fn is_zero(value: &usize) -> bool {
    *value == 0
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
    fn validate(&self, sheet: &SheetMetadata) -> Result<(), SidecarError> {
        if self.start_row > self.end_row || self.start_column > self.end_column {
            return Err(SidecarError::InvalidRequest("Range boundaries are reversed.".into()));
        }
        if self.end_row >= sheet.row_count || self.end_column >= sheet.column_count {
            return Err(SidecarError::InvalidRequest("Range is outside the worksheet.".into()));
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
}

#[derive(Debug)]
pub struct SharedString {
    text: String,
    runs: Option<Vec<RichRun>>,
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

#[derive(Debug, Default, Deserialize, Serialize)]
struct ChunkData {
    cells: Vec<CellRecord>,
    rows: Vec<RowProperty>,
}

#[derive(Debug)]
pub enum SidecarError {
    InvalidRequest(String),
    Io(String),
    Workbook(String),
}

impl std::fmt::Display for SidecarError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidRequest(message) | Self::Io(message) | Self::Workbook(message) => {
                formatter.write_str(message)
            }
        }
    }
}

impl std::error::Error for SidecarError {}

impl From<std::io::Error> for SidecarError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error.to_string())
    }
}

impl From<zip::result::ZipError> for SidecarError {
    fn from(error: zip::result::ZipError) -> Self {
        Self::Workbook(error.to_string())
    }
}

impl From<quick_xml::Error> for SidecarError {
    fn from(error: quick_xml::Error) -> Self {
        Self::Workbook(error.to_string())
    }
}

impl From<serde_json::Error> for SidecarError {
    fn from(error: serde_json::Error) -> Self {
        Self::Io(error.to_string())
    }
}

pub struct WorkbookSessions {
    sessions: HashMap<String, WorkbookSession>,
}

impl WorkbookSessions {
    pub fn new() -> Self {
        Self {
            sessions: HashMap::new(),
        }
    }

    pub fn open(&mut self, path: &Path) -> Result<WorkbookMetadata, SidecarError> {
        self.open_with_locale(path, "zh")
    }

    pub fn open_with_locale(
        &mut self,
        path: &Path,
        locale: &str,
    ) -> Result<WorkbookMetadata, SidecarError> {
        let canonical_path = path.canonicalize()?;
        let file = File::open(&canonical_path)?;
        let mut archive = ZipArchive::new(file)?;
        validate_archive(&mut archive)?;
        let entry_count = archive.len();
        let color_context = visuals::read_theme_palette(&mut archive)?;
        let theme_colors = color_context.palette_hex();
        let theme_fonts = visuals::read_theme_fonts(&mut archive)?;
        let shared_strings = Arc::new(read_shared_strings(&mut archive, &color_context)?);
        let (declarations, workbook_protection, active_tab, date_1904) =
            read_sheet_declarations(&mut archive)?;
        let relationships = read_workbook_relationships(&mut archive)?;
        let (styles, dxf_styles) =
            visuals::read_styles(&mut archive, &color_context, theme_fonts.as_ref(), locale)?;
        let custom_table_styles = read_custom_table_styles(&mut archive, &dxf_styles);
        let mut sheets = Vec::with_capacity(declarations.len());
        let mut runtimes = Vec::with_capacity(declarations.len());
        let mut visual_sources = Vec::with_capacity(declarations.len());

        for declaration in declarations {
            let target = relationships.get(&declaration.relationship_id).ok_or_else(|| {
                SidecarError::Workbook(format!(
                    "Missing worksheet relationship {}.",
                    declaration.relationship_id
                ))
            })?;
            let worksheet_path = normalize_worksheet_path(target)?;
            let dimensions = read_sheet_dimensions(&mut archive, &worksheet_path, &color_context)?;
            let tables = read_sheet_tables(
                &mut archive,
                &worksheet_path,
                &color_context,
                &custom_table_styles,
            )?;
            let comments = visuals::read_comments(&mut archive, &worksheet_path)?
                .into_iter()
                .filter_map(|(reference, author, text)| {
                    let anchor = reference.split(':').next().unwrap_or(&reference);
                    let (row, column) = parse_address(&anchor.replace('$', "")).ok()?;
                    Some(CommentInfo {
                        row,
                        column,
                        author,
                        text,
                    })
                })
                .collect();
            let sparklines = read_sheet_sparklines(&mut archive, &worksheet_path)?;
            let pivot_infos = visuals::read_pivot_tables(&mut archive, &worksheet_path)?;
            let pivot_ranges = pivot_infos
                .iter()
                .filter_map(|info| parse_range_reference(&info.output_ref))
                .collect();
            let pivot_tables = pivot_infos
                .into_iter()
                .map(|info| PivotTableInfo {
                    header_fill: pivot_style_fill(info.style_name.as_deref(), &color_context),
                    path: info.path,
                    cache_path: info.cache_path,
                    output_ref: info.output_ref,
                    first_data_row: info.first_data_row,
                    row_grand_totals: info.row_grand_totals,
                })
                .collect();
            let id = format!("sheet-{}", declaration.sheet_id);
            sheets.push(SheetMetadata {
                id: id.clone(),
                name: declaration.name,
                row_count: dimensions.row_count,
                column_count: dimensions.column_count,
                column_widths: dimensions.column_widths,
                default_row_height: dimensions.default_row_height,
                default_column_width: dimensions.default_column_width,
                freeze: dimensions.freeze,
                hidden: declaration.hidden,
                tab_color: dimensions.tab_color,
                show_grid_lines: dimensions.show_grid_lines,
                show_formulas: dimensions.show_formulas,
                show_row_col_headers: dimensions.show_row_col_headers,
                tables,
                comments,
                pivot_ranges,
                pivot_tables,
                sparklines,
            });
            visual_sources.push(SheetVisualSource {
                sheet_id: id,
                worksheet_path: worksheet_path.clone(),
            });
            runtimes.push(SheetRuntime::new(worksheet_path));
        }
        if sheets.is_empty() {
            return Err(SidecarError::Workbook(
                "Workbook contains no readable worksheets.".into(),
            ));
        }
        // Value-less cells are kept when their xf paints something visible
        // (fill/border) or differs from the default xf in number format,
        // font or alignment — formatting a user's future input inherits (#169).
        let default_style = styles.first().cloned().unwrap_or_default();
        let styled_xfs: Arc<Vec<bool>> = Arc::new(
            styles
                .iter()
                .map(|style| style.styles_blank_cell(&default_style))
                .collect(),
        );
        let visual_objects =
            visuals::read_visual_objects(&mut archive, &visual_sources, &color_context)?;
        let defined_names = read_defined_names(&mut archive)?;

        let session_id = Uuid::new_v4().to_string();
        let cache_directory =
            std::env::temp_dir().join(format!("genspark-ai-excel-{session_id}"));
        fs::create_dir(&cache_directory)?;
        let name = canonical_path
            .file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("workbook.xlsx")
            .to_owned();
        self.sessions.insert(
            session_id.clone(),
            WorkbookSession {
                path: canonical_path,
                sheets: sheets.clone(),
                runtimes,
                shared_strings,
                styled_xfs,
                color_context: Arc::new(color_context),
                visuals: visual_objects.clone(),
                cache_directory,
                cancelled: Arc::new(AtomicBool::new(false)),
            },
        );
        Ok(WorkbookMetadata {
            session_id,
            name,
            entry_count,
            sheets,
            active_tab,
            styles,
            dxf_styles,
            visuals: visual_objects,
            defined_names,
            theme_colors,
            theme_fonts,
            workbook_protection,
            date1904: date_1904,
        })
    }

    pub fn read_range(
        &mut self,
        session_id: &str,
        sheet_id: &str,
        range: &CellRange,
    ) -> Result<RangeResult, SidecarError> {
        let session = self
            .sessions
            .get_mut(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let sheet_index = session
            .sheets
            .iter()
            .position(|sheet| sheet.id == sheet_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown worksheet.".into()))?;
        range.validate(&session.sheets[sheet_index])?;
        session.ensure_parser(sheet_index)?;
        session.read_range(sheet_index, range)
    }

    pub fn read_formula_cells(
        &mut self,
        session_id: &str,
        sheet_id: &str,
    ) -> Result<FormulaCellsResult, SidecarError> {
        let session = self
            .sessions
            .get_mut(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let sheet_index = session
            .sheets
            .iter()
            .position(|sheet| sheet.id == sheet_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown worksheet.".into()))?;
        session.ensure_parser(sheet_index)?;
        session.read_formula_cells(sheet_index)
    }

    pub fn close(&mut self, session_id: &str) -> Result<(), SidecarError> {
        let session = self
            .sessions
            .remove(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        session.close()
    }

    pub fn cancel(&self, session_id: &str) -> Result<(), SidecarError> {
        let session = self
            .sessions
            .get(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        session.cancelled.store(true, Ordering::Release);
        Ok(())
    }

    pub fn read_media(
        &self,
        session_id: &str,
        visual_id: &str,
    ) -> Result<MediaResult, SidecarError> {
        let session = self
            .sessions
            .get(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let media_path = session
            .visuals
            .iter()
            .find(|visual| visual.id == visual_id)
            .and_then(|visual| visual.media_path.as_deref())
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook image.".into()))?;
        let file = File::open(&session.path)?;
        let mut archive = ZipArchive::new(file)?;
        visuals::read_media(&mut archive, media_path)
    }
}

impl Default for WorkbookSessions {
    fn default() -> Self {
        Self::new()
    }
}

impl Drop for WorkbookSessions {
    fn drop(&mut self) {
        for (_, session) in self.sessions.drain() {
            let _ = session.close();
        }
    }
}

struct WorkbookSession {
    path: PathBuf,
    sheets: Vec<SheetMetadata>,
    runtimes: Vec<SheetRuntime>,
    shared_strings: Arc<Vec<SharedString>>,
    styled_xfs: Arc<Vec<bool>>,
    color_context: Arc<ColorContext>,
    visuals: Vec<VisualObject>,
    cache_directory: PathBuf,
    cancelled: Arc<AtomicBool>,
}

impl WorkbookSession {
    fn ensure_parser(&mut self, sheet_index: usize) -> Result<(), SidecarError> {
        let runtime = &mut self.runtimes[sheet_index];
        if runtime.handle.is_some() {
            return Ok(());
        }
        let path = self.path.clone();
        let worksheet_path = runtime.worksheet_path.clone();
        let cache_directory = self.cache_directory.clone();
        let state = Arc::clone(&runtime.state);
        let shared_strings = Arc::clone(&self.shared_strings);
        let styled_xfs = Arc::clone(&self.styled_xfs);
        let color_context = Arc::clone(&self.color_context);
        let cancelled = Arc::clone(&self.cancelled);
        let handle = thread::Builder::new()
            .name(format!("xlsx-index-{sheet_index}"))
            .spawn(move || {
                let result = index_worksheet(
                    &path,
                    &worksheet_path,
                    sheet_index,
                    &cache_directory,
                    &shared_strings,
                    &styled_xfs,
                    &color_context,
                    &state,
                    &cancelled,
                );
                let (lock, condition) = &*state;
                if let Ok(mut index) = lock.lock() {
                    match result {
                        Ok(()) => index.complete = true,
                        Err(error) => index.error = Some(error.to_string()),
                    }
                    condition.notify_all();
                }
            })?;
        runtime.handle = Some(handle);
        Ok(())
    }

    fn read_range(
        &self,
        sheet_index: usize,
        range: &CellRange,
    ) -> Result<RangeResult, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, condition) = &*runtime.state;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        let far_ahead = index.error.is_none()
            && !index.complete
            && index
                .indexed_through_row
                .map_or(range.start_row, |row| range.start_row.saturating_sub(row))
                >= RANGE_WAIT_MAX_LAG_ROWS;
        let index = if far_ahead {
            index
        } else {
            condition
                .wait_timeout_while(index, RANGE_WAIT, |current| {
                    current.error.is_none()
                        && !current.complete
                        && current.indexed_through_row.is_none_or(|row| row < range.end_row)
                })
                .map_err(|_| SidecarError::Io("Worksheet index wait failed.".into()))?
                .0
        };
        if let Some(error) = &index.error {
            return Err(SidecarError::Workbook(error.clone()));
        }
        let indexed_through_row = index.indexed_through_row;
        let indexing_complete = index.complete;
        let merges = index
            .merges
            .iter()
            .filter(|merge| {
                merge.start_row <= range.end_row
                    && merge.end_row >= range.start_row
                    && merge.start_column <= range.end_column
                    && merge.end_column >= range.start_column
            })
            .copied()
            .collect::<Vec<_>>();
        let hyperlinks = index
            .hyperlinks
            .iter()
            .filter(|link| {
                link.row >= range.start_row
                    && link.row <= range.end_row
                    && link.column >= range.start_column
                    && link.column <= range.end_column
            })
            .cloned()
            .collect::<Vec<_>>();
        let conditional_rules = if indexing_complete {
            index.conditional_rules.clone()
        } else {
            Vec::new()
        };
        let auto_filter = if indexing_complete {
            index.auto_filter
        } else {
            None
        };
        let data_validations = if indexing_complete {
            index.data_validations.clone()
        } else {
            Vec::new()
        };
        let sheet_protection = if indexing_complete {
            index.sheet_protection
        } else {
            None
        };
        let row_breaks = if indexing_complete {
            index.row_breaks.clone()
        } else {
            Vec::new()
        };
        let col_breaks = if indexing_complete {
            index.col_breaks.clone()
        } else {
            Vec::new()
        };
        let protected_ranges = if indexing_complete {
            index.protected_ranges.clone()
        } else {
            Vec::new()
        };
        let first_chunk = range.start_row / CHUNK_ROW_COUNT;
        let last_available_row = indexed_through_row
            .map(|row| row.min(range.end_row))
            .unwrap_or(0);
        let last_chunk = last_available_row / CHUNK_ROW_COUNT;
        let paths = if indexed_through_row.is_none() || last_available_row < range.start_row {
            Vec::new()
        } else {
            (first_chunk..=last_chunk)
                .filter_map(|chunk| index.chunk_files.get(&chunk).cloned())
                .collect::<Vec<_>>()
        };
        drop(index);

        let mut cells = Vec::new();
        let mut rows = Vec::new();
        for path in paths {
            let file = File::open(path)?;
            let chunk: ChunkData = serde_json::from_reader(BufReader::new(file))?;
            cells.extend(chunk.cells.into_iter().filter(|cell| {
                cell.row >= range.start_row
                    && cell.row <= range.end_row
                    && cell.column >= range.start_column
                    && cell.column <= range.end_column
            }));
            rows.extend(
                chunk
                    .rows
                    .into_iter()
                    .filter(|row| row.row >= range.start_row && row.row <= range.end_row),
            );
        }
        Ok(RangeResult {
            cells,
            rows,
            merges,
            hyperlinks,
            conditional_rules,
            auto_filter,
            data_validations,
            sheet_protection,
            row_breaks,
            col_breaks,
            protected_ranges,
            indexed_through_row,
            indexing_complete,
        })
    }

    /// All formula cells indexed so far; the closure analyzer polls until
    /// `indexing_complete`. Capped — a sheet over the cap sets `truncated`.
    fn read_formula_cells(&self, sheet_index: usize) -> Result<FormulaCellsResult, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, _) = &*runtime.state;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        if let Some(error) = &index.error {
            return Err(SidecarError::Workbook(error.clone()));
        }
        Ok(FormulaCellsResult {
            cells: index.formula_cells.clone(),
            indexing_complete: index.complete,
            truncated: index.formula_truncated,
        })
    }

    fn close(mut self) -> Result<(), SidecarError> {
        self.cancelled.store(true, Ordering::Release);
        for runtime in &self.runtimes {
            let (_, condition) = &*runtime.state;
            condition.notify_all();
        }
        for runtime in &mut self.runtimes {
            if let Some(handle) = runtime.handle.take() {
                let _ = handle.join();
            }
        }
        if self.cache_directory.exists() {
            fs::remove_dir_all(&self.cache_directory)?;
        }
        Ok(())
    }
}

struct SheetRuntime {
    worksheet_path: String,
    state: Arc<(Mutex<SheetIndex>, Condvar)>,
    handle: Option<JoinHandle<()>>,
}

impl SheetRuntime {
    fn new(worksheet_path: String) -> Self {
        Self {
            worksheet_path,
            state: Arc::new((Mutex::new(SheetIndex::default()), Condvar::new())),
            handle: None,
        }
    }
}

#[derive(Default)]
struct SheetIndex {
    indexed_through_row: Option<usize>,
    complete: bool,
    error: Option<String>,
    chunk_files: HashMap<usize, PathBuf>,
    merges: Vec<MergedRange>,
    hyperlinks: Vec<HyperlinkRecord>,
    conditional_rules: Vec<ConditionalRule>,
    auto_filter: Option<MergedRange>,
    data_validations: Vec<DataValidationRule>,
    sheet_protection: Option<SheetProtectionInfo>,
    row_breaks: Vec<usize>,
    col_breaks: Vec<usize>,
    protected_ranges: Vec<ProtectedRangeInfo>,
    /// Formula cells collected while indexing, capped at MAX_FORMULA_CELLS.
    formula_cells: Vec<CellRecord>,
    formula_truncated: bool,
}

#[derive(Debug)]
struct SheetDeclaration {
    name: String,
    sheet_id: String,
    relationship_id: String,
    hidden: bool,
}

fn validate_archive(archive: &mut ZipArchive<File>) -> Result<(), SidecarError> {
    if archive.len() > MAX_ENTRY_COUNT {
        return Err(SidecarError::Workbook(
            "Workbook contains too many ZIP entries.".into(),
        ));
    }
    for index in 0..archive.len() {
        let entry = archive.by_index(index)?;
        if entry.enclosed_name().is_none() {
            return Err(SidecarError::Workbook(
                "Workbook contains an unsafe ZIP path.".into(),
            ));
        }
    }
    Ok(())
}

fn read_sheet_declarations(
    archive: &mut ZipArchive<File>,
) -> Result<(Vec<SheetDeclaration>, Option<WorkbookProtectionInfo>, usize, bool), SidecarError> {
    let xml = read_zip_string(archive, "xl/workbook.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut sheets = Vec::new();
    let mut protection = None;
    let mut active_tab = 0usize;
    let mut date_1904 = false;
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"workbookPr" =>
            {
                date_1904 = attribute_value(&reader, &element, b"date1904")?
                    .is_some_and(|value| value == "1" || value == "true");
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"workbookView" =>
            {
                if let Some(value) = attribute_value(&reader, &element, b"activeTab")? {
                    active_tab = value.parse::<usize>().unwrap_or(0);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"workbookProtection" =>
            {
                let lock_structure = attribute_value(&reader, &element, b"lockStructure")?
                    .is_some_and(|value| value == "1" || value == "true");
                let has_password =
                    attribute_value(&reader, &element, b"workbookPassword")?.is_some()
                        || attribute_value(&reader, &element, b"workbookHashValue")?.is_some();
                protection = Some(WorkbookProtectionInfo {
                    lock_structure,
                    has_password,
                });
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"sheet" =>
            {
                let name = attribute_value(&reader, &element, b"name")?
                    .ok_or_else(|| SidecarError::Workbook("Sheet has no name.".into()))?;
                let sheet_id = attribute_value(&reader, &element, b"sheetId")?
                    .ok_or_else(|| SidecarError::Workbook("Sheet has no sheetId.".into()))?;
                let relationship_id = attribute_value(&reader, &element, b"id")?
                    .ok_or_else(|| SidecarError::Workbook("Sheet has no relationship id.".into()))?;
                let hidden = matches!(
                    attribute_value(&reader, &element, b"state")?.as_deref(),
                    Some("hidden") | Some("veryHidden")
                );
                sheets.push(SheetDeclaration {
                    name,
                    sheet_id,
                    relationship_id,
                    hidden,
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((sheets, protection, active_tab, date_1904))
}

fn read_workbook_relationships(
    archive: &mut ZipArchive<File>,
) -> Result<HashMap<String, String>, SidecarError> {
    let xml = read_zip_string(archive, "xl/_rels/workbook.xml.rels")?;
    let mut reader = Reader::from_str(&xml);
    let mut relationships = HashMap::new();
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                if let (Some(id), Some(target)) = (
                    attribute_value(&reader, &element, b"Id")?,
                    attribute_value(&reader, &element, b"Target")?,
                ) {
                    relationships.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(relationships)
}

// Pure string handling: zip entry names always use '/', while PathBuf joins
// with '\' on Windows, which made by_name miss every worksheet there.
fn normalize_worksheet_path(target: &str) -> Result<String, SidecarError> {
    let candidate = if let Some(absolute) = target.strip_prefix('/') {
        absolute.to_owned()
    } else if target.starts_with("xl/") {
        target.to_owned()
    } else {
        format!("xl/{}", target.trim_start_matches("./"))
    };
    let mut normalized: Vec<&str> = Vec::new();
    for component in candidate.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                if normalized.pop().is_none() {
                    return Err(SidecarError::Workbook(
                        "Worksheet relationship escapes the package.".into(),
                    ));
                }
            }
            value => normalized.push(value),
        }
    }
    Ok(normalized.join("/"))
}

struct SheetDimensions {
    row_count: usize,
    column_count: usize,
    column_widths: Vec<ColumnWidth>,
    default_row_height: Option<f64>,
    default_column_width: Option<f64>,
    freeze: Option<FreezePane>,
    tab_color: Option<String>,
    show_grid_lines: bool,
    show_formulas: bool,
    show_row_col_headers: bool,
}

fn read_sheet_dimensions(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    colors: &ColorContext,
) -> Result<SheetDimensions, SidecarError> {
    let entry = archive.by_name(worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut maximum_row = 0;
    let mut maximum_column = 0;
    // Implicit-address fallbacks, mirroring the cell-stream parser.
    let mut current_row = 0usize;
    let mut first_row = true;
    let mut next_column = 0usize;
    let mut dimensions = None;
    let mut column_widths = Vec::new();
    let mut default_row_height = None;
    let mut default_column_width = None;
    let mut freeze = None;
    let mut tab_color = None;
    let mut show_grid_lines = true;
    let mut show_formulas = false;
    let mut show_row_col_headers = true;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"dimension" =>
            {
                if let Some(reference) = attribute_value(&reader, &element, b"ref")? {
                    // A malformed ref means "no usable dimension", not a broken file.
                    dimensions = dimensions_from_reference(&reference).ok();
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"tabColor" =>
            {
                tab_color = visuals::resolve_color(
                    attribute_value(&reader, &element, b"rgb")?.as_deref(),
                    attribute_value(&reader, &element, b"indexed")?.as_deref(),
                    attribute_value(&reader, &element, b"theme")?.as_deref(),
                    attribute_value(&reader, &element, b"tint")?.as_deref(),
                    colors,
                );
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"sheetView" =>
            {
                if let Some(value) = attribute_value(&reader, &element, b"showGridLines")? {
                    show_grid_lines = value != "0" && value != "false";
                }
                if let Some(value) = attribute_value(&reader, &element, b"showFormulas")? {
                    show_formulas = value == "1" || value == "true";
                }
                if let Some(value) = attribute_value(&reader, &element, b"showRowColHeaders")? {
                    show_row_col_headers = value != "0" && value != "false";
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"sheetFormatPr" =>
            {
                // defaultColWidth="0" is legal (Excel writes it with
                // zeroHeight="1" on all-hidden sheets) and means "no default —
                // use the built-in width"; normalize 0 to None (#173).
                default_row_height = attribute_value(&reader, &element, b"defaultRowHeight")?
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| *value > 0.0);
                default_column_width = attribute_value(&reader, &element, b"defaultColWidth")?
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| *value > 0.0);
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"pane" =>
            {
                let state = attribute_value(&reader, &element, b"state")?;
                if matches!(state.as_deref(), Some("frozen") | Some("frozenSplit")) {
                    let split = |name: &[u8]| -> Result<usize, SidecarError> {
                        Ok(attribute_value(&reader, &element, name)?
                            .and_then(|value| value.parse::<f64>().ok())
                            .map(|value| value as usize)
                            .unwrap_or(0))
                    };
                    freeze = Some(FreezePane {
                        frozen_columns: split(b"xSplit")?,
                        frozen_rows: split(b"ySplit")?,
                    });
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"col" =>
            {
                let minimum = attribute_value(&reader, &element, b"min")?
                    .and_then(|value| value.parse::<usize>().ok());
                let maximum = attribute_value(&reader, &element, b"max")?
                    .and_then(|value| value.parse::<usize>().ok());
                let width = attribute_value(&reader, &element, b"width")?
                    .and_then(|value| value.parse::<f64>().ok());
                let hidden = attribute_value(&reader, &element, b"hidden")?
                    .is_some_and(|value| value == "1" || value == "true");
                let outline_level = attribute_value(&reader, &element, b"outlineLevel")?
                    .and_then(|value| value.parse::<u8>().ok())
                    .map(|value| value.min(7))
                    .filter(|value| *value > 0);
                let collapsed = attribute_value(&reader, &element, b"collapsed")?
                    .is_some_and(|value| value == "1" || value == "true");
                // xf 0 is the workbook default; carrying it would only bloat the metadata.
                let style_index = attribute_value(&reader, &element, b"style")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|value| *value > 0);
                if let (Some(minimum), Some(maximum)) = (minimum, maximum) {
                    if width.is_some()
                        || hidden
                        || outline_level.is_some()
                        || collapsed
                        || style_index.is_some()
                    {
                        column_widths.push(ColumnWidth {
                            start_column: minimum.saturating_sub(1),
                            end_column: maximum.saturating_sub(1),
                            width,
                            hidden,
                            outline_level,
                            collapsed,
                            style_index,
                        });
                    }
                }
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"sheetData"
                    // Single-row/column dimensions are the signature of writers
                    // that never update <dimension> (POI SXSSF, OpenXML SDK
                    // producers emitting A1 or A1:G1); measure instead.
                    && dimensions.is_some_and(|(rows, columns)| rows > 1 && columns > 1) =>
            {
                let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                return Ok(SheetDimensions {
                    row_count,
                    column_count,
                    column_widths,
                    default_row_height,
                    default_column_width,
                    freeze,
                    tab_color,
                    show_grid_lines,
                    show_formulas,
                    show_row_col_headers,
                });
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"row" =>
            {
                current_row = attribute_value(&reader, &element, b"r")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|value| *value > 0)
                    .map(|value| value - 1)
                    .unwrap_or(if first_row { 0 } else { current_row + 1 });
                first_row = false;
                next_column = 0;
                maximum_row = maximum_row.max(current_row);
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"c" =>
            {
                let (row, column) = match attribute_value(&reader, &element, b"r")? {
                    Some(address) => parse_address(&address)?,
                    None => (current_row, next_column),
                };
                next_column = column + 1;
                maximum_row = maximum_row.max(row);
                maximum_column = maximum_column.max(column);
            }
            Event::Eof => {
                let (dim_rows, dim_columns) = dimensions.unwrap_or((0, 0));
                let (row_count, column_count) = (
                    dim_rows.max(maximum_row + 1),
                    dim_columns.max(maximum_column + 1),
                );
                return Ok(SheetDimensions {
                    row_count,
                    column_count,
                    column_widths,
                    default_row_height,
                    default_column_width,
                    freeze,
                    tab_color,
                    show_grid_lines,
                    show_formulas,
                    show_row_col_headers,
                });
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn dimensions_from_reference(reference: &str) -> Result<(usize, usize), SidecarError> {
    let last = reference
        .split(':')
        .next_back()
        .unwrap_or(reference)
        .replace('$', "");
    let (row, column) = parse_address(&last)?;
    Ok((row + 1, column + 1))
}

fn read_shared_strings(
    archive: &mut ZipArchive<File>,
    colors: &ColorContext,
) -> Result<Vec<SharedString>, SidecarError> {
    let Ok(entry) = archive.by_name("xl/sharedStrings.xml") else {
        return Ok(Vec::new());
    };
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut strings = Vec::new();
    let mut current = String::new();
    let mut runs: Vec<RichRun> = Vec::new();
    let mut current_run: Option<RichRun> = None;
    let mut in_text = false;
    let mut in_phonetic = false;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if element.local_name().as_ref() == b"si" => {
                current.clear();
                runs.clear();
                current_run = None;
            }
            Event::Start(element) if element.local_name().as_ref() == b"r" => {
                current_run = Some(RichRun::default());
            }
            // Phonetic guide text (<rPh><t>…</t></rPh>) is ruby annotation,
            // not cell content.
            Event::Start(element) if element.local_name().as_ref() == b"rPh" => {
                in_phonetic = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"rPh" => {
                in_phonetic = false;
            }
            Event::Start(element) | Event::Empty(element) if current_run.is_some() => {
                if element.local_name().as_ref() == b"t" {
                    in_text = true;
                } else if let Some(run) = current_run.as_mut() {
                    apply_run_property(run, &reader, &element, colors)?;
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"t" => {
                in_text = !in_phonetic;
            }
            Event::Text(text) if in_text => {
                let decoded = decode_text(&text)?;
                current.push_str(&decoded);
                if let Some(run) = &mut current_run {
                    run.text.push_str(&decoded);
                }
            }
            Event::GeneralRef(reference) if in_text => {
                let decoded = general_ref_text(&reference)?;
                current.push_str(&decoded);
                if let Some(run) = &mut current_run {
                    run.text.push_str(&decoded);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"t" => in_text = false,
            Event::End(element) if element.local_name().as_ref() == b"r" => {
                if let Some(run) = current_run.take() {
                    runs.push(run);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"si" => {
                strings.push(SharedString {
                    text: current.clone(),
                    runs: qualify_runs(std::mem::take(&mut runs)),
                });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(strings)
}

fn has_run_formatting(run: &RichRun) -> bool {
    run.bold
        || run.italic
        || run.underline
        || run.strikethrough
        || run.color.is_some()
        || run.size.is_some()
        || run.family.is_some()
}

fn qualify_runs(runs: Vec<RichRun>) -> Option<Vec<RichRun>> {
    (runs.len() > 1 || runs.iter().any(has_run_formatting)).then_some(runs)
}

/// Applies one rPr child element (b/i/u/strike/sz/rFont/color) to a run.
fn apply_run_property<R: std::io::BufRead>(
    run: &mut RichRun,
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    colors: &ColorContext,
) -> Result<(), SidecarError> {
    // CT_BooleanProperty: presence means true unless val says otherwise —
    // POI writes explicit off flags as non-self-closing <strike val="0">.
    let flag_on = |value: Option<String>| !matches!(value.as_deref(), Some("0") | Some("false"));
    match element.local_name().as_ref() {
        b"b" => run.bold = flag_on(attribute_value(reader, element, b"val")?),
        b"i" => run.italic = flag_on(attribute_value(reader, element, b"val")?),
        b"u" => {
            run.underline =
                attribute_value(reader, element, b"val")?.as_deref() != Some("none");
        }
        b"strike" => run.strikethrough = flag_on(attribute_value(reader, element, b"val")?),
        b"sz" => {
            run.size = attribute_value(reader, element, b"val")?
                .and_then(|value| value.parse::<f64>().ok());
        }
        b"rFont" => run.family = attribute_value(reader, element, b"val")?,
        b"color" => {
            run.color = visuals::resolve_color(
                attribute_value(reader, element, b"rgb")?.as_deref(),
                attribute_value(reader, element, b"indexed")?.as_deref(),
                attribute_value(reader, element, b"theme")?.as_deref(),
                attribute_value(reader, element, b"tint")?.as_deref(),
                colors,
            );
        }
        _ => {}
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn index_worksheet(
    workbook_path: &Path,
    worksheet_path: &str,
    sheet_index: usize,
    cache_directory: &Path,
    shared_strings: &[SharedString],
    styled_xfs: &[bool],
    colors: &ColorContext,
    state: &Arc<(Mutex<SheetIndex>, Condvar)>,
    cancelled: &AtomicBool,
) -> Result<(), SidecarError> {
    let file = File::open(workbook_path)?;
    let mut archive = ZipArchive::new(file)?;
    let link_targets = visuals::hyperlink_targets(&mut archive, worksheet_path)?;
    let entry = archive.by_name(worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut cell_builder: Option<CellBuilder> = None;
    // <row r=> and <c r=> are both optional: a missing row is one below its
    // predecessor, a missing cell one right of its predecessor (54288.xlsx
    // omits r on all but each row's first cell).
    let mut current_row = 0usize;
    let mut first_row = true;
    let mut next_column = 0usize;
    let mut in_formula = false;
    let mut in_value = false;
    let mut in_text = false;
    let mut in_phonetic = false;
    // Shared-formula groups: the master's text expands into every follower (#165).
    let mut shared_formulas = shared_formulas::SharedFormulas::default();
    let mut cell_shared_si: Option<u32> = None;
    let mut chunk_index = 0;
    let mut chunk = ChunkData::default();
    let mut pending_formulas: Vec<CellRecord> = Vec::new();
    let mut latest_row = 0;
    let mut merges = Vec::new();
    let mut hyperlinks = Vec::new();
    let mut cf_ranges: Vec<MergedRange> = Vec::new();
    let mut cf_rule: Option<ConditionalRule> = None;
    let mut in_cf_formula = false;
    let mut conditional_rules: Vec<ConditionalRule> = Vec::new();
    let mut in_cf_ext_id = false;
    let mut cf_rule_ext_id: Option<String> = None;
    let mut x14_rule_slots: HashMap<String, usize> = HashMap::new();
    let mut x14_current_id: Option<String> = None;
    let mut x14_negative_colors: HashMap<String, String> = HashMap::new();
    let mut x14_databars: HashMap<String, (Option<bool>, bool)> = HashMap::new();
    let mut auto_filter: Option<MergedRange> = None;
    let mut sheet_protection: Option<SheetProtectionInfo> = None;
    let mut row_breaks: Vec<usize> = Vec::new();
    let mut col_breaks: Vec<usize> = Vec::new();
    let mut in_row_breaks = false;
    let mut in_col_breaks = false;
    let mut protected_ranges: Vec<ProtectedRangeInfo> = Vec::new();
    let mut dv_rule: Option<DataValidationRule> = None;
    let mut in_dv_formula = false;
    let mut data_validations: Vec<DataValidationRule> = Vec::new();

    loop {
        if cancelled.load(Ordering::Acquire) {
            return Ok(());
        }
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"row" =>
            {
                current_row = attribute_value(&reader, &element, b"r")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|value| *value > 0)
                    .map(|value| value - 1)
                    .unwrap_or(if first_row { 0 } else { current_row + 1 });
                first_row = false;
                next_column = 0;
                if let Some(property) = row_property(&reader, &element, current_row)? {
                    let row_chunk = property.row / CHUNK_ROW_COUNT;
                    if row_chunk != chunk_index {
                        flush_chunk(
                            sheet_index,
                            chunk_index,
                            &mut chunk,
                            &mut pending_formulas,
                            cache_directory,
                            state,
                            row_chunk * CHUNK_ROW_COUNT - 1,
                        )?;
                        chunk_index = row_chunk;
                    }
                    chunk.rows.push(property);
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"c" => {
                let builder =
                    CellBuilder::from_element(&reader, &element, current_row, next_column)?;
                next_column = builder.column + 1;
                cell_builder = Some(builder);
            }
            Event::Empty(element) if element.local_name().as_ref() == b"c" => {
                // Self-closing cells carry no value but may carry a style
                // (borders, fills); keep them like any other cell.
                {
                    let builder =
                        CellBuilder::from_element(&reader, &element, current_row, next_column)?;
                    next_column = builder.column + 1;
                    latest_row = latest_row.max(builder.row);
                    let cell_chunk = builder.row / CHUNK_ROW_COUNT;
                    if cell_chunk != chunk_index {
                        flush_chunk(
                            sheet_index,
                            chunk_index,
                            &mut chunk,
                            &mut pending_formulas,
                            cache_directory,
                            state,
                            cell_chunk * CHUNK_ROW_COUNT - 1,
                        )?;
                        chunk_index = cell_chunk;
                    }
                    if let Some(cell) = builder.finish(shared_strings, styled_xfs)? {
                        chunk.cells.push(cell);
                    }
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"f" => {
                in_formula = true;
                cell_shared_si = shared_formula_si(&reader, &element)?;
                if let Some(builder) = cell_builder.as_mut() {
                    builder.array_ref = array_formula_ref(&reader, &element)?;
                }
            }
            Event::Empty(element) if element.local_name().as_ref() == b"f" => {
                // Self-closing shared-formula follower (<f t="shared" si="N"/>):
                // inherit the master's formula shifted by the cell offset (#165).
                if let (Some(si), Some(builder)) = (
                    shared_formula_si(&reader, &element)?,
                    cell_builder.as_mut(),
                ) {
                    if let Some(formula) = shared_formulas.expand(si, builder.row, builder.column) {
                        builder.formula = formula;
                    }
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"v" => in_value = true,
            // Inline strings can carry <rPh> ruby annotations too.
            Event::Start(element) if element.local_name().as_ref() == b"rPh" => {
                in_phonetic = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"rPh" => {
                in_phonetic = false;
            }
            Event::Start(element) if element.local_name().as_ref() == b"t" => {
                in_text = !in_phonetic;
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"r" && cell_builder.is_some() =>
            {
                if let Some(builder) = &mut cell_builder {
                    builder.current_run = Some(RichRun::default());
                }
            }
            Event::End(element)
                if element.local_name().as_ref() == b"r" && cell_builder.is_some() =>
            {
                if let Some(builder) = &mut cell_builder {
                    if let Some(run) = builder.current_run.take() {
                        builder.inline_runs.push(run);
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if cell_builder
                    .as_ref()
                    .is_some_and(|builder| builder.current_run.is_some()) =>
            {
                if let Some(run) = cell_builder
                    .as_mut()
                    .and_then(|builder| builder.current_run.as_mut())
                {
                    apply_run_property(run, &reader, &element, colors)?;
                }
            }
            Event::Text(text) => {
                if let Some(builder) = &mut cell_builder {
                    let decoded = decode_text(&text)?;
                    if in_formula {
                        builder.formula.push_str(&decoded);
                    } else if in_value {
                        builder.raw_value.push_str(&decoded);
                    } else if in_text {
                        builder.inline_text.push_str(&decoded);
                        if let Some(run) = &mut builder.current_run {
                            run.text.push_str(&decoded);
                        }
                    }
                } else if in_cf_formula {
                    if let Some(formula) = cf_rule
                        .as_mut()
                        .and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decode_text(&text)?);
                    }
                } else if in_cf_ext_id {
                    cf_rule_ext_id
                        .get_or_insert_with(String::new)
                        .push_str(&decode_text(&text)?);
                } else if in_dv_formula {
                    if let Some(formula) = dv_rule
                        .as_mut()
                        .and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decode_text(&text)?);
                    }
                }
            }
            Event::GeneralRef(reference) => {
                let decoded = general_ref_text(&reference)?;
                if let Some(builder) = &mut cell_builder {
                    if in_formula {
                        builder.formula.push_str(&decoded);
                    } else if in_value {
                        builder.raw_value.push_str(&decoded);
                    } else if in_text {
                        builder.inline_text.push_str(&decoded);
                        if let Some(run) = &mut builder.current_run {
                            run.text.push_str(&decoded);
                        }
                    }
                } else if in_cf_formula {
                    if let Some(formula) = cf_rule
                        .as_mut()
                        .and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decoded);
                    }
                } else if in_dv_formula {
                    if let Some(formula) = dv_rule
                        .as_mut()
                        .and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decoded);
                    }
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"f" => {
                in_formula = false;
                if let (Some(si), Some(builder)) = (cell_shared_si.take(), cell_builder.as_mut()) {
                    if builder.formula.is_empty() {
                        // open-empty follower (<f t="shared" si="N"></f>)
                        if let Some(formula) =
                            shared_formulas.expand(si, builder.row, builder.column)
                        {
                            builder.formula = formula;
                        }
                    } else {
                        shared_formulas.register(si, builder.row, builder.column, &builder.formula);
                    }
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"v" => in_value = false,
            Event::End(element) if element.local_name().as_ref() == b"t" => in_text = false,
            Event::End(element) if element.local_name().as_ref() == b"c" => {
                if let Some(builder) = cell_builder.take() {
                    latest_row = latest_row.max(builder.row);
                    let cell_chunk = builder.row / CHUNK_ROW_COUNT;
                    if cell_chunk != chunk_index {
                        flush_chunk(
                            sheet_index,
                            chunk_index,
                            &mut chunk,
                            &mut pending_formulas,
                            cache_directory,
                            state,
                            cell_chunk * CHUNK_ROW_COUNT - 1,
                        )?;
                        chunk_index = cell_chunk;
                    }
                    if let Some(cell) = builder.finish(shared_strings, styled_xfs)? {
                        if cell.formula.is_some() {
                            pending_formulas.push(cell.clone());
                        }
                        chunk.cells.push(cell);
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"mergeCell" =>
            {
                if let Some(reference) = attribute_value(&reader, &element, b"ref")? {
                    if let Some(merge) = parse_merge_reference(&reference) {
                        merges.push(merge);
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"hyperlink" =>
            {
                if let Some(reference) = attribute_value(&reader, &element, b"ref")? {
                    let anchor = reference.split(':').next().unwrap_or(&reference);
                    if let Ok((row, column)) = parse_address(&anchor.replace('$', "")) {
                        let target = match attribute_value(&reader, &element, b"id")? {
                            Some(id) => link_targets.get(&id).cloned(),
                            None => attribute_value(&reader, &element, b"location")?
                                .map(|location| format!("#{location}")),
                        };
                        if let Some(target) = target {
                            hyperlinks.push(HyperlinkRecord {
                                row,
                                column,
                                target,
                            });
                        }
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"autoFilter" =>
            {
                if auto_filter.is_none() {
                    auto_filter = attribute_value(&reader, &element, b"ref")?
                        .and_then(|reference| parse_area_reference(&reference));
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"sheetProtection" =>
            {
                let protected = attribute_value(&reader, &element, b"sheet")?
                    .is_some_and(|value| value == "1" || value == "true");
                let has_password = attribute_value(&reader, &element, b"password")?.is_some()
                    || attribute_value(&reader, &element, b"hashValue")?.is_some();
                sheet_protection = Some(SheetProtectionInfo {
                    protected,
                    has_password,
                });
            }
            Event::Start(element) if element.local_name().as_ref() == b"rowBreaks" => {
                in_row_breaks = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"rowBreaks" => {
                in_row_breaks = false;
            }
            Event::Start(element) if element.local_name().as_ref() == b"colBreaks" => {
                in_col_breaks = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"colBreaks" => {
                in_col_breaks = false;
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"brk"
                    && (in_row_breaks || in_col_breaks) =>
            {
                // Only manual breaks: Excel also caches automatic ones when
                // printer metrics are embedded, which we recompute instead.
                let manual = attribute_value(&reader, &element, b"man")?
                    .is_some_and(|value| value == "1" || value == "true");
                if manual {
                    if let Some(id) = attribute_value(&reader, &element, b"id")?
                        .and_then(|value| value.parse::<usize>().ok())
                    {
                        if in_row_breaks {
                            row_breaks.push(id);
                        } else {
                            col_breaks.push(id);
                        }
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"protectedRange" =>
            {
                if let (Some(name), Some(sqref)) = (
                    attribute_value(&reader, &element, b"name")?,
                    attribute_value(&reader, &element, b"sqref")?,
                ) {
                    // securityDescriptor carries per-user permissions the
                    // rewrite cannot preserve — treat like a password.
                    let has_password = attribute_value(&reader, &element, b"password")?
                        .is_some()
                        || attribute_value(&reader, &element, b"hashValue")?.is_some()
                        || attribute_value(&reader, &element, b"securityDescriptor")?.is_some();
                    protected_ranges.push(ProtectedRangeInfo {
                        name,
                        sqref,
                        has_password,
                    });
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"securityDescriptor" =>
            {
                // The child-element form of the ACL, nested in the
                // protectedRange just pushed.
                if let Some(range) = protected_ranges.last_mut() {
                    range.has_password = true;
                }
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidation" =>
            {
                dv_rule = parse_dv_rule(&reader, &element)?;
            }
            Event::Empty(element)
                if element.local_name().as_ref() == b"dataValidation" =>
            {
                if let Some(rule) = parse_dv_rule(&reader, &element)? {
                    data_validations.push(rule);
                }
            }
            Event::Start(element)
                if element.local_name().as_ref().starts_with(b"formula")
                    && dv_rule.is_some() =>
            {
                in_dv_formula = true;
                if let Some(rule) = &mut dv_rule {
                    rule.formulas.push(String::new());
                }
            }
            Event::End(element)
                if element.local_name().as_ref().starts_with(b"formula")
                    && dv_rule.is_some() =>
            {
                in_dv_formula = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"dataValidation" => {
                if let Some(rule) = dv_rule.take() {
                    data_validations.push(rule);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"conditionalFormatting" =>
            {
                cf_ranges = attribute_value(&reader, &element, b"sqref")?
                    .map(|sqref| {
                        sqref
                            .split_whitespace()
                            .filter_map(parse_area_reference)
                            .collect()
                    })
                    .unwrap_or_default();
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"cfRule" && !cf_ranges.is_empty() =>
            {
                cf_rule = Some(parse_cf_rule(&reader, &element, &cf_ranges)?);
            }
            Event::Empty(element)
                if element.local_name().as_ref() == b"cfRule" && !cf_ranges.is_empty() =>
            {
                conditional_rules.push(parse_cf_rule(&reader, &element, &cf_ranges)?);
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"iconSet" && cf_rule.is_some() =>
            {
                if let Some(rule) = &mut cf_rule {
                    rule.icon_set_name = Some(
                        attribute_value(&reader, &element, b"iconSet")?
                            .unwrap_or_else(|| "3TrafficLights1".into()),
                    );
                    rule.icon_reverse = attribute_value(&reader, &element, b"reverse")?
                        .is_some_and(|value| value == "1" || value == "true");
                    rule.show_value = attribute_value(&reader, &element, b"showValue")?
                        .map(|value| value != "0" && value != "false")
                        .unwrap_or(true);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"dataBar" && cf_rule.is_some() =>
            {
                if let Some(rule) = &mut cf_rule {
                    rule.show_value = attribute_value(&reader, &element, b"showValue")?
                        .map(|value| value != "0" && value != "false")
                        .unwrap_or(true);
                }
            }
            // x14:dataBar carries the attributes the 2006 element lacks.
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"dataBar" && cf_rule.is_none() =>
            {
                if let Some(id) = &x14_current_id {
                    let gradient = attribute_value(&reader, &element, b"gradient")?
                        .map(|value| value != "0" && value != "false");
                    let same_as_positive =
                        attribute_value(&reader, &element, b"negativeBarColorSameAsPositive")?
                            .map(|value| value != "0" && value != "false")
                            .unwrap_or(true);
                    x14_databars.insert(id.clone(), (gradient, same_as_positive));
                }
            }
            // <x14:id> inside a main cfRule's extLst links the rule to its
            // x14 twin in the worksheet extLst.
            Event::Start(element)
                if element.local_name().as_ref() == b"id" && cf_rule.is_some() =>
            {
                in_cf_ext_id = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"id" => {
                in_cf_ext_id = false;
            }
            // x14:cfRule (its parent x14:conditionalFormatting has no sqref
            // attribute, so cf_ranges is empty and the main arms skip it).
            Event::Start(element)
                if element.local_name().as_ref() == b"cfRule" && cf_ranges.is_empty() =>
            {
                x14_current_id = attribute_value(&reader, &element, b"id")?;
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"negativeFillColor" =>
            {
                if let Some(id) = &x14_current_id {
                    if let Some(color) = visuals::resolve_color(
                        attribute_value(&reader, &element, b"rgb")?.as_deref(),
                        attribute_value(&reader, &element, b"indexed")?.as_deref(),
                        attribute_value(&reader, &element, b"theme")?.as_deref(),
                        attribute_value(&reader, &element, b"tint")?.as_deref(),
                        colors,
                    ) {
                        x14_negative_colors.insert(id.clone(), color);
                    }
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"cfvo" && cf_rule.is_some() =>
            {
                if let Some(rule) = &mut cf_rule {
                    rule.cfvos.push(Cfvo {
                        kind: attribute_value(&reader, &element, b"type")?
                            .unwrap_or_else(|| "num".into()),
                        value: attribute_value(&reader, &element, b"val")?,
                        gte: attribute_value(&reader, &element, b"gte")?
                            .and_then(|value| (value == "0" || value == "false").then_some(false)),
                    });
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"color" && cf_rule.is_some() =>
            {
                if let Some(rule) = &mut cf_rule {
                    if let Some(color) = visuals::resolve_color(
                        attribute_value(&reader, &element, b"rgb")?.as_deref(),
                        attribute_value(&reader, &element, b"indexed")?.as_deref(),
                        attribute_value(&reader, &element, b"theme")?.as_deref(),
                        attribute_value(&reader, &element, b"tint")?.as_deref(),
                        colors,
                    ) {
                        rule.colors.push(color);
                    }
                }
            }
            Event::Start(element)
                if element.local_name().as_ref() == b"formula" && cf_rule.is_some() =>
            {
                in_cf_formula = true;
                if let Some(rule) = &mut cf_rule {
                    rule.formulas.push(String::new());
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"formula" => {
                in_cf_formula = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"cfRule" => {
                if let Some(rule) = cf_rule.take() {
                    conditional_rules.push(rule);
                    if let Some(id) = cf_rule_ext_id.take() {
                        x14_rule_slots.insert(id, conditional_rules.len() - 1);
                    }
                } else {
                    x14_current_id = None;
                }
            }
            Event::Eof => {
                flush_chunk(
                    sheet_index,
                    chunk_index,
                    &mut chunk,
                    &mut pending_formulas,
                    cache_directory,
                    state,
                    latest_row,
                )?;
                break;
            }
            _ => {}
        }
        buffer.clear();
    }
    let (lock, condition) = &**state;
    let mut index = lock
        .lock()
        .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
    index.merges = merges;
    index.hyperlinks = hyperlinks;
    for (id, color) in x14_negative_colors {
        if let Some(rule) = x14_rule_slots
            .get(&id)
            .and_then(|slot| conditional_rules.get_mut(*slot))
        {
            rule.negative_color = Some(color);
        }
    }
    for (id, (gradient, same_as_positive)) in x14_databars {
        if let Some(rule) = x14_rule_slots
            .get(&id)
            .and_then(|slot| conditional_rules.get_mut(*slot))
        {
            rule.gradient = gradient;
            rule.negative_same_as_positive = Some(same_as_positive);
        }
    }
    index.conditional_rules = conditional_rules;
    index.auto_filter = auto_filter;
    index.sheet_protection = sheet_protection;
    // Wire caps (schema and preload reject larger sets; the save side caps
    // at the same sizes) — a pathological file loses tail entries, not the
    // whole sheet.
    row_breaks.truncate(1_023);
    col_breaks.truncate(1_023);
    protected_ranges.truncate(1_000);
    index.row_breaks = row_breaks;
    index.col_breaks = col_breaks;
    index.protected_ranges = protected_ranges;
    index.data_validations = data_validations;
    condition.notify_all();
    drop(index);
    Ok(())
}

fn parse_dv_rule<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<Option<DataValidationRule>, SidecarError> {
    let Some(sqref) = attribute_value(reader, element, b"sqref")? else {
        return Ok(None);
    };
    let ranges = sqref
        .split_whitespace()
        .filter_map(parse_area_reference)
        .collect::<Vec<_>>();
    if ranges.is_empty() {
        return Ok(None);
    }
    let flag = |name: &[u8]| -> Result<bool, SidecarError> {
        Ok(attribute_value(reader, element, name)?
            .is_some_and(|value| value == "1" || value == "true"))
    };
    Ok(Some(DataValidationRule {
        ranges,
        rule_type: attribute_value(reader, element, b"type")?
            .unwrap_or_else(|| "none".into()),
        operator: attribute_value(reader, element, b"operator")?,
        formulas: Vec::new(),
        allow_blank: flag(b"allowBlank")?,
        suppress_dropdown: flag(b"showDropDown")?,
        show_input_message: flag(b"showInputMessage")?,
        show_error_message: flag(b"showErrorMessage")?,
        error_style: attribute_value(reader, element, b"errorStyle")?,
        error_title: attribute_value(reader, element, b"errorTitle")?,
        error: attribute_value(reader, element, b"error")?,
        prompt_title: attribute_value(reader, element, b"promptTitle")?,
        prompt: attribute_value(reader, element, b"prompt")?,
    }))
}

fn parse_cf_rule<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    ranges: &[MergedRange],
) -> Result<ConditionalRule, SidecarError> {
    Ok(ConditionalRule {
        ranges: ranges.to_vec(),
        rule_type: attribute_value(reader, element, b"type")?
            .unwrap_or_else(|| "unknown".into()),
        operator: attribute_value(reader, element, b"operator")?,
        formulas: Vec::new(),
        text: attribute_value(reader, element, b"text")?,
        dxf_index: attribute_value(reader, element, b"dxfId")?
            .and_then(|value| value.parse::<usize>().ok()),
        priority: attribute_value(reader, element, b"priority")?
            .and_then(|value| value.parse::<i64>().ok())
            .unwrap_or(0),
        rank: attribute_value(reader, element, b"rank")?
            .and_then(|value| value.parse::<u32>().ok()),
        percent: attribute_value(reader, element, b"percent")?
            .is_some_and(|value| value == "1" || value == "true"),
        bottom: attribute_value(reader, element, b"bottom")?
            .is_some_and(|value| value == "1" || value == "true"),
        cfvos: Vec::new(),
        colors: Vec::new(),
        icon_set_name: None,
        icon_reverse: false,
        show_value: true,
        negative_color: None,
        negative_same_as_positive: None,
        gradient: None,
    })
}

fn read_defined_names(archive: &mut ZipArchive<File>) -> Result<Vec<DefinedName>, SidecarError> {
    let xml = read_zip_string(archive, "xl/workbook.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut names = Vec::new();
    let mut current: Option<DefinedName> = None;
    loop {
        match reader.read_event()? {
            Event::Start(element) if element.local_name().as_ref() == b"definedName" => {
                let name = attribute_value(&reader, &element, b"name")?.unwrap_or_default();
                let hidden = attribute_value(&reader, &element, b"hidden")?
                    .is_some_and(|value| value == "1" || value == "true");
                let sheet_index = attribute_value(&reader, &element, b"localSheetId")?
                    .and_then(|value| value.parse::<usize>().ok());
                // _xlnm.* built-ins and hidden names stay file-only (the save
                // preserves them verbatim; the editor never models them).
                current = (!name.is_empty() && !name.starts_with("_xlnm") && !hidden)
                    .then_some(DefinedName {
                        name,
                        formula: String::new(),
                        sheet_index,
                    });
            }
            Event::Text(text) => {
                if let Some(defined) = &mut current {
                    defined.formula.push_str(&decode_text(&text)?);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some(defined) = &mut current {
                    defined.formula.push_str(&general_ref_text(&reference)?);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"definedName" => {
                if let Some(defined) = current.take() {
                    if !defined.formula.is_empty() {
                        names.push(defined);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(names)
}

const DEFAULT_ACCENTS: [(u8, u8, u8); 6] = [
    (0x44, 0x72, 0xC4),
    (0xED, 0x7D, 0x31),
    (0xA5, 0xA5, 0xA5),
    (0xFF, 0xC0, 0x00),
    (0x5B, 0x9B, 0xD5),
    (0x70, 0xAD, 0x47),
];

fn rgb_hex((red, green, blue): (u8, u8, u8)) -> String {
    format!("#{red:02X}{green:02X}{blue:02X}")
}

/// Approximates the built-in table style families using the workbook's real
/// theme accents: accent cycle is (n-1) % 7 with 0 = neutral gray, variants
/// come in blocks of 7 (Light 1-21, Medium 1-28, Dark 1-11).
/// Fills/font colors a table band can draw, resolved from a custom
/// `<tableStyle>` in styles.xml (each element references a dxf).
#[derive(Clone, Debug, Default)]
struct CustomTablePalette {
    header_fill: Option<String>,
    header_font_color: Option<String>,
    stripe_fill: Option<String>,
    second_row_stripe_fill: Option<String>,
    column_stripe_fill: Option<String>,
    second_column_stripe_fill: Option<String>,
    whole_table_fill: Option<String>,
    first_column_fill: Option<String>,
    last_column_fill: Option<String>,
    total_row_fill: Option<String>,
    total_row_font_color: Option<String>,
    first_header_cell_font_color: Option<String>,
}

/// styles.xml `<tableStyles>` → per-style band palette. Only fills and font
/// colors are carried; band borders stay unsupported.
fn read_custom_table_styles(
    archive: &mut ZipArchive<File>,
    dxf_styles: &[visuals::CellStyle],
) -> HashMap<String, CustomTablePalette> {
    let mut palettes = HashMap::new();
    let Ok(xml) = read_zip_string(archive, "xl/styles.xml") else {
        return palettes;
    };
    let mut reader = Reader::from_str(&xml);
    let mut current: Option<(String, CustomTablePalette)> = None;
    loop {
        let Ok(event) = reader.read_event() else {
            break;
        };
        match event {
            Event::Start(element) if element.local_name().as_ref() == b"tableStyle" => {
                if let Ok(Some(name)) = attribute_value(&reader, &element, b"name") {
                    current = Some((name, CustomTablePalette::default()));
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"tableStyleElement" =>
            {
                let Some((_, palette)) = current.as_mut() else {
                    continue;
                };
                let kind = attribute_value(&reader, &element, b"type").ok().flatten();
                let dxf = attribute_value(&reader, &element, b"dxfId")
                    .ok()
                    .flatten()
                    .and_then(|value| value.parse::<usize>().ok())
                    .and_then(|index| dxf_styles.get(index));
                let Some(dxf) = dxf else { continue };
                match kind.as_deref() {
                    Some("wholeTable") => {
                        palette.whole_table_fill = dxf.fill_color.clone();
                    }
                    Some("headerRow") => {
                        palette.header_fill = dxf.fill_color.clone();
                        palette.header_font_color = dxf.font_color.clone();
                    }
                    Some("totalRow") => {
                        palette.total_row_fill = dxf.fill_color.clone();
                        palette.total_row_font_color = dxf.font_color.clone();
                    }
                    Some("firstColumn") => {
                        palette.first_column_fill = dxf.fill_color.clone();
                    }
                    Some("lastColumn") => {
                        palette.last_column_fill = dxf.fill_color.clone();
                    }
                    Some("firstRowStripe") => {
                        palette.stripe_fill = dxf.fill_color.clone();
                    }
                    Some("secondRowStripe") => {
                        palette.second_row_stripe_fill = dxf.fill_color.clone();
                    }
                    Some("firstColumnStripe") => {
                        palette.column_stripe_fill = dxf.fill_color.clone();
                    }
                    Some("secondColumnStripe") => {
                        palette.second_column_stripe_fill = dxf.fill_color.clone();
                    }
                    Some("firstHeaderCell") => {
                        palette.first_header_cell_font_color = dxf.font_color.clone();
                    }
                    _ => {}
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"tableStyle" => {
                if let Some((name, palette)) = current.take() {
                    palettes.insert(name, palette);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    palettes
}

/// Light-family pivot styles paint header and grand-total bands with a light
/// tint of the style's accent; Medium/Dark families are left unpainted until
/// their (solid-fill, white-text) bands are modelled.
fn pivot_style_fill(style_name: Option<&str>, colors: &ColorContext) -> Option<String> {
    let rest = style_name?.strip_prefix("PivotStyleLight")?;
    let number = rest.parse::<usize>().ok()?;
    let accent_index = number.saturating_sub(1) % 7;
    let base = if accent_index == 0 {
        visuals::theme_dark1(colors).unwrap_or((0x00, 0x00, 0x00))
    } else {
        visuals::theme_accent(colors, accent_index).unwrap_or(DEFAULT_ACCENTS[accent_index - 1])
    };
    Some(visuals::tint_to_hex(base, 0.8))
}

/// Table style column 1 of the Light block (Light1-7) draws its frame in the
/// base color; the filled families' borders stay unmodelled.
fn table_style_border(style_name: Option<&str>, colors: &ColorContext) -> Option<String> {
    let rest = style_name?.strip_prefix("TableStyleLight")?;
    let number = rest.parse::<usize>().ok()?;
    if number.saturating_sub(1) / 7 != 0 {
        return None;
    }
    let accent_index = number.saturating_sub(1) % 7;
    let base = if accent_index == 0 {
        visuals::theme_dark1(colors).unwrap_or((0x00, 0x00, 0x00))
    } else {
        visuals::theme_accent(colors, accent_index).unwrap_or(DEFAULT_ACCENTS[accent_index - 1])
    };
    Some(rgb_hex(base))
}

fn table_style_palette(
    style_name: Option<&str>,
    colors: &ColorContext,
) -> (Option<String>, Option<String>, Option<String>) {
    // A tableStyleInfo without a name (or none at all) is Excel's style
    // "None": nothing gets painted.
    let Some(name) = style_name else {
        return (None, None, None);
    };
    let (family, number) = if let Some(rest) = name.strip_prefix("TableStyleLight") {
        ("light", rest.parse::<usize>().unwrap_or(1))
    } else if let Some(rest) = name.strip_prefix("TableStyleMedium") {
        ("medium", rest.parse::<usize>().unwrap_or(2))
    } else if let Some(rest) = name.strip_prefix("TableStyleDark") {
        ("dark", rest.parse::<usize>().unwrap_or(1))
    } else {
        ("medium", 2)
    };
    let accent_index = number.saturating_sub(1) % 7;
    let base = if accent_index == 0 {
        // Column 1 of each family is the theme's neutral text color (dk1),
        // not a literal gray.
        visuals::theme_dark1(colors).unwrap_or((0x00, 0x00, 0x00))
    } else {
        visuals::theme_accent(colors, accent_index)
            .unwrap_or(DEFAULT_ACCENTS[accent_index - 1])
    };
    let variant = number.saturating_sub(1) / 7;
    match family {
        "light" => {
            let stripe = Some(visuals::tint_to_hex(base, 0.85));
            if variant == 1 {
                (Some(rgb_hex(base)), Some("#FFFFFF".into()), stripe)
            } else {
                // Light 1-7 / 15-21: unfilled bold header in the accent color.
                (None, Some(rgb_hex(base)), stripe)
            }
        }
        "dark" => (
            Some(visuals::tint_to_hex(base, -0.25)),
            Some("#FFFFFF".into()),
            Some(visuals::tint_to_hex(base, 0.4)),
        ),
        _ => {
            let header = if variant == 2 {
                visuals::tint_to_hex(base, -0.25)
            } else {
                rgb_hex(base)
            };
            // Medium 8-14 are Excel's "full color" block: darker stripes
            // (accent tint 0.6, e.g. Medium9 #B8CCE4) than the white-bodied
            // blocks (tint 0.8, e.g. Medium2 #DCE6F1).
            let stripe_tint = if variant == 1 { 0.6 } else { 0.8 };
            (
                Some(header),
                Some("#FFFFFF".into()),
                Some(visuals::tint_to_hex(base, stripe_tint)),
            )
        }
    }
}

fn read_sheet_tables(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    colors: &ColorContext,
    custom_styles: &HashMap<String, CustomTablePalette>,
) -> Result<Vec<TableInfo>, SidecarError> {
    let mut tables = Vec::new();
    for table_path in visuals::table_part_paths(archive, worksheet_path)? {
        let xml = read_zip_string(archive, &table_path)?;
        let mut reader = Reader::from_str(&xml);
        let mut range = None;
        let mut header_row_count = 1;
        let mut totals_row_count = 0;
        let mut show_first_column = false;
        let mut show_last_column = false;
        let mut show_row_stripes = false;
        let mut show_column_stripes = false;
        let mut style_name = None;
        let mut name = None;
        let mut columns = Vec::new();
        loop {
            match reader.read_event()? {
                // Match by local name so `x:`-prefixed writers are covered,
                // but require a ref attribute: the extLst twin <x14:table
                // altText=…/> shares the local name and would otherwise reset
                // the parsed range, dropping the whole table.
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"table" =>
                {
                    let Some(reference) = attribute_value(&reader, &element, b"ref")? else {
                        continue;
                    };
                    range = parse_area_reference(&reference);
                    header_row_count = attribute_value(&reader, &element, b"headerRowCount")?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1);
                    totals_row_count = attribute_value(&reader, &element, b"totalsRowCount")?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(0);
                    // displayName is the token structured references use; name is a fallback
                    name = attribute_value(&reader, &element, b"displayName")?
                        .or(attribute_value(&reader, &element, b"name")?);
                }
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"tableColumn" =>
                {
                    if let Some(column) = attribute_value(&reader, &element, b"name")? {
                        columns.push(column);
                    }
                }
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"tableStyleInfo" =>
                {
                    style_name = attribute_value(&reader, &element, b"name")?;
                    show_first_column = attribute_value(&reader, &element, b"showFirstColumn")?
                        .is_some_and(|value| value == "1" || value == "true");
                    show_last_column = attribute_value(&reader, &element, b"showLastColumn")?
                        .is_some_and(|value| value == "1" || value == "true");
                    show_row_stripes = attribute_value(&reader, &element, b"showRowStripes")?
                        .is_some_and(|value| value == "1" || value == "true");
                    show_column_stripes =
                        attribute_value(&reader, &element, b"showColumnStripes")?
                            .is_some_and(|value| value == "1" || value == "true");
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if let Some(range) = range {
            let custom = style_name
                .as_deref()
                .and_then(|style| custom_styles.get(style));
            let (header_fill, header_font_color, stripe_fill) = match custom {
                Some(palette) => (
                    palette.header_fill.clone(),
                    palette.header_font_color.clone(),
                    palette.stripe_fill.clone(),
                ),
                None => table_style_palette(style_name.as_deref(), colors),
            };
            let border_color = if custom.is_some() {
                None
            } else {
                table_style_border(style_name.as_deref(), colors)
            };
            tables.push(TableInfo {
                range,
                header_row_count,
                show_row_stripes,
                show_column_stripes,
                name,
                columns,
                style_name,
                header_fill,
                header_font_color,
                stripe_fill,
                second_row_stripe_fill: custom.and_then(|p| p.second_row_stripe_fill.clone()),
                column_stripe_fill: custom.and_then(|p| p.column_stripe_fill.clone()),
                second_column_stripe_fill: custom
                    .and_then(|p| p.second_column_stripe_fill.clone()),
                whole_table_fill: custom.and_then(|p| p.whole_table_fill.clone()),
                // First/last column emphasis (and the first-header corner
                // cell) only paint when tableStyleInfo turns them on.
                first_column_fill: custom
                    .filter(|_| show_first_column)
                    .and_then(|p| p.first_column_fill.clone()),
                last_column_fill: custom
                    .filter(|_| show_last_column)
                    .and_then(|p| p.last_column_fill.clone()),
                total_row_fill: custom.and_then(|p| p.total_row_fill.clone()),
                total_row_font_color: custom.and_then(|p| p.total_row_font_color.clone()),
                first_header_cell_font_color: custom
                    .filter(|_| show_first_column)
                    .and_then(|p| p.first_header_cell_font_color.clone()),
                totals_row_count,
                border_color,
            });
        }
    }
    Ok(tables)
}

const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
const MAX_SPARKLINE_GROUPS: usize = 100;
const MAX_SPARKLINE_CELLS: usize = 500;

fn sparkline_type(value: Option<&str>) -> &'static str {
    match value {
        Some("column") => "column",
        Some("stacked") => "stacked",
        _ => "line",
    }
}

/// ARGB or RGB hex attribute → `#RRGGBB` (alpha dropped, no theme resolution).
fn argb_to_hex(value: &str) -> Option<String> {
    let hex = value.trim();
    let hex = match hex.len() {
        8 => &hex[2..],
        6 => hex,
        _ => return None,
    };
    hex.bytes()
        .all(|byte| byte.is_ascii_hexdigit())
        .then(|| format!("#{}", hex.to_ascii_uppercase()))
}

fn read_sheet_sparklines(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<SparklineGroupInfo>, SidecarError> {
    let entry = archive.by_name(worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    parse_sparkline_groups(&mut reader)
}

fn parse_sparkline_groups<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<Vec<SparklineGroupInfo>, SidecarError> {
    enum Target {
        None,
        Formula,
        Sqref,
    }
    let mut groups = Vec::new();
    let mut buffer = Vec::new();
    let mut in_ext = false;
    let mut group: Option<SparklineGroupInfo> = None;
    let mut in_sparkline = false;
    let mut target = Target::None;
    let mut source_ref = String::new();
    let mut host_cell = String::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if element.local_name().as_ref() == b"ext" => {
                in_ext = attribute_value(reader, &element, b"uri")?
                    .is_some_and(|uri| uri.eq_ignore_ascii_case(SPARKLINE_EXT_URI));
            }
            Event::End(element) if element.local_name().as_ref() == b"ext" => {
                in_ext = false;
                group = None;
                in_sparkline = false;
                target = Target::None;
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"sparklineGroup" =>
            {
                group = Some(SparklineGroupInfo {
                    kind: sparkline_type(
                        attribute_value(reader, &element, b"type")?.as_deref(),
                    )
                    .into(),
                    color: None,
                    negative_color: None,
                    cells: Vec::new(),
                });
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"colorSeries" =>
            {
                if let Some(group) = &mut group {
                    group.color = attribute_value(reader, &element, b"rgb")?
                        .as_deref()
                        .and_then(argb_to_hex);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"colorNegative" =>
            {
                if let Some(group) = &mut group {
                    group.negative_color = attribute_value(reader, &element, b"rgb")?
                        .as_deref()
                        .and_then(argb_to_hex);
                }
            }
            Event::Start(element)
                if in_ext && element.local_name().as_ref() == b"sparkline" =>
            {
                in_sparkline = true;
                source_ref.clear();
                host_cell.clear();
            }
            Event::Start(element) if in_sparkline && element.local_name().as_ref() == b"f" => {
                target = Target::Formula;
            }
            Event::Start(element)
                if in_sparkline && element.local_name().as_ref() == b"sqref" =>
            {
                target = Target::Sqref;
            }
            Event::Text(text) if in_sparkline => {
                let value = decode_text(&text)?;
                match target {
                    Target::Formula => source_ref.push_str(&value),
                    Target::Sqref => host_cell.push_str(&value),
                    Target::None => {}
                }
            }
            Event::GeneralRef(reference) if in_sparkline => {
                let value = general_ref_text(&reference)?;
                match target {
                    Target::Formula => source_ref.push_str(&value),
                    Target::Sqref => host_cell.push_str(&value),
                    Target::None => {}
                }
            }
            Event::End(element)
                if in_sparkline
                    && (element.local_name().as_ref() == b"f"
                        || element.local_name().as_ref() == b"sqref") =>
            {
                target = Target::None;
            }
            Event::End(element)
                if in_sparkline && element.local_name().as_ref() == b"sparkline" =>
            {
                in_sparkline = false;
                target = Target::None;
                let cell = host_cell.split_whitespace().next().unwrap_or("");
                let formula = source_ref.trim();
                if let Some(group) = &mut group {
                    if !cell.is_empty()
                        && !formula.is_empty()
                        && group.cells.len() < MAX_SPARKLINE_CELLS
                    {
                        group.cells.push(SparklineCellInfo {
                            cell: cell.to_owned(),
                            source_ref: formula.to_owned(),
                        });
                    }
                }
            }
            Event::End(element)
                if in_ext && element.local_name().as_ref() == b"sparklineGroup" =>
            {
                if let Some(group) = group.take() {
                    if !group.cells.is_empty() && groups.len() < MAX_SPARKLINE_GROUPS {
                        groups.push(group);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(groups)
}

fn parse_area_reference(reference: &str) -> Option<MergedRange> {
    let cleaned = reference.replace('$', "");
    match cleaned.split_once(':') {
        Some((start, end)) => {
            let (start_row, start_column) = parse_address(start).ok()?;
            let (end_row, end_column) = parse_address(end).ok()?;
            (end_row >= start_row && end_column >= start_column).then_some(MergedRange {
                start_row,
                start_column,
                end_row,
                end_column,
            })
        }
        None => {
            let (row, column) = parse_address(&cleaned).ok()?;
            Some(MergedRange {
                start_row: row,
                start_column: column,
                end_row: row,
                end_column: column,
            })
        }
    }
}

fn row_property<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    row: usize,
) -> Result<Option<RowProperty>, SidecarError> {
    // ht is reported with or without customHeight="1", so the renderer can
    // use Excel's laid-out height as a floor (re-measuring with substitute
    // fonts clipped wrapped CJK rows on Windows). customHeight distinguishes
    // a user-fixed height (clip like Excel) from an auto-fit result the
    // renderer may still grow past.
    let height = attribute_value(reader, element, b"ht")?
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| *value >= 0.0);
    let custom_height = attribute_value(reader, element, b"customHeight")?
        .is_some_and(|value| value == "1" || value == "true");
    let hidden = attribute_value(reader, element, b"hidden")?
        .is_some_and(|value| value == "1" || value == "true");
    let outline_level = attribute_value(reader, element, b"outlineLevel")?
        .and_then(|value| value.parse::<u8>().ok())
        .map(|value| value.min(7))
        .filter(|value| *value > 0);
    let collapsed = attribute_value(reader, element, b"collapsed")?
        .is_some_and(|value| value == "1" || value == "true");
    // <row s=> only styles the row when customFormat is set (spec: without it
    // the attribute is leftover noise Excel ignores). xf 0 is the default.
    let style_index = attribute_value(reader, element, b"s")?
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|value| *value > 0)
        .filter(|_| {
            matches!(
                attribute_value(reader, element, b"customFormat").ok().flatten().as_deref(),
                Some("1") | Some("true")
            )
        });
    if height.is_none() && !hidden && outline_level.is_none() && !collapsed && style_index.is_none()
    {
        return Ok(None);
    }
    Ok(Some(RowProperty {
        row,
        height,
        custom_height: custom_height && height.is_some(),
        hidden,
        outline_level,
        collapsed,
        style_index,
    }))
}

fn parse_merge_reference(reference: &str) -> Option<MergedRange> {
    let (start, end) = reference.split_once(':')?;
    let (start_row, start_column) = parse_address(&start.replace('$', "")).ok()?;
    let (end_row, end_column) = parse_address(&end.replace('$', "")).ok()?;
    if end_row < start_row || end_column < start_column {
        return None;
    }
    Some(MergedRange {
        start_row,
        start_column,
        end_row,
        end_column,
    })
}

fn flush_chunk(
    sheet_index: usize,
    chunk_index: usize,
    chunk: &mut ChunkData,
    pending_formulas: &mut Vec<CellRecord>,
    cache_directory: &Path,
    state: &Arc<(Mutex<SheetIndex>, Condvar)>,
    indexed_through_row: usize,
) -> Result<(), SidecarError> {
    let path = cache_directory.join(format!("sheet-{sheet_index}-chunk-{chunk_index}.json"));
    let has_data = !chunk.cells.is_empty() || !chunk.rows.is_empty();
    if has_data {
        let file = File::create(&path)?;
        let mut writer = BufWriter::new(file);
        serde_json::to_writer(&mut writer, chunk)?;
        writer.flush()?;
        chunk.cells.clear();
        chunk.rows.clear();
    }
    let (lock, condition) = &**state;
    let mut index = lock
        .lock()
        .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
    if has_data {
        index.chunk_files.insert(chunk_index, path);
    }
    if !pending_formulas.is_empty() {
        let room = MAX_FORMULA_CELLS.saturating_sub(index.formula_cells.len());
        if pending_formulas.len() > room {
            index.formula_truncated = true;
        }
        index.formula_cells.extend(pending_formulas.drain(..).take(room));
    }
    index.indexed_through_row = Some(indexed_through_row);
    condition.notify_all();
    Ok(())
}

struct CellBuilder {
    row: usize,
    column: usize,
    style_index: Option<usize>,
    cell_type: Option<String>,
    raw_value: String,
    formula: String,
    array_ref: Option<String>,
    inline_text: String,
    inline_runs: Vec<RichRun>,
    current_run: Option<RichRun>,
}

impl CellBuilder {
    fn from_element<R: std::io::BufRead>(
        reader: &Reader<R>,
        element: &BytesStart<'_>,
        fallback_row: usize,
        fallback_column: usize,
    ) -> Result<Self, SidecarError> {
        // <c r=> is optional: an unaddressed cell sits one right of its predecessor.
        let (row, column) = match attribute_value(reader, element, b"r")? {
            Some(address) => parse_address(&address)?,
            None => (fallback_row, fallback_column),
        };
        Ok(Self {
            row,
            column,
            style_index: attribute_value(reader, element, b"s")?
                .and_then(|value| value.parse::<usize>().ok()),
            cell_type: attribute_value(reader, element, b"t")?,
            raw_value: String::new(),
            formula: String::new(),
            array_ref: None,
            inline_text: String::new(),
            inline_runs: Vec::new(),
            current_run: None,
        })
    }

    fn finish(
        self,
        shared_strings: &[SharedString],
        styled_xfs: &[bool],
    ) -> Result<Option<CellRecord>, SidecarError> {
        let formula = if self.formula.is_empty() {
            None
        } else {
            Some(format!("={}", self.formula))
        };
        let mut rich = None;
        let value = match self.cell_type.as_deref() {
            // Empty <v/> or a stale index degrades to a valueless styled cell;
            // erroring here used to blank the whole sheet.
            Some("s") => match self
                .raw_value
                .parse::<usize>()
                .ok()
                .and_then(|index| shared_strings.get(index))
            {
                Some(shared) => {
                    rich = shared.runs.clone();
                    Some(CellValue::String(shared.text.clone()))
                }
                None => None,
            },
            Some("inlineStr") => {
                rich = qualify_runs(self.inline_runs);
                Some(CellValue::String(self.inline_text))
            }
            Some("str") | Some("e") => Some(CellValue::String(self.raw_value)),
            Some("b") => Some(CellValue::Boolean(self.raw_value == "1")),
            _ if self.raw_value.is_empty() => None,
            // Real-world exporters write stray non-numeric values without a
            // type attribute; degrade to text instead of failing the sheet.
            _ => match self.raw_value.parse::<f64>() {
                Ok(number) => Some(CellValue::Number(number)),
                Err(_) => Some(CellValue::String(self.raw_value)),
            },
        };
        // Unknown indices keep the cell so a bad styles part can't drop data.
        let styled = self
            .style_index
            .is_some_and(|index| styled_xfs.get(index).copied().unwrap_or(true));
        if value.is_none() && formula.is_none() && !styled {
            return Ok(None);
        }
        Ok(Some(CellRecord {
            row: self.row,
            column: self.column,
            value,
            array_ref: if formula.is_some() { self.array_ref } else { None },
            formula,
            style_index: self.style_index,
            rich,
        }))
    }
}

/// "A3:C20" (or a single "A3") → MergedRange; None on anything unparsable.
fn parse_range_reference(reference: &str) -> Option<MergedRange> {
    let cleaned = reference.replace('$', "");
    let mut parts = cleaned.split(':');
    let (start_row, start_column) = parse_address(parts.next()?).ok()?;
    let (end_row, end_column) = match parts.next() {
        Some(end) => parse_address(end).ok()?,
        None => (start_row, start_column),
    };
    Some(MergedRange {
        start_row: start_row.min(end_row),
        start_column: start_column.min(end_column),
        end_row: start_row.max(end_row),
        end_column: start_column.max(end_column),
    })
}

/// `ref` of a `<f t="array">` element; None for ordinary formulas.
fn array_formula_ref<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<Option<String>, SidecarError> {
    if attribute_value(reader, element, b"t")?.as_deref() != Some("array") {
        return Ok(None);
    }
    attribute_value(reader, element, b"ref")
}

/// `si` of a `<f t="shared">` element; None for ordinary formulas.
fn shared_formula_si<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<Option<u32>, SidecarError> {
    if attribute_value(reader, element, b"t")?.as_deref() != Some("shared") {
        return Ok(None);
    }
    Ok(attribute_value(reader, element, b"si")?.and_then(|value| value.parse::<u32>().ok()))
}

fn parse_address(address: &str) -> Result<(usize, usize), SidecarError> {
    let mut column = 0usize;
    let mut split = 0usize;
    for (index, byte) in address.bytes().enumerate() {
        if byte.is_ascii_alphabetic() {
            column = column
                .checked_mul(26)
                .and_then(|value| value.checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize))
                .ok_or_else(|| SidecarError::Workbook("Cell column overflows.".into()))?;
            split = index + 1;
        } else {
            break;
        }
    }
    if split == 0 || split == address.len() {
        return Err(SidecarError::Workbook(format!(
            "Invalid cell address {address}."
        )));
    }
    let row = address[split..]
        .parse::<usize>()
        .map_err(|_| SidecarError::Workbook(format!("Invalid cell address {address}.")))?;
    if row == 0 || column == 0 {
        return Err(SidecarError::Workbook(format!(
            "Invalid cell address {address}."
        )));
    }
    Ok((row - 1, column - 1))
}

fn read_zip_string(
    archive: &mut ZipArchive<File>,
    path: &str,
) -> Result<String, SidecarError> {
    let mut entry = archive.by_name(path)?;
    let mut value = String::new();
    entry.read_to_string(&mut value)?;
    Ok(value)
}

fn attribute_value<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, SidecarError> {
    for attribute in element.attributes().with_checks(false) {
        let attribute =
            attribute.map_err(|error| SidecarError::Workbook(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            return Ok(Some(
                attribute
                    .decode_and_unescape_value(reader.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// quick-xml 0.38 emits entity references (`&#26679;`, `&amp;`) as separate
/// GeneralRef events, not as part of Text. Exporters like openpyxl encode all
/// non-ASCII text as numeric character refs, so dropping these loses CJK text.
pub(crate) fn general_ref_text(
    reference: &quick_xml::events::BytesRef<'_>,
) -> Result<String, SidecarError> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?;
    Ok(match name.as_ref() {
        "amp" => "&".into(),
        "lt" => "<".into(),
        "gt" => ">".into(),
        "apos" => "'".into(),
        "quot" => "\"".into(),
        other => format!("&{other};"),
    })
}

fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String, SidecarError> {
    let decoded = text
        .decode()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?;
    quick_xml::escape::unescape(&decoded)
        .map(|value| value.into_owned())
        .map_err(|error| SidecarError::Workbook(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizes_worksheet_paths_with_forward_slashes() {
        assert_eq!(
            normalize_worksheet_path("worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            normalize_worksheet_path("./worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            normalize_worksheet_path("/xl/worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            normalize_worksheet_path("xl/worksheets/../worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert!(normalize_worksheet_path("../../etc/passwd").is_err());
    }

    #[test]
    fn parses_excel_addresses() {
        assert_eq!(parse_address("A1").unwrap(), (0, 0));
        assert_eq!(parse_address("XFD1048576").unwrap(), (1_048_575, 16_383));
    }

    #[test]
    fn rejects_oversized_ranges() {
        let sheet = SheetMetadata {
            id: "sheet-1".into(),
            name: "Sheet1".into(),
            row_count: 100_000,
            column_count: 100,
            column_widths: Vec::new(),
            default_row_height: None,
            default_column_width: None,
            freeze: None,
            hidden: false,
            tab_color: None,
            show_grid_lines: true,
            show_formulas: false,
            show_row_col_headers: true,
            tables: Vec::new(),
            comments: Vec::new(),
            pivot_ranges: Vec::new(),
            pivot_tables: Vec::new(),
            sparklines: Vec::new(),
        };
        let range = CellRange {
            start_row: 0,
            end_row: 999,
            start_column: 0,
            end_column: 99,
        };
        assert!(range.validate(&sheet).is_err());
    }

    fn sparklines_from(xml: &str) -> Vec<SparklineGroupInfo> {
        let mut reader = Reader::from_reader(xml.as_bytes());
        parse_sparkline_groups(&mut reader).unwrap()
    }

    fn sparkline_ext(groups_xml: &str) -> String {
        format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<extLst>
<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/></ext>
<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}">
<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">{groups_xml}</x14:sparklineGroups>
</ext>
</extLst>
</worksheet>"#
        )
    }

    #[test]
    fn parses_sparkline_group_types_and_colors() {
        let xml = sparkline_ext(
            r#"<x14:sparklineGroup displayEmptyCellsAs="gap">
<x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/>
<x14:sparklines><x14:sparkline><xm:f>Sheet1!A2:C2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline></x14:sparklines>
</x14:sparklineGroup>
<x14:sparklineGroup type="column">
<x14:colorSeries theme="4"/>
<x14:sparklines><x14:sparkline><xm:f>Sheet1!A3:C3</xm:f><xm:sqref>D3</xm:sqref></x14:sparkline></x14:sparklines>
</x14:sparklineGroup>
<x14:sparklineGroup type="stacked">
<x14:sparklines><x14:sparkline><xm:f>Sheet1!A4:C4</xm:f><xm:sqref>D4</xm:sqref></x14:sparkline></x14:sparklines>
</x14:sparklineGroup>
<x14:sparklineGroup type="bogus">
<x14:sparklines><x14:sparkline><xm:f>Sheet1!A5:C5</xm:f><xm:sqref>D5</xm:sqref></x14:sparkline></x14:sparklines>
</x14:sparklineGroup>"#,
        );
        let groups = sparklines_from(&xml);
        assert_eq!(groups.len(), 4);
        assert_eq!(groups[0].kind, "line");
        assert_eq!(groups[0].color.as_deref(), Some("#376092"));
        assert_eq!(groups[0].negative_color.as_deref(), Some("#D00000"));
        assert_eq!(groups[0].cells.len(), 1);
        assert_eq!(groups[0].cells[0].cell, "D2");
        assert_eq!(groups[0].cells[0].source_ref, "Sheet1!A2:C2");
        assert_eq!(groups[1].kind, "column");
        assert_eq!(groups[1].color, None);
        assert_eq!(groups[1].negative_color, None);
        assert_eq!(groups[2].kind, "stacked");
        assert_eq!(groups[3].kind, "line");
    }

    #[test]
    fn parses_multiple_sparkline_cells_and_skips_empty_groups() {
        let xml = sparkline_ext(
            r#"<x14:sparklineGroup/>
<x14:sparklineGroup type="line">
<x14:sparklines>
<x14:sparkline><xm:f>Sheet1!A2:C2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline>
<x14:sparkline><xm:f>Sheet1!A3:C3</xm:f><xm:sqref>D3</xm:sqref></x14:sparkline>
<x14:sparkline><xm:f>Sheet1!A4:C4</xm:f><xm:sqref>D4</xm:sqref></x14:sparkline>
</x14:sparklines>
</x14:sparklineGroup>"#,
        );
        let groups = sparklines_from(&xml);
        assert_eq!(groups.len(), 1);
        let cells: Vec<_> = groups[0].cells.iter().map(|entry| entry.cell.as_str()).collect();
        assert_eq!(cells, ["D2", "D3", "D4"]);
    }

    #[test]
    fn truncates_oversized_sparkline_groups() {
        let mut entries = String::new();
        for index in 0..(MAX_SPARKLINE_CELLS + 5) {
            entries.push_str(&format!(
                "<x14:sparkline><xm:f>Sheet1!A{row}:C{row}</xm:f><xm:sqref>D{row}</xm:sqref></x14:sparkline>",
                row = index + 1
            ));
        }
        let xml = sparkline_ext(&format!(
            "<x14:sparklineGroup><x14:sparklines>{entries}</x14:sparklines></x14:sparklineGroup>"
        ));
        let groups = sparklines_from(&xml);
        assert_eq!(groups[0].cells.len(), MAX_SPARKLINE_CELLS);

        let mut group_xml = String::new();
        for index in 0..(MAX_SPARKLINE_GROUPS + 3) {
            group_xml.push_str(&format!(
                "<x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A{row}:C{row}</xm:f><xm:sqref>D{row}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>",
                row = index + 1
            ));
        }
        let groups = sparklines_from(&sparkline_ext(&group_xml));
        assert_eq!(groups.len(), MAX_SPARKLINE_GROUPS);
    }

    #[test]
    fn ignores_worksheets_without_sparkline_ext() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><other/></ext></extLst>
</worksheet>"#;
        assert!(sparklines_from(xml).is_empty());
    }

    #[test]
    fn omits_empty_sparklines_field_from_metadata_json() {
        let sheet = SheetMetadata {
            id: "sheet-1".into(),
            name: "Sheet1".into(),
            row_count: 1,
            column_count: 1,
            column_widths: Vec::new(),
            default_row_height: None,
            default_column_width: None,
            freeze: None,
            hidden: false,
            tab_color: None,
            show_grid_lines: true,
            show_formulas: false,
            show_row_col_headers: true,
            tables: Vec::new(),
            comments: Vec::new(),
            pivot_ranges: Vec::new(),
            pivot_tables: Vec::new(),
            sparklines: Vec::new(),
        };
        let json = serde_json::to_string(&sheet).unwrap();
        assert!(!json.contains("sparklines"));
    }

    fn open_fixture(entries: &[(&str, &str)]) -> (tempfile::TempDir, PathBuf) {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("fixture.xlsx");
        let mut writer = zip::ZipWriter::new(File::create(&path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(*name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
        (dir, path)
    }

    /// Writers that never update `<dimension>` leave `ref="A1"` on sheets
    /// with real data (POI SXSSF et al.); the extent must be measured, not
    /// trusted, or every cell outside A1 is unreachable.
    #[test]
    fn stale_single_cell_dimension_is_remeasured() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1"/>
<sheetData><row r="3"><c r="C3"><v>7</v></c></row></sheetData>
</worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.sheets[0].row_count, 3);
        assert_eq!(metadata.sheets[0].column_count, 3);
    }

    /// A malformed dimension ref degrades to "measure the sheet", not a
    /// workbook-open failure.
    #[test]
    fn malformed_dimension_ref_is_ignored() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="1:20"/>
<sheetData><row r="2"><c r="B2"><v>1</v></c></row></sheetData>
</worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.sheets[0].row_count, 2);
        assert_eq!(metadata.sheets[0].column_count, 2);
    }

    /// OpenXML-SDK writers leave a stale single-row dimension (A1:G1) behind;
    /// trusting it hides every following data row.
    #[test]
    fn stale_single_row_dimension_is_remeasured() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:G1"/>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="4"><c r="I4"><v>2</v></c></row></sheetData>
</worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.sheets[0].row_count, 4);
        assert_eq!(metadata.sheets[0].column_count, 9);
    }

    /// An empty <v/> on a t="s" cell degrades to a valueless styled cell;
    /// erroring used to blank every cell of the sheet.
    #[test]
    fn empty_shared_string_value_keeps_sheet_readable() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Partner ID</t></si></sst>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v/></c><c r="C1" t="s"><v>99</v></c></row></sheetData></worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 0,
            start_column: 0,
            end_column: 2,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        assert_eq!(result.cells.len(), 1);
        assert!(matches!(&result.cells[0].value, Some(CellValue::String(text)) if text == "Partner ID"));
    }

    /// Fully `x:`-prefixed table parts (common .NET writers) must still parse.
    #[test]
    fn prefixed_table_part_is_parsed() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:tableParts count="1"><x:tablePart r:id="rId2"/></x:tableParts></x:worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
            ),
            (
                "xl/tables/table1.xml",
                r#"<x:table xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="ReportTable" displayName="ReportTable" ref="A1:I4" headerRowCount="1"><x:tableColumns count="2"><x:tableColumn id="1" name="Unit"/><x:tableColumn id="2" name="Goal"/></x:tableColumns><x:tableStyleInfo name="TableStyleMedium1" showRowStripes="1"/></x:table>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let table = &metadata.sheets[0].tables[0];
        assert_eq!(table.name.as_deref(), Some("ReportTable"));
        assert_eq!(table.columns, vec!["Unit", "Goal"]);
        assert_eq!(table.range.end_row, 3);
        assert_eq!(table.range.end_column, 8);
    }

    #[test]
    fn parses_workbook_pr_date1904() {
        let sheet_xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet_xml),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert!(metadata.date1904);
    }

    #[test]
    fn parses_workbook_view_active_tab() {
        let sheet_xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/worksheets/sheet2.xml", sheet_xml),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.active_tab, 1);
    }

    /// POI writes explicit off-flags as non-self-closing <strike val="0">;
    /// rich-run boolean properties must honor val, not mere presence.
    #[test]
    fn rich_run_bool_props_honor_val_zero() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b val="0"></b><i val="0"></i><strike val="0"></strike><u val="none"></u><sz val="10"/><rFont val="Arial"/></rPr><t>Last Run :</t></r><r><rPr><b/><strike/></rPr><t>on</t></r></si></sst>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 0,
            start_column: 0,
            end_column: 0,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let cell = &result.cells[0];
        let runs = cell.rich.as_ref().expect("rich runs");
        assert!(!runs[0].bold);
        assert!(!runs[0].italic);
        assert!(!runs[0].strikethrough);
        assert!(!runs[0].underline);
        assert!(runs[1].bold);
        assert!(runs[1].strikethrough);
    }

    /// An OLE embed's hidden drawing fallback gets the object's progId as
    /// its caption so the placeholder reads as an embedded object.
    #[test]
    fn ole_placeholder_carries_prog_id() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId2"/><oleObjects><oleObject progId="Acrobat Document" shapeId="1025" r:id="rId3"/></oleObjects></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            (
                "xl/drawings/drawing1.xml",
                r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1025" name="Object 1" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let shape = metadata
            .visuals
            .iter()
            .find(|visual| visual.kind == "shape")
            .expect("shape visual");
        assert_eq!(shape.text.as_deref(), Some("Acrobat Document"));
    }

    /// Phonetic ruby annotations (<rPh>) must not leak into cell text.
    #[test]
    fn phonetic_runs_stay_out_of_shared_strings() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>豊田</t><rPh sb="0" eb="1"><t>トヨ</t></rPh><rPh sb="1" eb="2"><t>タ</t></rPh><phoneticPr fontId="1"/></si></sst>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 0,
            start_column: 0,
            end_column: 0,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        assert!(
            matches!(&result.cells[0].value, Some(CellValue::String(text)) if text == "\u{8c4a}\u{7530}")
        );
    }

    /// A tableStyleInfo without a name is Excel's style "None": no palette.
    #[test]
    fn nameless_table_style_gets_no_palette() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></tableParts></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
            ),
            (
                "xl/tables/table1.xml",
                r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C2"><tableColumns count="1"><tableColumn id="1" name="C1"/></tableColumns><tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let table = &metadata.sheets[0].tables[0];
        assert!(table.style_name.is_none());
        assert!(table.header_fill.is_none());
        assert!(table.header_font_color.is_none());
        assert!(table.stripe_fill.is_none());
        assert!(table.border_color.is_none());
    }

    /// TableStyleLight1 draws its frame in the neutral dk1 color.
    #[test]
    fn light_table_style_carries_border_color() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></tableParts></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
            ),
            (
                "xl/tables/table1.xml",
                r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C3"><tableColumns count="1"><tableColumn id="1" name="C1"/></tableColumns><tableStyleInfo name="TableStyleLight1" showRowStripes="1"/></table>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let table = &metadata.sheets[0].tables[0];
        assert_eq!(table.border_color.as_deref(), Some("#000000"));
        // Light 1-7: unfilled header, font in the base color.
        assert!(table.header_fill.is_none());
        assert_eq!(table.header_font_color.as_deref(), Some("#000000"));
    }

    /// pivotTableStyleInfo Light-family styles resolve a band fill.
    #[test]
    fn pivot_style_light_resolves_band_fill() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="8"><c r="A8" t="str"><v>Row Labels</v></c></row></sheetData></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#,
            ),
            (
                "xl/pivotTables/pivotTable1.xml",
                r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1" rowGrandTotals="1"><location ref="A8:B11" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1"/></pivotTableDefinition>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let pivot = &metadata.sheets[0].pivot_tables[0];
        assert_eq!(pivot.output_ref, "A8:B11");
        assert_eq!(pivot.first_data_row, 1);
        assert!(pivot.row_grand_totals);
        // Light16 -> accent1 tint 0.8; no theme part, so the default accent.
        assert!(pivot.header_fill.is_some());
    }

    /// rowGrandTotals="0" must reach the renderer as an explicit false.
    #[test]
    fn pivot_disabled_grand_totals_serialize_false() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#,
            ),
            (
                "xl/pivotTables/pivotTable1.xml",
                r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1" rowGrandTotals="0"><location ref="A1:B3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotTableStyleInfo name="PivotStyleLight16"/></pivotTableDefinition>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let pivot = &metadata.sheets[0].pivot_tables[0];
        assert!(!pivot.row_grand_totals);
        let json = serde_json::to_string(pivot).unwrap();
        assert!(json.contains("\"rowGrandTotals\":false"));
    }

    /// The extLst twin <x14:table altText=…/> shares the local name "table";
    /// matching by local name reset the parsed range and dropped the table.
    #[test]
    fn keeps_tables_with_x14_alt_text_ext() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></tableParts></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
            ),
            (
                "xl/tables/table1.xml",
                r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblIncome" displayName="tblIncome" ref="B4:C7"><tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Amount"/></tableColumns><tableStyleInfo name="X" showRowStripes="1" showColumnStripes="0"/><extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:table altText="Monthly Income"/></ext></extLst></table>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let table = &metadata.sheets[0].tables[0];
        assert_eq!(table.name.as_deref(), Some("tblIncome"));
        assert_eq!(table.columns, vec!["Item", "Amount"]);
        assert_eq!(table.range.start_row, 3);
        assert_eq!(table.range.end_row, 6);
    }

    /// A custom `<tableStyle>` resolves its band fills/fonts from the dxf
    /// table instead of falling back to the built-in Medium2 approximation.
    #[test]
    fn custom_table_style_resolves_dxf_bands() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData><row r="1"><c r="A1" t="str"><v>h</v></c></row></sheetData>
<tableParts count="1"><tablePart r:id="rId2"/></tableParts>
</worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
            ),
            (
                "xl/tables/table1.xml",
                r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B4" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns><tableStyleInfo name="MyStyle" showRowStripes="1" showColumnStripes="1"/></table>"#,
            ),
            (
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font/></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
<dxfs count="3">
<dxf><font><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FF5B9BD5"/></patternFill></fill></dxf>
<dxf><fill><patternFill><bgColor rgb="FFDEEBF7"/></patternFill></fill></dxf>
<dxf><fill><patternFill><bgColor rgb="FFFFC000"/></patternFill></fill></dxf>
</dxfs>
<tableStyles count="1"><tableStyle name="MyStyle" pivot="0" count="3">
<tableStyleElement type="headerRow" dxfId="0"/>
<tableStyleElement type="firstRowStripe" dxfId="1"/>
<tableStyleElement type="totalRow" dxfId="2"/>
</tableStyle></tableStyles>
</styleSheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let table = &metadata.sheets[0].tables[0];
        assert_eq!(table.header_fill.as_deref(), Some("#5B9BD5"));
        assert_eq!(table.header_font_color.as_deref(), Some("#FF0000"));
        assert_eq!(table.stripe_fill.as_deref(), Some("#DEEBF7"));
        assert_eq!(table.total_row_fill.as_deref(), Some("#FFC000"));
        assert_eq!(table.totals_row_count, 1);
    }

    /// Chart title semantics + per-side axes + scatter/series line parsing,
    /// and text-box paragraphs with run styling, end to end through open().
    #[test]
    fn parses_chart_semantics_and_textbox_paragraphs() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId2"/></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            (
                "xl/drawings/drawing1.xml",
                r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>
<xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr anchor="t"/><a:p><a:pPr algn="r"/><a:r><a:rPr lang="en" sz="1100" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Hi</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>
</xdr:wsDr>"#,
            ),
            (
                "xl/drawings/_rels/drawing1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
            ),
            (
                "xl/charts/chart1.xml",
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title/><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:pt idx="0"><c:v>Demo</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:xVal><c:numRef><c:f>S!$A$2:$A$3</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>S!$B$2:$B$3</c:f><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart><c:valAx><c:axId val="1"/><c:scaling><c:min val="-180"/><c:max val="180"/></c:scaling><c:axPos val="b"/><c:majorGridlines/><c:title><c:tx><c:rich><a:p><a:r><a:t>XT</a:t></a:r></a:p></c:rich></c:tx></c:title><c:numFmt formatCode="0&quot;X&quot;" sourceLinked="0"/><c:majorUnit val="60"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:title/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let chart = metadata
            .visuals
            .iter()
            .find(|visual| visual.kind == "chart")
            .and_then(|visual| visual.chart.as_ref())
            .expect("chart visual");
        // Present-but-empty <c:title> with a single series → the series name.
        assert_eq!(chart.title, "Demo");
        assert_eq!(chart.scatter_style.as_deref(), Some("lineMarker"));
        let x_axis = chart.x_axis.as_ref().expect("x axis");
        assert_eq!(x_axis.title.as_deref(), Some("XT"));
        assert_eq!(x_axis.min, Some(-180.0));
        assert_eq!(x_axis.max, Some(180.0));
        assert_eq!(x_axis.major_unit, Some(60.0));
        assert_eq!(x_axis.num_fmt.as_deref(), Some("0\"X\""));
        assert!(x_axis.major_gridlines);
        // Present-but-empty axis title → the "Axis Title" placeholder.
        let y_axis = chart.y_axis.as_ref().expect("y axis");
        assert_eq!(y_axis.title.as_deref(), Some("Axis Title"));
        let series = &chart.series[0];
        assert_eq!(series.line_color.as_deref(), Some("none"));
        assert_eq!(series.marker.as_deref(), Some("none"));
        let shape = metadata
            .visuals
            .iter()
            .find(|visual| visual.kind == "shape")
            .expect("shape visual");
        assert_eq!(shape.line_color.as_deref(), Some("#808080"));
        assert_eq!(shape.text_anchor.as_deref(), Some("t"));
        assert_eq!(shape.text.as_deref(), Some("Hi\nSecond"));
        let paragraphs = shape.paragraphs.as_ref().expect("paragraphs");
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].align.as_deref(), Some("r"));
        assert_eq!(paragraphs[0].runs[0].text, "Hi");
        assert!(paragraphs[0].runs[0].bold);
        assert_eq!(paragraphs[0].runs[0].color.as_deref(), Some("#FF0000"));
    }

    /// `defaultColWidth="0"` (Excel's all-hidden-sheet form) is
    /// legal and means "no default"; it must not fail the whole workbook.
    #[test]
    fn zero_default_sizes_normalize_to_none() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="TTS" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetFormatPr defaultColWidth="0" defaultRowHeight="12.75" customHeight="1" zeroHeight="1"/>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.sheets[0].default_column_width, None);
        assert_eq!(metadata.sheets[0].default_row_height, Some(12.75));
    }

    /// Shared-formula followers (`<f t="shared" si="N"/>`) must
    /// inherit the master's formula with relative references shifted.
    #[test]
    fn expands_shared_formulas_into_followers() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // Master E3 with self-closing (F3) and open-empty (G3)
                // followers, plus an unrelated second group whose master
                // mixes absolute and relative references.
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="3">
<c r="E3"><f t="shared" ref="E3:G3" si="0">D3+1</f><v>2</v></c>
<c r="F3"><f t="shared" si="0"/><v>3</v></c>
<c r="G3"><f t="shared" si="0"></f><v>4</v></c>
</row>
<row r="6"><c r="E6"><f t="shared" ref="E6:E7" si="1">SUM($A$1:B5)</f><v>9</v></c></row>
<row r="7"><c r="E7"><f t="shared" si="1"/><v>9</v></c></row>
</sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 6,
            start_column: 0,
            end_column: 6,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let formula_at = |row: usize, column: usize| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .and_then(|cell| cell.formula.clone())
        };
        assert_eq!(formula_at(2, 4).as_deref(), Some("=D3+1")); // master E3
        assert_eq!(formula_at(2, 5).as_deref(), Some("=E3+1")); // follower F3
        assert_eq!(formula_at(2, 6).as_deref(), Some("=F3+1")); // follower G3
        assert_eq!(formula_at(5, 4).as_deref(), Some("=SUM($A$1:B5)")); // master E6
        assert_eq!(formula_at(6, 4).as_deref(), Some("=SUM($A$1:B6)")); // follower E7
        // The formula index (readWorkbookFormulas) sees the followers too.
        let formulas = sessions
            .read_formula_cells(&metadata.session_id, &sheet_id)
            .unwrap();
        assert_eq!(formulas.cells.len(), 5);
    }

    /// Excel writes ht on auto-fitted rows without customHeight="1"; those
    /// heights must survive parsing (dropping them forced a re-measure with
    /// substitute fonts, which clipped wrapped CJK rows on Windows).
    #[test]
    fn reports_row_heights_without_custom_height_flag() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // Row 1: Excel auto-fit height (no customHeight); row 2: an
                // explicit user height; row 3: no height attributes at all.
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1" ht="38.25"><c r="A1"><v>1</v></c></row>
<row r="2" ht="56" customHeight="1"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
</sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 2,
            start_column: 0,
            end_column: 0,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let row_of = |row: usize| result.rows.iter().find(|property| property.row == row);
        assert_eq!(row_of(0).and_then(|property| property.height), Some(38.25));
        assert!(!row_of(0).unwrap().custom_height);
        assert_eq!(row_of(1).and_then(|property| property.height), Some(56.0));
        assert!(row_of(1).unwrap().custom_height);
        assert!(row_of(2).is_none());
    }

    #[test]
    fn captures_array_formula_refs_on_master_cells() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // A1 is a CSE master spanning A1:A2 (follower A2 has only the
                // dead cached value); C1 is a single-cell array formula.
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1"><f t="array" ref="A1:A2">TRANSPOSE(B1:C1)</f><v>7</v></c>
<c r="B1"><v>7</v></c>
<c r="C1" t="str"><f t="array" ref="C1">B1&amp;""</f><v>7</v></c>
</row>
<row r="2"><c r="A2"><v>8</v></c></row>
</sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 2,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let cell_at = |row: usize, column: usize| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .unwrap()
        };
        let master = cell_at(0, 0);
        assert_eq!(master.formula.as_deref(), Some("=TRANSPOSE(B1:C1)"));
        assert_eq!(master.array_ref.as_deref(), Some("A1:A2"));
        let single = cell_at(0, 2);
        assert_eq!(single.formula.as_deref(), Some("=B1&\"\""));
        assert_eq!(single.array_ref.as_deref(), Some("C1"));
        // Followers stay plain cached values; plain cells carry no ref.
        assert_eq!(cell_at(1, 0).array_ref, None);
        assert_eq!(cell_at(1, 0).formula, None);
        assert_eq!(cell_at(0, 1).array_ref, None);
    }

    /// Value-less `<c s="…"/>` cells used to be dropped, losing
    /// borders/fills on blank and merged cells; later widened to
    /// any xf differing from the default (number format / font / alignment).
    #[test]
    fn keeps_value_less_cells_whose_style_paints_borders_or_fill() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // xf1: bordered, xf2: filled, xf3: bold-only (kept since #169).
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font/><font><b/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf borderId="1" applyBorder="1"/><xf fillId="2" applyFill="1"/><xf fontId="1" applyFont="1"/></cellXfs>
</styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/><c r="C1" s="1"></c><c r="D1" s="0"/><c r="E1"/><c r="F1" s="3"/></row>
<row r="2"><c r="A2" s="2"/></row>
</sheetData>
<mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let sheet_id = metadata.sheets[0].id.clone();
        // Merges are parsed after sheetData (OOXML order) and are sheet-wide
        // data that stays empty until worksheet indexing completes, while the
        // range wait only covers row indexing — poll until the index is
        // complete so the merge assertion is deterministic.
        let range = CellRange {
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 5,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };

        let mut cells: Vec<(usize, usize, bool, Option<usize>)> = result
            .cells
            .iter()
            .map(|cell| (cell.row, cell.column, cell.value.is_some(), cell.style_index))
            .collect();
        cells.sort();
        assert_eq!(
            cells,
            vec![
                (0, 0, true, None),
                // Self-closing and open-empty styled blanks both survive.
                (0, 1, false, Some(1)),
                (0, 2, false, Some(1)),
                // Font-only blank survives too: the format applies the moment
                // the user types into the cell (#169). D1 (s="0", the default
                // xf) and E1 (no style) stay dropped.
                (0, 5, false, Some(3)),
                // Filled blank anchoring the merged range survives.
                (1, 0, false, Some(2)),
            ]
        );
        assert_eq!(result.merges.len(), 1);
    }

    /// <row r=> and <c r=> are both optional (54288.xlsx omits r on all but
    /// each row's first cell); implicit addresses continue from the
    /// predecessor. Dimension inference must count them too.
    #[test]
    fn assigns_implicit_addresses_without_r() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // No <dimension>: the extent must be inferred from implicit
                // addresses as well. Row 3 anchors at C3 then continues
                // implicitly (D3, E3); the second row has no r (row 4) and
                // starts at column A.
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="3" spans="3:5"><c r="C3"><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>
<row><c><v>9</v></c></row>
</sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        assert_eq!(metadata.sheets[0].row_count, 4);
        assert_eq!(metadata.sheets[0].column_count, 5);
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 3,
            start_column: 0,
            end_column: 4,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let mut cells: Vec<(usize, usize, f64)> = result
            .cells
            .iter()
            .filter_map(|cell| match cell.value {
                Some(CellValue::Number(number)) => Some((cell.row, cell.column, number)),
                _ => None,
            })
            .collect();
        cells.sort_by(|a, b| a.partial_cmp(b).unwrap());
        assert_eq!(
            cells,
            vec![(2, 2, 1.0), (2, 3, 2.0), (2, 4, 3.0), (3, 0, 9.0)]
        );
    }

    /// <row s= customFormat="1"> and <col style=> carry the default xf for
    /// cells without one of their own; both used to be dropped entirely.
    #[test]
    fn reports_row_and_column_default_styles() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF000080"/></patternFill></fill></fills>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fillId="0"/><xf fillId="2" applyFill="1"/><xf fillId="2" applyFill="1"/></cellXfs>
</styleSheet>"#,
            ),
            (
                // Row 1 styles via customFormat; row 2's s without customFormat
                // is noise Excel ignores. The style-only <col> used to be
                // dropped with it.
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cols><col min="2" max="3" style="2"/><col min="5" max="5" width="12" customWidth="1"/></cols>
<sheetData>
<row r="1" s="1" customFormat="1"><c r="A1"><v>1</v></c></row>
<row r="2" s="1"><c r="A2"><v>2</v></c></row>
</sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let widths = &metadata.sheets[0].column_widths;
        let styled = widths
            .iter()
            .find(|width| width.style_index.is_some())
            .expect("style-only <col> must be kept");
        assert_eq!(
            (styled.start_column, styled.end_column, styled.style_index),
            (1, 2, Some(2))
        );
        assert!(widths
            .iter()
            .any(|width| width.width == Some(12.0) && width.style_index.is_none()));
        let sheet_id = metadata.sheets[0].id.clone();
        let range = CellRange {
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 0,
        };
        let result = loop {
            let result = sessions
                .read_range(&metadata.session_id, &sheet_id, &range)
                .unwrap();
            if result.indexing_complete {
                break result;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        let style_of = |row: usize| {
            result
                .rows
                .iter()
                .find(|property| property.row == row)
                .and_then(|property| property.style_index)
        };
        assert_eq!(style_of(0), Some(1));
        assert_eq!(style_of(1), None);
    }

    /// zh Excel/WPS date cells reference locale-reserved builtin numFmtIds
    /// (27-36, 50-58) that carry no formatCode in styles.xml; they used to
    /// resolve to None and render as raw date serials (e.g. 46230).
    #[test]
    fn resolves_locale_reserved_builtin_number_formats() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                // xf1: zh date (58), xf2: accounting (44), xf3: an explicit
                // numFmt entry reusing a reserved id, which must win over
                // the builtin table.
                "xl/styles.xml",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="57" formatCode="yyyy/m/d"/></numFmts>
<fonts count="1"><font/></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="58" applyNumberFormat="1"/><xf numFmtId="44" applyNumberFormat="1"/><xf numFmtId="57" applyNumberFormat="1"/><xf numFmtId="27" applyNumberFormat="1"/><xf numFmtId="30" applyNumberFormat="1"/></cellXfs>
</styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="1"><v>46230</v></c><c r="B1" s="4"><v>45651</v></c><c r="C1" s="5"><v>45343</v></c></row></sheetData>
</worksheet>"#,
            ),
        ]);

        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        let format = |index: usize| metadata.styles[index].number_format.as_deref();
        // The zh-CN month/day date format (U+6708 month, U+65E5 day),
        // escaped to keep the source ASCII-only.
        assert_eq!(format(1), Some("m\"\u{6708}\"d\"\u{65e5}\""));
        assert_eq!(
            format(2),
            Some(r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#),
        );
        assert_eq!(format(3), Some("yyyy/m/d"));

        // Locale-reserved ids must follow the UI locale. In Portuguese (and
        // other day-first locales), id 58 is a full date rather than the
        // zh-CN month/day pattern that previously leaked into every workbook.
        let mut portuguese_sessions = WorkbookSessions::new();
        let portuguese = portuguese_sessions.open_with_locale(&path, "pt").unwrap();
        let portuguese_format =
            |index: usize| portuguese.styles[index].number_format.as_deref();
        assert_eq!(portuguese_format(1), Some("d/m/yyyy"));
        assert_eq!(portuguese_format(3), Some("yyyy/m/d"));
        assert_eq!(portuguese_format(4), Some("d/m/yyyy"));
        assert_eq!(portuguese_format(5), Some("d/m/yyyy"));
    }

    #[test]
    fn serializes_sparklines_in_expected_shape() {
        let group = SparklineGroupInfo {
            kind: "line".into(),
            color: Some("#376092".into()),
            negative_color: None,
            cells: vec![SparklineCellInfo {
                cell: "D2".into(),
                source_ref: "Sheet1!A2:C2".into(),
            }],
        };
        let json = serde_json::to_value(&group).unwrap();
        assert_eq!(
            json,
            serde_json::json!({
                "type": "line",
                "color": "#376092",
                "cells": [{ "cell": "D2", "sourceRef": "Sheet1!A2:C2" }]
            })
        );
    }
}

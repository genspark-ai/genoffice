use std::collections::{HashMap, HashSet};
use std::fs::{self, File};
use std::io::{BufReader, BufWriter, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::Duration;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use serde::{Deserialize, Serialize};
use uuid::Uuid;
use zip::ZipArchive;

pub mod archive;
pub mod convert;
pub mod recalc;
mod richdata;
mod shared_formulas;
mod visuals;

mod cell_images;
mod refs;
mod sparklines;
mod table_styles;
#[cfg(test)]
mod tests;
mod types;
mod workbook;
mod worksheet;
mod xml_util;

use cell_images::*;
use refs::*;
use sparklines::*;
use table_styles::*;
pub use types::*;
use workbook::*;
use worksheet::*;
use xml_util::*;

pub use visuals::{CellStyle, MediaResult, ThemeFonts, VisualObject};
use visuals::{ColorContext, SheetVisualSource};

const CHUNK_ROW_COUNT: usize = 256;
const MAX_RANGE_CELLS: usize = 100_000;
const RANGE_WAIT: Duration = Duration::from_millis(750);
/// Reads this far past the indexed row would only burn the full RANGE_WAIT;
/// answer immediately instead and let the caller poll.
const RANGE_WAIT_MAX_LAG_ROWS: usize = 16 * CHUNK_ROW_COUNT;
const MAX_FORMULA_CELLS: usize = 100_000;
const MAX_ENTRY_COUNT: usize = 10_000;

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
        self.open_with_locale(path, "zh", None)
    }

    pub fn open_with_locale(
        &mut self,
        path: &Path,
        locale: &str,
        short_date_format: Option<&str>,
    ) -> Result<WorkbookMetadata, SidecarError> {
        let canonical_path = path.canonicalize()?;
        let file = File::open(&canonical_path)?;
        let mut archive = ZipArchive::new(file)?;
        validate_archive(&mut archive)?;
        let entry_count = archive.len();
        let mut color_context = visuals::read_theme_palette(&mut archive)?;
        visuals::read_indexed_palette(&mut archive, &mut color_context)?;
        let theme_colors = color_context.palette_hex();
        let theme_fonts = visuals::read_theme_fonts(&mut archive)?;
        let shared_strings = Arc::new(read_shared_strings(&mut archive, &color_context)?);
        let (declarations, workbook_protection, active_tab, date_1904) =
            read_sheet_declarations(&mut archive)?;
        let relationships = read_workbook_relationships(&mut archive)?;
        let (styles, dxf_styles, normal_font_name) = visuals::read_styles(
            &mut archive,
            &color_context,
            theme_fonts.as_ref(),
            locale,
            short_date_format,
        )?;
        let custom_table_styles = read_custom_table_styles(&mut archive, &dxf_styles);
        let rich_value_images = richdata::read_rich_value_images(&mut archive);
        let mut cell_image_count = 0usize;
        let mut sheets = Vec::with_capacity(declarations.len());
        let mut runtimes = Vec::with_capacity(declarations.len());
        let mut visual_sources = Vec::with_capacity(declarations.len());

        for declaration in declarations {
            let target = relationships
                .get(&declaration.relationship_id)
                .ok_or_else(|| {
                    SidecarError::Workbook(format!(
                        "Missing worksheet relationship {}.",
                        declaration.relationship_id
                    ))
                })?;
            let worksheet_path = normalize_worksheet_path(target)?;
            let source_xml_bytes = zip_entry(&mut archive, &worksheet_path)?.size();
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
            let cell_images = read_sheet_cell_images(
                &mut archive,
                &worksheet_path,
                &rich_value_images,
                &mut cell_image_count,
            )?;
            let pivot_infos = visuals::read_pivot_tables(&mut archive, &worksheet_path)?;
            let pivot_ranges = pivot_infos
                .iter()
                .filter_map(|info| parse_range_reference(&info.output_ref))
                .collect();
            let pivot_tables = pivot_infos
                .into_iter()
                .map(|info| {
                    let mut palette =
                        pivot_style_palette(info.style_name.as_deref(), &color_context);
                    if !info.show_row_stripes {
                        palette.stripe_fill = None;
                        palette.second_row_stripe_fill = None;
                    }
                    if !info.show_col_stripes {
                        palette.column_stripe_fill = None;
                        palette.second_column_stripe_fill = None;
                    }
                    PivotTableInfo {
                        palette,
                        styled: info
                            .style_name
                            .as_deref()
                            .is_some_and(|name| !name.is_empty()),
                        path: info.path,
                        cache_path: info.cache_path,
                        output_ref: info.output_ref,
                        first_data_row: info.first_data_row,
                        first_data_col: info.first_data_col,
                        row_grand_totals: info.row_grand_totals,
                        row_kinds: info.row_kinds,
                    }
                })
                .collect();
            let id = format!("sheet-{}", declaration.sheet_id);
            // A table may extend past the last written cell (a header-only
            // sheet whose table spans empty data rows); its banding and
            // frame still render there, so the grid must reach it.
            let (mut row_count, mut column_count) = (dimensions.row_count, dimensions.column_count);
            for table in &tables {
                row_count = row_count.max(table.range.end_row + 1);
                column_count = column_count.max(table.range.end_column + 1);
            }
            sheets.push(SheetMetadata {
                id: id.clone(),
                name: declaration.name,
                row_count,
                column_count,
                source_xml_bytes,
                column_widths: dimensions.column_widths,
                default_row_height: dimensions.default_row_height,
                default_row_height_fixed: dimensions.default_row_height_fixed,
                default_column_width: dimensions.default_column_width,
                base_column_width: dimensions.base_column_width,
                freeze: dimensions.freeze,
                hidden: declaration.hidden,
                tab_color: dimensions.tab_color,
                show_grid_lines: dimensions.show_grid_lines,
                show_formulas: dimensions.show_formulas,
                show_row_col_headers: dimensions.show_row_col_headers,
                right_to_left: dimensions.right_to_left,
                zoom_scale: dimensions.zoom_scale,
                tables,
                comments,
                pivot_ranges,
                pivot_tables,
                sparklines,
                print_area: None,
                print_titles: None,
                has_scoped_defined_names: false,
                cell_images,
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
        let (defined_names, print_names, scoped_sheets) = read_defined_names(&mut archive)?;
        for (sheet_index, sheet) in sheets.iter_mut().enumerate() {
            if let Some(names) = print_names.get(&sheet_index) {
                sheet.print_area = names.area.clone();
                sheet.print_titles = names.titles.clone();
            }
            sheet.has_scoped_defined_names = scoped_sheets.contains(&sheet_index);
        }

        let session_id = Uuid::new_v4().to_string();
        let cache_directory = std::env::temp_dir().join(format!("genspark-ai-excel-{session_id}"));
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
            normal_font_name,
            workbook_protection,
            date1904: date_1904,
            short_date_format: short_date_format.map(ToOwned::to_owned),
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

    /// The workbook file behind a session, for cross-subsystem cleanup
    /// (the recalc cache keys resident models by path).
    pub fn session_path(&self, session_id: &str) -> Option<PathBuf> {
        self.sessions
            .get(session_id)
            .map(|session| session.path.clone())
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
            .and_then(|visual| {
                visual
                    .media_path
                    .as_deref()
                    .or(visual.fill_media_path.as_deref())
            })
            .or_else(|| {
                session
                    .sheets
                    .iter()
                    .flat_map(|sheet| sheet.cell_images.iter())
                    .find(|image| image.id == visual_id)
                    .map(|image| image.media_path.as_str())
            })
            .map(str::to_owned)
            .or_else(|| session.header_footer_picture_path(visual_id))
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook image.".into()))?;
        let file = File::open(&session.path)?;
        let mut archive = ZipArchive::new(file)?;
        visuals::read_media(&mut archive, &media_path)
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
    /// Media path behind a header/footer picture id (`&G` slots live on the
    /// indexed sheet state, so only completed sheets can resolve one).
    fn header_footer_picture_path(&self, picture_id: &str) -> Option<String> {
        self.runtimes.iter().find_map(|runtime| {
            let (lock, _) = &*runtime.state;
            let index = lock.lock().ok()?;
            index
                .page_setup
                .as_ref()?
                .header_footer_pictures
                .iter()
                .find(|picture| picture.id == picture_id)
                .map(|picture| picture.media_path.clone())
        })
    }

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
        // The capped overlay list is the single source of truth: only cells
        // that actually got a picture record lose their cached error.
        let rich_image_cells: Arc<HashSet<(usize, usize)>> = Arc::new(
            self.sheets[sheet_index]
                .cell_images
                .iter()
                .map(|image| (image.row, image.column))
                .collect(),
        );
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
                    &rich_image_cells,
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
                        && current
                            .indexed_through_row
                            .is_none_or(|row| row < range.end_row)
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
        let page_setup = if indexing_complete {
            index.page_setup.clone()
        } else {
            None
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
            page_setup,
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
    page_setup: Option<PagePrintInfo>,
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

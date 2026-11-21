use std::collections::{HashMap, HashSet, VecDeque};
use std::fs::{self, File};
use std::io::{BufReader, BufWriter, Read, Seek, Write};
#[cfg(unix)]
use std::os::unix::fs::DirBuilderExt;
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
mod filter;
pub mod find;
mod formula_dependencies;
mod index_scheduler;
pub mod odf;
pub mod protocol;
pub mod recalc;
mod rels_normalize;
mod richdata;
mod shared_formulas;
mod shared_strings;
pub mod spreadsheetml;
pub mod stream_patch;
mod visuals;

mod cell_images;
mod chunks;
mod escher;
mod external_links;
mod refs;
mod sparklines;
mod table_styles;
#[cfg(test)]
mod tests;
mod types;
mod workbook;
mod worksheet;
mod xls_biff;
mod xls_styles;
mod xlsx_package;
mod xml_util;

use cell_images::*;
use formula_dependencies::DependencyCatalog;
use index_scheduler::{IndexScheduler, IndexState, IndexWorker, MetadataJob};
use refs::*;
use shared_strings::SharedStrings;
use sparklines::*;
use table_styles::*;
pub use types::*;
use workbook::*;
use worksheet::*;
use xml_util::*;

pub use visuals::{CellStyle, MediaResult, ThemeFonts, VisualObject};
use visuals::{ColorContext, SheetVisualSource, SourceFormats};

const CHUNK_ROW_COUNT: usize = 256;
const PREFETCH_CACHE_BUDGET: usize = 256 * 1024 * 1024;
const MIN_PREFETCH_BYTES_PER_SHEET: usize = 64 * 1024 * 1024;
const MAX_RANGE_CELLS: usize = 100_000;
const RANGE_WAIT: Duration = Duration::from_millis(750);
/// Reads this far past the indexed row would only burn the full RANGE_WAIT;
/// answer immediately instead and let the caller poll.
const RANGE_WAIT_MAX_LAG_ROWS: usize = 16 * CHUNK_ROW_COUNT;
const MAX_FORMULA_CELLS: usize = 100_000;

#[derive(Debug, Default)]
struct ChunkData {
    cells: Vec<CellRecord>,
    rows: Vec<RowProperty>,
}

#[derive(Debug)]
pub enum SidecarError {
    InvalidRequest(String),
    Io(String),
    Workbook(String),
    /// A cancel that landed after the request was already executing. Kept
    /// apart from a failure so the host can tell an abandoned open from a
    /// workbook it could not read.
    Cancelled,
}

impl std::fmt::Display for SidecarError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidRequest(message) | Self::Io(message) | Self::Workbook(message) => {
                formatter.write_str(message)
            }
            Self::Cancelled => formatter.write_str("Request was cancelled by the client."),
        }
    }
}

impl SidecarError {
    /// The cooperative abort an in-flight request reports.
    pub fn cancelled() -> Self {
        Self::Cancelled
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

/// Owns an open's cache directory from the moment it is created until the
/// open commits it to a session. Drop-based so every exit after creation — a
/// cooperative cancel, a failed read, a panic — reclaims the directory; only
/// `commit` lets it outlive the open, and a committed one is removed by
/// `WorkbookSession::close`.
#[derive(Debug)]
struct CacheDirectory {
    path: PathBuf,
    committed: bool,
}

impl CacheDirectory {
    fn create(path: PathBuf) -> Result<Self, SidecarError> {
        let mut builder = fs::DirBuilder::new();
        #[cfg(unix)]
        builder.mode(0o700);
        builder.create(&path)?;
        Ok(Self {
            path,
            committed: false,
        })
    }

    fn path(&self) -> &Path {
        &self.path
    }

    /// Hands the directory over to a session; only now may it outlive the open.
    fn commit(&mut self) {
        self.committed = true;
    }
}

impl Drop for CacheDirectory {
    fn drop(&mut self) {
        if !self.committed && self.path.exists() {
            let _ = fs::remove_dir_all(&self.path);
        }
    }
}

/// Whether this build can run work on background threads. WebAssembly
/// without threads (wasm32-wasip1, a browser without cross-origin isolation)
/// and the `single-thread` feature run the same work inline: a sheet is
/// indexed completely the first time it is read, and a recalc answers before
/// the next request is read.
pub const THREADS: bool = !cfg!(any(
    feature = "single-thread",
    all(target_family = "wasm", not(target_feature = "atomics"))
));

/// `std::env::temp_dir` panics on WASI, which has no notion of one: there the
/// host maps a directory and names it in `TMPDIR` (`/tmp` when unset).
fn temp_dir() -> PathBuf {
    if cfg!(target_os = "wasi") {
        std::env::var_os("TMPDIR").map_or_else(|| PathBuf::from("/tmp"), PathBuf::from)
    } else {
        std::env::temp_dir()
    }
}

/// `Path::canonicalize`, which WASI does not support: there the host hands over absolute
/// paths in a sandbox without symlinks or a working directory, so the path is already
/// canonical and only has to exist.
pub(crate) fn canonicalize(path: &Path) -> std::io::Result<PathBuf> {
    if cfg!(target_os = "wasi") {
        fs::metadata(path)?;
        Ok(path.to_path_buf())
    } else {
        path.canonicalize()
    }
}

/// Creates the cache directory an open needs. This is the one boundary where
/// the open touches the filesystem, so it is the one abort that happens after
/// the directory exists: returning here drops the guard, which reclaims it,
/// making the cancel path as clean as the success path.
fn create_cache_directory(
    session_id: &str,
    cancelled: &AtomicBool,
) -> Result<CacheDirectory, SidecarError> {
    let path = temp_dir().join(format!("genspark-ai-excel-{session_id}"));
    let directory = CacheDirectory::create(path)?;
    if cancelled.load(Ordering::Acquire) {
        return Err(SidecarError::cancelled());
    }
    Ok(directory)
}

pub struct WorkbookSessions {
    sessions: HashMap<String, WorkbookSession>,
}

/// Excel wraps password-encrypted workbooks in an OLE2 compound container
/// rather than a ZIP package. Detect that signature before the zip reader
/// runs so the open fails with an explanation instead of "invalid Zip
/// archive: Could not find EOCD" (genoffice#2076). The file cursor is rewound so the
/// caller can keep reading from the start.
fn is_ole2_container(file: &mut File) -> std::io::Result<bool> {
    const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut magic = [0u8; 8];
    let read = file.read(&mut magic)?;
    file.rewind()?;
    Ok(read == OLE2_MAGIC.len() && magic == OLE2_MAGIC)
}

impl WorkbookSessions {
    pub fn new() -> Self {
        Self {
            sessions: HashMap::new(),
        }
    }

    #[cfg(test)]
    /// Sessions currently registered. A cancelled open must leave none, or a
    /// later read-range would hit a session the host never opened.
    fn session_count(&self) -> usize {
        self.sessions.len()
    }

    pub fn open(&mut self, path: &Path) -> Result<WorkbookMetadata, SidecarError> {
        self.open_with_locale(path, "zh", None, &AtomicBool::new(false))
    }

    /// Opens a workbook, polling `cancelled` at stage boundaries so a cancel
    /// that lands once the host has already dequeued this request still
    /// abandons the open instead of running it to completion for a result
    /// nobody will read. A cancel arriving past the last boundary lets the
    /// open finish, as before.
    pub fn open_with_locale(
        &mut self,
        path: &Path,
        locale: &str,
        short_date_format: Option<&str>,
        cancelled: &AtomicBool,
    ) -> Result<WorkbookMetadata, SidecarError> {
        self.open_with_options(path, locale, short_date_format, false, cancelled)
    }

    pub(crate) fn open_with_options(
        &mut self,
        path: &Path,
        locale: &str,
        short_date_format: Option<&str>,
        progressive: bool,
        cancelled: &AtomicBool,
    ) -> Result<WorkbookMetadata, SidecarError> {
        // Before any work: the host may have cancelled before this ran at all.
        if cancelled.load(Ordering::Acquire) {
            return Err(SidecarError::cancelled());
        }
        let canonical_path = canonicalize(path)?;
        let mut file = File::open(&canonical_path)?;
        if is_ole2_container(&mut file)? {
            return Err(SidecarError::Workbook(
                "This workbook is password-protected and must be decrypted by the host application.".into(),
            ));
        }
        let mut archive = ZipArchive::new(file)?;
        let workbook_path = canonical_path.clone();
        archive::validate_entries(&mut archive)?;
        // The archive is open and validated and nothing has been written to
        // disk yet, so abandoning here costs only the parse already done.
        if cancelled.load(Ordering::Acquire) {
            return Err(SidecarError::cancelled());
        }
        let entry_count = archive.len();
        let mut color_context = visuals::read_theme_palette(&mut archive)?;
        visuals::read_indexed_palette(&mut archive, &mut color_context)?;
        let theme_colors = color_context.palette_hex();
        let theme_fonts = visuals::read_theme_fonts(&mut archive)?;
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
        let wps_cell_images = read_wps_cell_images(&mut archive);
        let mut cell_image_count = 0usize;
        let mut note_image_count = 0usize;
        let mut sheets = Vec::with_capacity(declarations.len());
        let mut runtimes = Vec::with_capacity(declarations.len());
        let mut visual_sources = Vec::with_capacity(declarations.len());
        let mut sheet_names = Vec::with_capacity(declarations.len());
        let mut stored_cell_budget = STORED_CELL_COUNT_BUDGET;
        let sources =
            worksheet_open_sources(&mut archive, &declarations, &relationships, cancelled)?;
        // Drawings (notably source-linked chart formats) and place-in-cell
        // images depend on the eager visual metadata path.
        let progressive = progressive
            && THREADS
            && rich_value_images.is_empty()
            && wps_cell_images.is_empty()
            && !archive.file_names().filter_map(Result::ok).any(|name| {
                name.starts_with("xl/drawings/")
                    || name.starts_with("xl/charts/")
                    || name.starts_with("xl/embeddings/")
            });
        let mut headers = Vec::with_capacity(sources.len());
        let mut untrusted_extents = Vec::with_capacity(sources.len());
        for source in &sources {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::cancelled());
            }
            let can_stream = progressive
                && source.is_large()
                && !visuals::read_relationships(&mut archive, &source.path)?
                    .values()
                    .any(|rel| {
                        let kind = &rel.relationship_type;
                        kind.ends_with("/drawing") || kind.ends_with("/oleObject")
                    });
            let header = if can_stream {
                read_progressive_sheet_header(&mut archive, &source.path, &color_context)?
            } else {
                None
            };
            untrusted_extents.push(can_stream && header.is_none());
            headers.push(header);
        }
        let parallelism = if THREADS && sources.len() > 1 {
            thread::available_parallelism().map_or(1, |count| count.get().min(4))
        } else {
            1
        };
        let mut scanned_metadata = VecDeque::new();

        for (index, declaration) in declarations.into_iter().enumerate() {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::cancelled());
            }
            let source = &sources[index];
            let worksheet_path = &source.path;
            // Scan the XML surrounding sheetData in small batches. These scans
            // are independent; the cell-count budget and metadata assembly
            // still run in workbook order.
            if scanned_metadata.is_empty()
                && parallelism > 1
                && source.is_large()
                && headers[index].is_none()
            {
                let end = (index + parallelism).min(sources.len());
                scanned_metadata = scan_worksheet_metadata_batch(
                    &workbook_path,
                    &sources[index..end],
                    &headers[index..end],
                    &untrusted_extents[index..end],
                    &color_context,
                    cancelled,
                );
            }
            let header = headers[index].take();
            let progressive = header.is_some();
            let mut scanned = scanned_metadata.pop_front().flatten().transpose()?;
            let dimensions = match header.or_else(|| scanned.as_mut()?.dimensions.take()) {
                Some(header) => header,
                None => read_sheet_dimensions(
                    &mut archive,
                    worksheet_path,
                    &color_context,
                    &mut stored_cell_budget,
                )?,
            };
            let tables = read_sheet_tables(
                &mut archive,
                worksheet_path,
                &color_context,
                &custom_table_styles,
            )?;
            let comments = visuals::read_comments(&mut archive, worksheet_path)?
                .into_iter()
                .filter_map(|comment| {
                    let anchor = comment
                        .reference
                        .split(':')
                        .next()
                        .unwrap_or(&comment.reference);
                    let (row, column) = parse_address(&anchor.replace('$', "")).ok()?;
                    let image = comment.image_media_path.map(|media_path| {
                        note_image_count += 1;
                        CommentImageInfo {
                            id: format!("note-picture-{note_image_count}"),
                            media_path,
                        }
                    });
                    Some(CommentInfo {
                        row,
                        column,
                        author: comment.author,
                        text: comment.text,
                        thread: comment.thread,
                        image,
                    })
                })
                .collect();
            let (protection, sparklines, ole_objects) = if progressive {
                (SheetProtectionMetadata::default(), Vec::new(), Vec::new())
            } else {
                let xml = match scanned {
                    Some(scanned) => scanned.envelope,
                    None => visuals::read_worksheet_without_cells(
                        &mut archive,
                        worksheet_path,
                        cancelled,
                    )?,
                };
                (
                    read_sheet_protection(&xml)?,
                    read_sheet_sparklines(&xml)?,
                    visuals::read_ole_objects(&mut archive, &xml, worksheet_path)?,
                )
            };
            let cell_images = read_sheet_cell_images(
                &mut archive,
                worksheet_path,
                &rich_value_images,
                &wps_cell_images,
                &mut cell_image_count,
            )?;
            let pivot_infos = visuals::read_pivot_tables(&mut archive, worksheet_path)?;
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
            sheet_names.push((declaration.name.clone(), worksheet_path.clone()));
            sheets.push(SheetMetadata {
                id: id.clone(),
                name: declaration.name,
                row_count,
                column_count,
                source_xml_bytes: source.xml_bytes,
                stored_cell_count: dimensions.stored_cell_count,
                value_cell_count: dimensions.value_cell_count,
                value_row_count: dimensions.value_extent.map(|(rows, _)| rows),
                value_column_count: dimensions.value_extent.map(|(_, columns)| columns),
                column_widths: dimensions.column_widths,
                default_row_height: dimensions.default_row_height,
                default_row_height_fixed: dimensions.default_row_height_fixed,
                default_column_width: dimensions.default_column_width,
                base_column_width: dimensions.base_column_width,
                freeze: dimensions.freeze,
                hidden: declaration.hidden,
                very_hidden: declaration.very_hidden,
                tab_color: dimensions.tab_color,
                show_grid_lines: dimensions.show_grid_lines,
                show_formulas: dimensions.show_formulas,
                show_row_col_headers: dimensions.show_row_col_headers,
                right_to_left: dimensions.right_to_left,
                zoom_scale: dimensions.zoom_scale,
                outline_level_row: dimensions.outline_level_row,
                outline_level_col: dimensions.outline_level_col,
                outline_summary_below: dimensions.outline_summary_below,
                outline_summary_right: dimensions.outline_summary_right,
                tables,
                comments,
                pivot_ranges,
                pivot_tables,
                sparklines,
                print_area: None,
                print_titles: None,
                has_scoped_defined_names: false,
                cell_images,
                progressive,
                protection,
            });
            visual_sources.push(SheetVisualSource {
                sheet_id: id,
                worksheet_path: worksheet_path.clone(),
                ole_objects,
            });
            runtimes.push((
                worksheet_path.clone(),
                Arc::new((Mutex::new(SheetIndex::default()), Condvar::new())),
            ));
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
        let mut source_formats = SourceFormats::new(sheet_names, &styles);
        let visual_objects = visuals::read_visual_objects(
            &mut archive,
            &visual_sources,
            &color_context,
            &mut source_formats,
        )?;
        let (defined_names, print_names, scoped_sheets) = read_defined_names(&mut archive)?;
        for (sheet_index, sheet) in sheets.iter_mut().enumerate() {
            if let Some(names) = print_names.get(&sheet_index) {
                sheet.print_area = names.area.clone();
                sheet.print_titles = names.titles.clone();
            }
            sheet.has_scoped_defined_names = scoped_sheets.contains(&sheet_index);
        }

        // Last boundary before the open touches the filesystem; past here the
        // only way to stop it is create_cache_directory's own poll.
        if cancelled.load(Ordering::Acquire) {
            return Err(SidecarError::cancelled());
        }
        let session_id = Uuid::new_v4().to_string();
        let mut cache_directory = create_cache_directory(&session_id, cancelled)?;
        let name = canonical_path
            .file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("workbook.xlsx")
            .to_owned();
        let progressive = sheets.iter().any(|sheet| sheet.progressive);
        let session_cancelled = Arc::new(AtomicBool::new(false));
        let shared_strings = SharedStrings::open(
            &mut archive,
            &workbook_path,
            cache_directory.path(),
            &color_context,
            progressive,
            Arc::clone(&session_cancelled),
        )?;
        if cancelled.load(Ordering::Acquire) {
            return Err(SidecarError::cancelled());
        }
        let sources = Arc::new(IndexSources {
            path: workbook_path,
            cache_directory: cache_directory.path().to_path_buf(),
            shared_strings,
            styled_xfs,
            color_context,
            cancelled: session_cancelled,
            dependencies: (THREADS && progressive)
                .then(|| DependencyCatalog::new(&sheets, &defined_names)),
        });
        let workers = runtimes
            .iter()
            .zip(&sheets)
            .enumerate()
            .map(|(index, ((path, state), sheet))| {
                sources.worker(index, path, Arc::clone(state), sheet)
            })
            .collect();
        let visible = sheets
            .iter()
            .enumerate()
            .filter_map(|(index, sheet)| (!sheet.hidden).then_some(index))
            .collect();
        let scheduler = IndexScheduler::new(workers, visible, Arc::clone(&sources.cancelled));
        let mut session = WorkbookSession {
            path: sources.path.clone(),
            sheets: sheets.clone(),
            runtimes: runtimes.into_iter().map(|(_, state)| state).collect(),
            filter_columns: Default::default(),
            visuals: visual_objects.clone(),
            cache_directory: sources.cache_directory.clone(),
            progressive,
            scheduler,
        };
        if progressive
            && let Some(active) = std::iter::once(active_tab)
                .chain(0..sheets.len())
                .find(|&index| sheets.get(index).is_some_and(|sheet| !sheet.hidden))
        {
            session.scheduler.focus(active);
            session.ensure_parser(active)?;
            session.scheduler.ensure_metadata(active);
        }
        self.sessions.insert(session_id.clone(), session);
        // The session now owns the directory; Close reclaims it. Committing
        // only after the insert keeps the two in lockstep, so a directory
        // that survives the open always has a session behind it.
        cache_directory.commit();
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
        self.read_range_with_mode(
            session_id,
            sheet_id,
            range,
            RangeReadMode::Wait,
            None,
            false,
        )
    }

    pub(crate) fn read_range_with_mode(
        &mut self,
        session_id: &str,
        sheet_id: &str,
        range: &CellRange,
        mode: RangeReadMode,
        sparse_rows: Option<&[usize]>,
        foreground: bool,
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
        range.validate(&session.sheets[sheet_index], sparse_rows)?;
        if session.progressive && foreground {
            session.scheduler.focus(sheet_index);
        }
        if !matches!(mode, RangeReadMode::Status) {
            session.ensure_parser(sheet_index)?;
            session.scheduler.ensure_metadata(sheet_index);
        }
        let result = session.read_range(sheet_index, range, mode, sparse_rows)?;
        if THREADS
            && session.progressive
            && (result.indexing_complete
                || (result.metadata_ready
                    && result
                        .indexed_through_row
                        .is_some_and(|row| row >= range.end_row)))
        {
            session.scheduler.viewport_ready(sheet_index);
        }
        Ok(result)
    }

    pub fn read_formula_cells(
        &mut self,
        session_id: &str,
        sheet_id: &str,
        background: bool,
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
        if !background {
            session.ensure_parser(sheet_index)?;
        }
        session.read_formula_cells(sheet_index)
    }

    pub fn read_row_outline(
        &mut self,
        session_id: &str,
        sheet_id: &str,
        background: bool,
    ) -> Result<RowOutlineResult, SidecarError> {
        let session = self
            .sessions
            .get_mut(session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let sheet_index = session
            .sheets
            .iter()
            .position(|sheet| sheet.id == sheet_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown worksheet.".into()))?;
        if !background {
            session.ensure_parser(sheet_index)?;
        }
        session.read_row_outline(sheet_index)
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
            .or_else(|| {
                // A picture inside a note is reachable the same way, under the
                // id its comment carries.
                session
                    .sheets
                    .iter()
                    .flat_map(|sheet| sheet.comments.iter())
                    .find_map(|comment| {
                        comment
                            .image
                            .as_ref()
                            .filter(|image| image.id == visual_id)
                            .map(|image| image.media_path.as_str())
                    })
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
    runtimes: Vec<IndexState>,
    filter_columns: filter::ColumnCache,
    visuals: Vec<VisualObject>,
    cache_directory: PathBuf,
    progressive: bool,
    scheduler: Arc<IndexScheduler>,
}

struct IndexSources {
    path: PathBuf,
    cache_directory: PathBuf,
    shared_strings: SharedStrings,
    styled_xfs: Arc<Vec<bool>>,
    color_context: ColorContext,
    cancelled: Arc<AtomicBool>,
    dependencies: Option<DependencyCatalog>,
}

impl IndexSources {
    fn worker(
        self: &Arc<Self>,
        sheet_index: usize,
        worksheet_path: &str,
        state: IndexState,
        sheet: &SheetMetadata,
    ) -> IndexWorker {
        let metadata: Option<MetadataJob> = sheet.progressive.then(|| {
            let source = Arc::clone(self);
            let state = Arc::clone(&state);
            let path = worksheet_path.to_owned();
            Box::new(move || {
                if let Err(error) = read_worksheet_metadata(
                    &source.path,
                    &path,
                    sheet_index,
                    &source.color_context,
                    &state,
                    &source.cancelled,
                ) && !source.cancelled.load(Ordering::Acquire)
                {
                    record_metadata_scan_failure(&state, sheet_index, &error);
                }
            }) as MetadataJob
        });
        let sources = Arc::clone(self);
        let index_state = Arc::clone(&state);
        let path = worksheet_path.to_owned();
        let progressive = sheet.progressive;
        // Only cells that received a rich-value picture lose their cached error.
        let rich_image_cells = sheet
            .cell_images
            .iter()
            .map(|image| (image.row, image.column))
            .collect();
        let index = Box::new(move |report: &mut dyn FnMut(usize)| {
            let mut dependencies = HashSet::new();
            let mut names = HashSet::new();
            let mut scan = |formula: &str| {
                if let Some(catalog) = &sources.dependencies {
                    catalog.scan(sheet_index, formula, &mut dependencies, &mut names, report);
                }
            };
            let result = index_worksheet(
                &sources.path,
                &path,
                sheet_index,
                &sources.cache_directory,
                &sources.shared_strings,
                &sources.styled_xfs,
                &sources.color_context,
                &rich_image_cells,
                progressive,
                &index_state,
                &sources.cancelled,
                &mut scan,
            );
            let (lock, condition) = &*index_state;
            if let Ok(mut index) = lock.lock() {
                match result {
                    Ok(()) => index.complete = true,
                    Err(error) => index.error = Some(error.to_string()),
                }
                condition.notify_all();
            }
        });
        IndexWorker::new(state, index, metadata)
    }
}

impl WorkbookSession {
    /// Media path behind a header/footer picture id (`&G` slots live on the
    /// indexed sheet state, so only completed sheets can resolve one).
    fn header_footer_picture_path(&self, picture_id: &str) -> Option<String> {
        self.runtimes.iter().find_map(|runtime| {
            let (lock, _) = &**runtime;
            let index = lock.lock().ok()?;
            index
                .decorations
                .as_ref()?
                .page_setup
                .as_ref()?
                .header_footer_pictures
                .iter()
                .find(|picture| picture.id == picture_id)
                .map(|picture| picture.media_path.clone())
        })
    }

    fn ensure_parser(&mut self, sheet_index: usize) -> Result<(), SidecarError> {
        self.scheduler.demand(sheet_index)
    }

    fn read_range(
        &self,
        sheet_index: usize,
        range: &CellRange,
        mode: RangeReadMode,
        sparse_rows: Option<&[usize]>,
    ) -> Result<RangeResult, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, condition) = &**runtime;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        let far_ahead = index.error.is_none()
            && !index.complete
            && index
                .indexed_through_row
                .map_or(range.start_row, |row| range.start_row.saturating_sub(row))
                >= RANGE_WAIT_MAX_LAG_ROWS;
        let index = if far_ahead || !matches!(mode, RangeReadMode::Wait) {
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
        let metadata_ready = index.decorations.is_some();
        let metadata_scan_failed = !metadata_ready && index.metadata_scan_failed;
        if matches!(mode, RangeReadMode::Status) {
            return Ok(RangeResult {
                indexed_through_row,
                indexing_complete,
                metadata_ready,
                metadata_scan_failed,
                ..RangeResult::default()
            });
        }
        let contains_row = |row| {
            sparse_rows.map_or(row >= range.start_row && row <= range.end_row, |rows| {
                rows.binary_search(&row).is_ok()
            })
        };
        let empty = SheetDecorations::default();
        let decorations = index.decorations.as_ref().unwrap_or(&empty);
        let merges = decorations
            .merges
            .iter()
            .filter(|merge| {
                merge.start_row <= range.end_row
                    && merge.end_row >= range.start_row
                    && merge.start_column <= range.end_column
                    && merge.end_column >= range.start_column
                    && sparse_rows.is_none_or(|rows| {
                        rows.get(rows.partition_point(|&row| row < merge.start_row))
                            .is_some_and(|&row| row <= merge.end_row)
                    })
            })
            .copied()
            .collect::<Vec<_>>();
        let hyperlinks = decorations
            .hyperlinks
            .iter()
            .filter(|link| {
                contains_row(link.row)
                    && link.column >= range.start_column
                    && link.column <= range.end_column
            })
            .cloned()
            .collect::<Vec<_>>();
        let conditional_rules = decorations.conditional_rules.clone();
        let auto_filter = decorations.auto_filter;
        let auto_filter_columns = decorations.auto_filter_columns.clone();
        let data_validations = decorations.data_validations.clone();
        let file_protection = decorations
            .protection
            .as_ref()
            .unwrap_or(&self.sheets[sheet_index].protection);
        let sheet_protection = metadata_ready
            .then(|| file_protection.sheet_protection.clone())
            .flatten();
        let protected_ranges = if metadata_ready {
            file_protection.protected_ranges.clone()
        } else {
            Vec::new()
        };
        let row_breaks = decorations.row_breaks.clone();
        let col_breaks = decorations.col_breaks.clone();
        let page_setup = decorations.page_setup.clone();
        let sparklines =
            (!decorations.sparklines.is_empty()).then(|| decorations.sparklines.clone());
        let cell_counts = index.cell_counts.filter(|_| indexing_complete);
        let first_chunk = range.start_row / CHUNK_ROW_COUNT;
        let last_available_row = indexed_through_row
            .map(|row| row.min(range.end_row))
            .unwrap_or(0);
        let last_chunk = last_available_row / CHUNK_ROW_COUNT;
        let paths = if indexed_through_row.is_none() || last_available_row < range.start_row {
            Vec::new()
        } else if let Some(rows) = sparse_rows {
            let mut previous = None;
            rows.iter()
                .take_while(|&&row| row <= last_available_row)
                .filter_map(|&row| {
                    let chunk = row / CHUNK_ROW_COUNT;
                    if previous.replace(chunk) == Some(chunk) {
                        return None;
                    }
                    index.chunk_files.get(&chunk).cloned()
                })
                .collect::<Vec<_>>()
        } else {
            (first_chunk..=last_chunk)
                .filter_map(|chunk| index.chunk_files.get(&chunk).cloned())
                .collect::<Vec<_>>()
        };
        drop(index);

        let mut cells = Vec::new();
        let mut rows = Vec::new();
        for path in paths {
            let chunk = chunks::read_chunk(&path, range.start_column..=range.end_column, true)?;
            cells.extend(chunk.cells.into_iter().filter(|cell| {
                contains_row(cell.row)
                    && cell.column >= range.start_column
                    && cell.column <= range.end_column
            }));
            rows.extend(chunk.rows.into_iter().filter(|row| contains_row(row.row)));
        }
        Ok(RangeResult {
            cells,
            rows,
            merges,
            hyperlinks,
            conditional_rules,
            auto_filter,
            auto_filter_columns,
            data_validations,
            sheet_protection,
            row_breaks,
            col_breaks,
            protected_ranges,
            sparklines,
            cell_counts,
            page_setup,
            indexed_through_row,
            indexing_complete,
            metadata_ready,
            metadata_scan_failed,
        })
    }

    /// All formula cells indexed so far; the closure analyzer polls until
    /// `indexing_complete`. Capped — a sheet over the cap sets `truncated`.
    fn read_formula_cells(&self, sheet_index: usize) -> Result<FormulaCellsResult, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, _) = &**runtime;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        if let Some(error) = &index.error {
            return Err(SidecarError::Workbook(error.clone()));
        }
        Ok(FormulaCellsResult {
            cells: index.formula_cells.clone(),
            // A truncated list is unusable anyway; the groups alone could
            // exceed the response cap.
            shared_groups: if index.formula_truncated {
                Vec::new()
            } else {
                index.shared_formula_groups.clone()
            },
            indexing_complete: index.complete,
            truncated: index.formula_truncated,
        })
    }

    fn read_row_outline(&self, sheet_index: usize) -> Result<RowOutlineResult, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, _) = &**runtime;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        if let Some(error) = &index.error {
            return Err(SidecarError::Workbook(error.clone()));
        }
        Ok(RowOutlineResult {
            rows: index.outline_rows.clone(),
            indexing_complete: index.complete,
        })
    }

    fn close(self) -> Result<(), SidecarError> {
        self.scheduler.close();
        // Drop the string table and join its writer before removing the cache.
        drop(self.scheduler);
        if self.cache_directory.exists() {
            fs::remove_dir_all(&self.cache_directory)?;
        }
        Ok(())
    }
}

fn record_metadata_scan_failure(
    state: &Arc<(Mutex<SheetIndex>, Condvar)>,
    sheet_index: usize,
    error: &SidecarError,
) {
    eprintln!("Early XLSX metadata scan failed for worksheet {sheet_index}: {error}");
    let (lock, condition) = &**state;
    if let Ok(mut index) = lock.lock() {
        index.metadata_scan_failed = true;
        condition.notify_all();
    }
}

#[derive(Default)]
struct SheetIndex {
    indexed_through_row: Option<usize>,
    complete: bool,
    error: Option<String>,
    chunk_files: HashMap<usize, PathBuf>,
    prefetch_bytes_remaining: Option<usize>,
    decorations: Option<SheetDecorations>,
    metadata_scan_failed: bool,
    /// Formula cells collected while indexing, capped at MAX_FORMULA_CELLS.
    /// Shared-formula followers are not counted: they live in
    /// `shared_formula_groups`, published once indexing completes.
    formula_cells: Vec<CellRecord>,
    formula_truncated: bool,
    shared_formula_groups: Vec<SharedFormulaGroup>,
    outline_rows: Vec<RowOutlineEntry>,
    cell_counts: Option<IndexedCellCounts>,
}

#[derive(Default)]
struct SheetDecorations {
    merges: Vec<MergedRange>,
    hyperlinks: Vec<HyperlinkRecord>,
    conditional_rules: Vec<ConditionalRule>,
    auto_filter: Option<MergedRange>,
    auto_filter_columns: Vec<FilterColumnCriteria>,
    data_validations: Vec<DataValidationRule>,
    row_breaks: Vec<usize>,
    col_breaks: Vec<usize>,
    page_setup: Option<PagePrintInfo>,
    protection: Option<SheetProtectionMetadata>,
    sparklines: Vec<SparklineGroupInfo>,
}

#[derive(Debug)]
struct SheetDeclaration {
    name: String,
    sheet_id: String,
    relationship_id: String,
    hidden: bool,
    very_hidden: bool,
}

//! Workbook-level parts: sheet declarations and relationships,
//! `<dimension>` / sheet-view scans, the shared-string table and workbook
//! defined names.

use super::*;

pub(crate) fn read_sheet_declarations(
    archive: &mut ZipArchive<File>,
) -> Result<
    (
        Vec<SheetDeclaration>,
        Option<WorkbookProtectionInfo>,
        usize,
        bool,
    ),
    SidecarError,
> {
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
                let has_password = attribute_value(&reader, &element, b"workbookPassword")?
                    .is_some()
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
                let relationship_id =
                    attribute_value(&reader, &element, b"id")?.ok_or_else(|| {
                        SidecarError::Workbook("Sheet has no relationship id.".into())
                    })?;
                let state = attribute_value(&reader, &element, b"state")?;
                let very_hidden = state.as_deref() == Some("veryHidden");
                let hidden = very_hidden || state.as_deref() == Some("hidden");
                sheets.push(SheetDeclaration {
                    name,
                    sheet_id,
                    relationship_id,
                    hidden,
                    very_hidden,
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((sheets, protection, active_tab, date_1904))
}

pub(crate) fn read_workbook_relationships(
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

pub(crate) fn normalize_worksheet_path(target: &str) -> Result<String, SidecarError> {
    let target = target.trim();
    let target = target
        .strip_prefix("xl/")
        .or_else(|| target.strip_prefix("xl\\"))
        .unwrap_or(target);
    crate::archive::resolve_relationship_target("xl/workbook.xml", target)
}

pub(crate) struct SheetDimensions {
    pub(crate) row_count: usize,
    pub(crate) column_count: usize,
    /// `<c>` elements carrying a value, formula or style; None when the
    /// workbook's counting budget ran out before this sheet was walked.
    pub(crate) stored_cell_count: Option<usize>,
    /// The subset of stored cells carrying a value or formula, and the
    /// (rows, columns) box those cells span; None with stored_cell_count.
    pub(crate) value_cell_count: Option<usize>,
    pub(crate) value_extent: Option<(usize, usize)>,
    pub(crate) column_widths: Vec<ColumnWidth>,
    pub(crate) default_row_height: Option<f64>,
    pub(crate) default_row_height_fixed: bool,
    pub(crate) default_column_width: Option<f64>,
    pub(crate) base_column_width: Option<f64>,
    pub(crate) freeze: Option<FreezePane>,
    pub(crate) tab_color: Option<String>,
    pub(crate) show_grid_lines: bool,
    pub(crate) show_formulas: bool,
    pub(crate) show_row_col_headers: bool,
    pub(crate) right_to_left: bool,
    pub(crate) zoom_scale: Option<u16>,
    pub(crate) outline_level_row: u8,
    pub(crate) outline_level_col: u8,
    pub(crate) outline_summary_below: bool,
    pub(crate) outline_summary_right: bool,
}

/// Excel keeps a zoom per view type; the app always opens the normal view,
/// so a sheet saved in page-break/layout view falls back to
/// zoomScaleNormal (absent = 100%). 100 maps to None so unzoomed sheets
/// serialize nothing.
pub(crate) fn normal_view_zoom(
    view: Option<&str>,
    zoom_scale: Option<u16>,
    zoom_scale_normal: Option<u16>,
) -> Option<u16> {
    let zoom = match view {
        Some("pageBreakPreview") | Some("pageLayout") => zoom_scale_normal,
        _ => zoom_scale.or(zoom_scale_normal),
    };
    zoom.map(|value| value.clamp(10, 400))
        .filter(|value| *value != 100)
}

pub(crate) fn read_sheet_dimensions(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    colors: &ColorContext,
    count_budget: &mut usize,
) -> Result<SheetDimensions, SidecarError> {
    let entry = zip_entry(archive, worksheet_path)?;
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
    let mut default_row_height_fixed = false;
    let mut default_column_width = None;
    let mut base_column_width = None;
    let mut freeze = None;
    let mut tab_color = None;
    let mut show_grid_lines = true;
    let mut show_formulas = false;
    let mut show_row_col_headers = true;
    let mut right_to_left = false;
    let mut zoom_scale = None;
    let mut dimension_trusted = false;
    let mut stored_cell_count = 0usize;
    let mut value_cell_count = 0usize;
    let mut value_extent = (0usize, 0usize);
    // Some(true) once the open <c> has shown a value/formula/style.
    let mut open_cell_stored: Option<bool> = None;
    let mut open_cell_position = (0usize, 0usize);
    let mut open_cell_has_value = false;
    let mut outline_level_row = 0u8;
    let mut outline_level_col = 0u8;
    let mut outline_summary_below = true;
    let mut outline_summary_right = true;
    let counted = loop {
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
                if element.local_name().as_ref() == b"outlinePr" =>
            {
                let flag = |name: &[u8]| -> Result<bool, SidecarError> {
                    Ok(!matches!(
                        attribute_value(&reader, &element, name)?.as_deref(),
                        Some("0") | Some("false")
                    ))
                };
                outline_summary_below = flag(b"summaryBelow")?;
                outline_summary_right = flag(b"summaryRight")?;
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
                if let Some(value) = attribute_value(&reader, &element, b"rightToLeft")? {
                    right_to_left = value == "1" || value == "true";
                }
                let percent = |name: &[u8]| -> Result<Option<u16>, SidecarError> {
                    Ok(attribute_value(&reader, &element, name)?
                        .and_then(|value| value.parse::<f64>().ok())
                        .filter(|value| *value > 0.0)
                        .map(|value| value.round() as u16))
                };
                zoom_scale = normal_view_zoom(
                    attribute_value(&reader, &element, b"view")?.as_deref(),
                    percent(b"zoomScale")?,
                    percent(b"zoomScaleNormal")?,
                );
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
                default_row_height_fixed = matches!(
                    attribute_value(&reader, &element, b"customHeight")?.as_deref(),
                    Some("1") | Some("true")
                );
                default_column_width = attribute_value(&reader, &element, b"defaultColWidth")?
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| *value > 0.0);
                base_column_width = attribute_value(&reader, &element, b"baseColWidth")?
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| *value > 0.0);
                let level = |name: &[u8]| -> Result<u8, SidecarError> {
                    Ok(attribute_value(&reader, &element, name)?
                        .and_then(|value| value.parse::<u8>().ok())
                        .map_or(0, |value| value.min(7)))
                };
                outline_level_row = level(b"outlineLevelRow")?;
                outline_level_col = level(b"outlineLevelCol")?;
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
                    // producers emitting A1 or A1:G1). Small dimensions can be
                    // stale too (tdf113271 declares A1:F5 over 462 rows) and
                    // are cheap to verify — trust only refs large enough that
                    // scanning them would cost real time. A ref reaching the
                    // sheet's last column or row (A1:XFD32, Yozo) declares the
                    // whole sheet, not the used range. A trusted ref still
                    // walks the cells, count-only, within the counting budget.
                    && dimensions.is_some_and(|(rows, columns)| {
                        rows > 1
                            && columns > 1
                            && !spans_full_axis(rows, columns)
                            && rows * columns >= DIMENSION_TRUST_CELLS
                    }) =>
            {
                dimension_trusted = true;
                if *count_budget == 0 {
                    let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                    break (row_count, column_count, false);
                }
            }
            Event::End(element)
                if element.local_name().as_ref() == b"sheetData" && dimension_trusted =>
            {
                let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                *count_budget -= stored_cell_count;
                break (row_count, column_count, true);
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
            Event::Start(element) if element.local_name().as_ref() == b"c" => {
                let (row, column) = match attribute_value(&reader, &element, b"r")? {
                    Some(address) => parse_address(&address).unwrap_or((current_row, next_column)),
                    None => (current_row, next_column),
                };
                next_column = column + 1;
                maximum_row = maximum_row.max(row);
                maximum_column = maximum_column.max(column);
                open_cell_stored = Some(attribute_value(&reader, &element, b"s")?.is_some());
                open_cell_position = (row, column);
                open_cell_has_value = false;
            }
            Event::Empty(element) if element.local_name().as_ref() == b"c" => {
                let (row, column) = match attribute_value(&reader, &element, b"r")? {
                    Some(address) => parse_address(&address).unwrap_or((current_row, next_column)),
                    None => (current_row, next_column),
                };
                next_column = column + 1;
                maximum_row = maximum_row.max(row);
                maximum_column = maximum_column.max(column);
                if attribute_value(&reader, &element, b"s")?.is_some() {
                    stored_cell_count += 1;
                }
                if dimension_trusted && stored_cell_count >= *count_budget {
                    let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                    *count_budget = 0;
                    break (row_count, column_count, false);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if open_cell_stored.is_some()
                    && matches!(element.local_name().as_ref(), b"v" | b"f" | b"is") =>
            {
                open_cell_stored = Some(true);
                open_cell_has_value = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"c" => {
                if open_cell_stored.take() == Some(true) {
                    stored_cell_count += 1;
                }
                if open_cell_has_value {
                    value_cell_count += 1;
                    value_extent.0 = value_extent.0.max(open_cell_position.0 + 1);
                    value_extent.1 = value_extent.1.max(open_cell_position.1 + 1);
                }
                if dimension_trusted && stored_cell_count >= *count_budget {
                    let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                    *count_budget = 0;
                    break (row_count, column_count, false);
                }
            }
            Event::Eof => {
                let (dim_rows, dim_columns) = dimensions.unwrap_or((0, 0));
                let (row_count, column_count) = (
                    declared_extent(dim_rows, SHEET_MAX_ROWS).max(maximum_row + 1),
                    declared_extent(dim_columns, SHEET_MAX_COLUMNS).max(maximum_column + 1),
                );
                break (row_count, column_count, true);
            }
            _ => {}
        }
        buffer.clear();
    };
    let (row_count, column_count, counted) = counted;
    Ok(SheetDimensions {
        row_count,
        column_count,
        stored_cell_count: counted.then_some(stored_cell_count),
        value_cell_count: counted.then_some(value_cell_count),
        value_extent: counted.then_some(value_extent),
        column_widths,
        default_row_height,
        default_row_height_fixed,
        default_column_width,
        base_column_width,
        freeze,
        tab_color,
        show_grid_lines,
        show_formulas,
        show_row_col_headers,
        right_to_left,
        zoom_scale,
        outline_level_row,
        outline_level_col,
        outline_summary_below,
        outline_summary_right,
    })
}

pub(crate) const DIMENSION_TRUST_CELLS: usize = 10_000;
/// Per-workbook budget for counting stored cells on trusted dimensions
/// (under half a microsecond per cell at open). Every host gate sits below it, so a
/// workbook that exhausts it is "too large" either way; sheets the budget
/// never reached report no count and the host falls back to their box.
pub(crate) const STORED_CELL_COUNT_BUDGET: usize = 2_000_000;
pub(crate) const SHEET_MAX_ROWS: usize = 1_048_576;
pub(crate) const SHEET_MAX_COLUMNS: usize = 16_384;

/// A declared extent reaching the sheet's last row/column means "whole
/// sheet", not a used range; only the measured cells count then.
fn declared_extent(declared: usize, sheet_max: usize) -> usize {
    if declared >= sheet_max { 0 } else { declared }
}

fn spans_full_axis(rows: usize, columns: usize) -> bool {
    rows >= SHEET_MAX_ROWS || columns >= SHEET_MAX_COLUMNS
}

pub(crate) fn dimensions_from_reference(reference: &str) -> Result<(usize, usize), SidecarError> {
    let last = reference
        .split(':')
        .next_back()
        .unwrap_or(reference)
        .replace('$', "");
    let (row, column) = parse_address(&last)?;
    Ok((row + 1, column + 1))
}

pub(crate) fn read_shared_strings(
    archive: &mut ZipArchive<File>,
    colors: &ColorContext,
) -> Result<Vec<SharedString>, SidecarError> {
    let Ok(entry) = zip_entry(archive, "xl/sharedStrings.xml") else {
        return Ok(Vec::new());
    };
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut strings = Vec::new();
    let mut current = String::new();
    let mut runs: Vec<RichRun> = Vec::new();
    let mut current_run: Option<RichRun> = None;
    let mut in_text = false;
    let mut text_preserve = false;
    let mut text_node = String::new();
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
                    text_preserve = preserves_space(&reader, &element)?;
                    text_node.clear();
                } else if let Some(run) = current_run.as_mut() {
                    apply_run_property(run, &reader, &element, colors)?;
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"t" => {
                in_text = !in_phonetic;
                text_preserve = preserves_space(&reader, &element)?;
                text_node.clear();
            }
            Event::Text(text) if in_text => text_node.push_str(&decode_text(&text)?),
            Event::CData(text) if in_text => text_node.push_str(&decode_cdata(&text)?),
            Event::GeneralRef(reference) if in_text => {
                text_node.push_str(&general_ref_text(&reference)?);
            }
            Event::End(element) if element.local_name().as_ref() == b"t" => {
                if in_text {
                    let text = text_node_content(std::mem::take(&mut text_node), text_preserve);
                    current.push_str(&text);
                    if let Some(run) = &mut current_run {
                        run.text.push_str(&text);
                    }
                }
                in_text = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"r" => {
                if let Some(run) = current_run.take() {
                    runs.push(run);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"si" => {
                let mut finished = std::mem::take(&mut runs);
                for run in &mut finished {
                    normalize_cell_text(&mut run.text);
                }
                let mut text = current.clone();
                normalize_cell_text(&mut text);
                strings.push(SharedString {
                    text,
                    runs: qualify_runs(finished),
                });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(strings)
}

pub(crate) fn has_run_formatting(run: &RichRun) -> bool {
    run.bold
        || run.italic
        || run.underline
        || run.strikethrough
        || run.color.is_some()
        || run.size.is_some()
        || run.family.is_some()
        || run.vert_align.is_some()
}

pub(crate) fn qualify_runs(runs: Vec<RichRun>) -> Option<Vec<RichRun>> {
    (runs.len() > 1 || runs.iter().any(has_run_formatting)).then_some(runs)
}

/// Applies one rPr child element (b/i/u/strike/sz/rFont/color/vertAlign) to a run.
pub(crate) fn apply_run_property<R: std::io::BufRead>(
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
            run.underline = attribute_value(reader, element, b"val")?.as_deref() != Some("none");
        }
        b"strike" => run.strikethrough = flag_on(attribute_value(reader, element, b"val")?),
        b"sz" => {
            run.size = attribute_value(reader, element, b"val")?
                .and_then(|value| value.parse::<f64>().ok());
        }
        b"rFont" => run.family = attribute_value(reader, element, b"val")?,
        b"vertAlign" => {
            run.vert_align = attribute_value(reader, element, b"val")?
                .filter(|value| value == "subscript" || value == "superscript");
        }
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

/// Sheet-scoped `_xlnm.Print_Area` / `_xlnm.Print_Titles` formulas, keyed by
/// localSheetId (workbook sheet order).
#[derive(Default)]
pub(crate) struct SheetPrintNames {
    pub(crate) area: Option<String>,
    pub(crate) titles: Option<String>,
}

pub(crate) fn read_defined_names(
    archive: &mut ZipArchive<File>,
) -> Result<
    (
        Vec<DefinedName>,
        HashMap<usize, SheetPrintNames>,
        HashSet<usize>,
    ),
    SidecarError,
> {
    let xml = read_zip_string(archive, "xl/workbook.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut names = Vec::new();
    let mut print_names: HashMap<usize, SheetPrintNames> = HashMap::new();
    // Sheets with ANY localSheetId-scoped definedName, including hidden and
    // _xlnm.* built-ins the editor never models — the save's duplicate guard
    // matches the raw attribute, so the host's early gate must too.
    let mut scoped_sheets: HashSet<usize> = HashSet::new();
    struct Pending {
        name: String,
        hidden: bool,
        sheet_index: Option<usize>,
        formula: String,
    }
    let mut current: Option<Pending> = None;
    loop {
        match reader.read_event()? {
            Event::Start(element) if element.local_name().as_ref() == b"definedName" => {
                let name = attribute_value(&reader, &element, b"name")?.unwrap_or_default();
                let hidden = attribute_value(&reader, &element, b"hidden")?
                    .is_some_and(|value| value == "1" || value == "true");
                let sheet_index = attribute_value(&reader, &element, b"localSheetId")?
                    .and_then(|value| value.parse::<usize>().ok());
                current = (!name.is_empty()).then_some(Pending {
                    name,
                    hidden,
                    sheet_index,
                    formula: String::new(),
                });
            }
            Event::Text(text) => {
                if let Some(defined) = &mut current {
                    defined.formula.push_str(&decode_text(&text)?);
                }
            }
            Event::CData(text) => {
                if let Some(defined) = &mut current {
                    defined.formula.push_str(&decode_cdata(&text)?);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some(defined) = &mut current {
                    defined.formula.push_str(&general_ref_text(&reference)?);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"definedName" => {
                if let Some(defined) = current.take() {
                    if let Some(sheet_index) = defined.sheet_index {
                        scoped_sheets.insert(sheet_index);
                    }
                    if defined.formula.is_empty() {
                        continue;
                    }
                    if defined.name.starts_with("_xlnm") {
                        // Print names feed the PDF export; other _xlnm.*
                        // built-ins stay file-only. Oversized refs exceed the
                        // wire cap and are dropped.
                        if defined.formula.len() > 2_000 {
                            continue;
                        }
                        if let Some(sheet_index) = defined.sheet_index {
                            let entry = print_names.entry(sheet_index).or_default();
                            if defined.name == "_xlnm.Print_Area" {
                                entry.area = Some(defined.formula);
                            } else if defined.name == "_xlnm.Print_Titles" {
                                entry.titles = Some(defined.formula);
                            }
                        }
                    // Hidden names stay file-only (the save preserves them
                    // verbatim; the editor never models them).
                    } else if !defined.hidden {
                        names.push(DefinedName {
                            name: defined.name,
                            formula: strip_future_function_markers(&defined.formula),
                            sheet_index: defined.sheet_index,
                        });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((names, print_names, scoped_sheets))
}

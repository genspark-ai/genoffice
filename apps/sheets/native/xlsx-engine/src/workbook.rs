//! Workbook-level parts: archive validation, sheet declarations and
//! relationships, `<dimension>` / sheet-view scans, the shared-string table
//! and workbook defined names.

use super::*;

pub(crate) fn validate_archive(archive: &mut ZipArchive<File>) -> Result<(), SidecarError> {
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

// Pure string handling: zip entry names always use '/', while PathBuf joins
// with '\' on Windows, which made by_name miss every worksheet there.
pub(crate) fn normalize_worksheet_path(target: &str) -> Result<String, SidecarError> {
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

pub(crate) struct SheetDimensions {
    pub(crate) row_count: usize,
    pub(crate) column_count: usize,
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
                    // scanning them would cost real time.
                    && dimensions.is_some_and(|(rows, columns)| {
                        rows > 1 && columns > 1 && rows * columns >= DIMENSION_TRUST_CELLS
                    }) =>
            {
                let (row_count, column_count) = dimensions.unwrap_or((1, 1));
                return Ok(SheetDimensions {
                    row_count,
                    column_count,
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
                });
            }
            _ => {}
        }
        buffer.clear();
    }
}

pub(crate) const DIMENSION_TRUST_CELLS: usize = 10_000;

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
            Event::CData(text) if in_text => {
                let decoded = decode_cdata(&text)?;
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
                let mut finished = std::mem::take(&mut runs);
                for run in &mut finished {
                    normalize_line_endings(&mut run.text);
                }
                let mut text = current.clone();
                normalize_line_endings(&mut text);
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

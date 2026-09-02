//! Streaming worksheet indexer: walks `sheetN.xml` once, emitting cell
//! chunks, row properties, merges, validations, conditional formats, page
//! setup and header/footer text into the session's `SheetIndex`.

use super::*;

#[allow(clippy::too_many_arguments)]
pub(crate) fn index_worksheet(
    workbook_path: &Path,
    worksheet_path: &str,
    sheet_index: usize,
    cache_directory: &Path,
    shared_strings: &[SharedString],
    styled_xfs: &[bool],
    colors: &ColorContext,
    rich_image_cells: &HashSet<(usize, usize)>,
    state: &Arc<(Mutex<SheetIndex>, Condvar)>,
    cancelled: &AtomicBool,
) -> Result<(), SidecarError> {
    let file = File::open(workbook_path)?;
    let mut archive = ZipArchive::new(file)?;
    let link_targets = visuals::hyperlink_targets(&mut archive, worksheet_path)?;
    let entry = zip_entry(&mut archive, worksheet_path)?;
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
    let mut x14_axis_colors: HashMap<String, String> = HashMap::new();
    let mut x14_databars: HashMap<String, X14DataBar> = HashMap::new();
    let mut x14_cfvo_kinds: HashMap<String, Vec<String>> = HashMap::new();
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
    let mut page_print = PagePrintInfo::default();
    // customSheetView carries its own print elements; only sheet-level ones
    // are what Excel prints.
    let mut in_custom_sheet_views = false;
    // Which <headerFooter> child the reader is inside (odd/even/first ×
    // header/footer), or None.
    let mut header_footer_section: Option<HeaderFooterSection> = None;
    let mut legacy_drawing_hf_id: Option<String> = None;

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
                    if let Some(cell) =
                        builder.finish(shared_strings, styled_xfs, rich_image_cells)?
                    {
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
                if let (Some(si), Some(builder)) =
                    (shared_formula_si(&reader, &element)?, cell_builder.as_mut())
                {
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
            event @ (Event::Text(_) | Event::CData(_)) => {
                let decoded = match &event {
                    Event::Text(text) => decode_text(text)?,
                    Event::CData(text) => decode_cdata(text)?,
                    _ => unreachable!(),
                };
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
                } else if let Some(section) = header_footer_section {
                    section
                        .slot(&mut page_print)
                        .get_or_insert_with(String::new)
                        .push_str(&decoded);
                } else if in_cf_formula {
                    if let Some(formula) =
                        cf_rule.as_mut().and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decoded);
                    }
                } else if in_cf_ext_id {
                    cf_rule_ext_id
                        .get_or_insert_with(String::new)
                        .push_str(&decoded);
                } else if in_dv_formula {
                    if let Some(formula) =
                        dv_rule.as_mut().and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decoded);
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
                } else if let Some(section) = header_footer_section {
                    section
                        .slot(&mut page_print)
                        .get_or_insert_with(String::new)
                        .push_str(&decoded);
                } else if in_cf_formula {
                    if let Some(formula) =
                        cf_rule.as_mut().and_then(|rule| rule.formulas.last_mut())
                    {
                        formula.push_str(&decoded);
                    }
                } else if in_dv_formula {
                    if let Some(formula) =
                        dv_rule.as_mut().and_then(|rule| rule.formulas.last_mut())
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
                    if let Some(cell) =
                        builder.finish(shared_strings, styled_xfs, rich_image_cells)?
                    {
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
            Event::Start(element) if element.local_name().as_ref() == b"customSheetViews" => {
                in_custom_sheet_views = true;
            }
            Event::End(element) if element.local_name().as_ref() == b"customSheetViews" => {
                in_custom_sheet_views = false;
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"pageSetUpPr" && !in_custom_sheet_views =>
            {
                page_print.fit_to_page = attribute_value(&reader, &element, b"fitToPage")?
                    .is_some_and(|value| value == "1" || value == "true");
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"printOptions" && !in_custom_sheet_views =>
            {
                page_print.print_gridlines = attribute_value(&reader, &element, b"gridLines")?
                    .is_some_and(|value| value == "1" || value == "true");
                page_print.print_headings = attribute_value(&reader, &element, b"headings")?
                    .is_some_and(|value| value == "1" || value == "true");
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"pageMargins" && !in_custom_sheet_views =>
            {
                if let (
                    Some(left),
                    Some(right),
                    Some(top),
                    Some(bottom),
                    Some(header),
                    Some(footer),
                ) = (
                    margin_attribute(&reader, &element, b"left")?,
                    margin_attribute(&reader, &element, b"right")?,
                    margin_attribute(&reader, &element, b"top")?,
                    margin_attribute(&reader, &element, b"bottom")?,
                    margin_attribute(&reader, &element, b"header")?,
                    margin_attribute(&reader, &element, b"footer")?,
                ) {
                    page_print.margins = Some(PageMarginsInfo {
                        left,
                        right,
                        top,
                        bottom,
                        header,
                        footer,
                    });
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"pageSetup" && !in_custom_sheet_views =>
            {
                page_print.orientation = attribute_value(&reader, &element, b"orientation")?
                    .filter(|value| value == "portrait" || value == "landscape");
                page_print.paper_size = attribute_value(&reader, &element, b"paperSize")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .filter(|value| (1..=256).contains(value));
                // Excel's valid scale range; anything else means "as saved by
                // a broken writer" and prints at 100.
                page_print.scale = attribute_value(&reader, &element, b"scale")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .filter(|value| (10..=400).contains(value));
                page_print.fit_to_width = attribute_value(&reader, &element, b"fitToWidth")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .filter(|value| *value <= 32_767);
                page_print.fit_to_height = attribute_value(&reader, &element, b"fitToHeight")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .filter(|value| *value <= 32_767);
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"headerFooter" && !in_custom_sheet_views =>
            {
                let flag = |name: &[u8]| -> Result<bool, SidecarError> {
                    Ok(attribute_value(&reader, &element, name)?
                        .is_some_and(|value| value == "1" || value == "true"))
                };
                page_print.different_odd_even = flag(b"differentOddEven")?;
                page_print.different_first = flag(b"differentFirst")?;
                page_print.header_footer_fixed_size =
                    attribute_value(&reader, &element, b"scaleWithDoc")?
                        .is_some_and(|value| value == "0" || value == "false");
            }
            Event::Start(element)
                if HeaderFooterSection::from_tag(element.local_name().as_ref()).is_some()
                    && !in_custom_sheet_views =>
            {
                header_footer_section =
                    HeaderFooterSection::from_tag(element.local_name().as_ref());
            }
            Event::End(element)
                if HeaderFooterSection::from_tag(element.local_name().as_ref()).is_some() =>
            {
                header_footer_section = None;
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"legacyDrawingHF" =>
            {
                legacy_drawing_hf_id = attribute_value(&reader, &element, b"id")?;
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
                if element.local_name().as_ref() == b"brk" && (in_row_breaks || in_col_breaks) =>
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
                    let has_password = attribute_value(&reader, &element, b"password")?.is_some()
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
            Event::Start(element) if element.local_name().as_ref() == b"dataValidation" => {
                dv_rule = parse_dv_rule(&reader, &element)?;
            }
            Event::Empty(element) if element.local_name().as_ref() == b"dataValidation" => {
                if let Some(rule) = parse_dv_rule(&reader, &element)? {
                    data_validations.push(rule);
                }
            }
            Event::Start(element)
                if element.local_name().as_ref().starts_with(b"formula") && dv_rule.is_some() =>
            {
                in_dv_formula = true;
                if let Some(rule) = &mut dv_rule {
                    rule.formulas.push(String::new());
                }
            }
            Event::End(element)
                if element.local_name().as_ref().starts_with(b"formula") && dv_rule.is_some() =>
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
                    // ECMA-376 defaults: the shortest bar spans 10% of the
                    // cell, the longest 90% (Excel 2007 look). An x14 twin
                    // overrides both below.
                    rule.min_length =
                        Some(bar_length_attribute(&reader, &element, b"minLength", 10)?);
                    rule.max_length =
                        Some(bar_length_attribute(&reader, &element, b"maxLength", 90)?);
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
                    // Only the three spec values travel; anything else reads
                    // as the default (automatic).
                    let axis_position = attribute_value(&reader, &element, b"axisPosition")?
                        .filter(|value| matches!(value.as_str(), "automatic" | "middle" | "none"));
                    x14_databars.insert(
                        id.clone(),
                        X14DataBar {
                            gradient,
                            same_as_positive,
                            axis_position,
                            // x14 defaults: 0/100 (the shortest bar is empty).
                            min_length: bar_length_attribute(&reader, &element, b"minLength", 0)?,
                            max_length: bar_length_attribute(&reader, &element, b"maxLength", 100)?,
                        },
                    );
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
                if element.local_name().as_ref() == b"axisColor" =>
            {
                if let Some(id) = &x14_current_id {
                    if let Some(color) = visuals::resolve_color(
                        attribute_value(&reader, &element, b"rgb")?.as_deref(),
                        attribute_value(&reader, &element, b"indexed")?.as_deref(),
                        attribute_value(&reader, &element, b"theme")?.as_deref(),
                        attribute_value(&reader, &element, b"tint")?.as_deref(),
                        colors,
                    ) {
                        x14_axis_colors.insert(id.clone(), color);
                    }
                }
            }
            // x14:cfvo inside an x14 twin: autoMin/autoMax refine the 2006
            // twin's plain min/max (Excel anchors auto bounds at zero).
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"cfvo" && cf_rule.is_none() =>
            {
                if let Some(id) = &x14_current_id {
                    if let Some(kind) = attribute_value(&reader, &element, b"type")? {
                        x14_cfvo_kinds.entry(id.clone()).or_default().push(kind);
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
    // The reader holds the worksheet entry (and with it the archive) until
    // dropped; the picture lookup needs the archive again.
    drop(reader);
    if let Some(relationship_id) = legacy_drawing_hf_id {
        page_print.header_footer_pictures = visuals::read_header_footer_pictures(
            &mut archive,
            worksheet_path,
            &relationship_id,
            sheet_index,
        )?;
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
    for (id, color) in x14_axis_colors {
        if let Some(rule) = x14_rule_slots
            .get(&id)
            .and_then(|slot| conditional_rules.get_mut(*slot))
        {
            rule.axis_color = Some(color);
        }
    }
    for (id, twin) in x14_databars {
        if let Some(rule) = x14_rule_slots
            .get(&id)
            .and_then(|slot| conditional_rules.get_mut(*slot))
        {
            rule.gradient = twin.gradient;
            rule.negative_same_as_positive = Some(twin.same_as_positive);
            rule.axis_position = twin.axis_position;
            rule.min_length = Some(twin.min_length);
            rule.max_length = Some(twin.max_length);
        }
    }
    for (id, kinds) in x14_cfvo_kinds {
        if let Some(rule) = x14_rule_slots
            .get(&id)
            .and_then(|slot| conditional_rules.get_mut(*slot))
        {
            for (cfvo, kind) in rule.cfvos.iter_mut().zip(kinds) {
                if kind == "autoMin" || kind == "autoMax" {
                    cfvo.kind = kind;
                }
            }
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
    for text in [
        &mut page_print.odd_header,
        &mut page_print.odd_footer,
        &mut page_print.even_header,
        &mut page_print.even_footer,
        &mut page_print.first_header,
        &mut page_print.first_footer,
    ] {
        if let Some(value) = text {
            truncate_utf16_units(value, 500);
        }
    }
    index.page_setup = (page_print != PagePrintInfo::default()).then_some(page_print);
    index.data_validations = data_validations;
    condition.notify_all();
    drop(index);
    Ok(())
}

pub(crate) fn parse_dv_rule<R: std::io::BufRead>(
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
        rule_type: attribute_value(reader, element, b"type")?.unwrap_or_else(|| "none".into()),
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

pub(crate) fn parse_cf_rule<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    ranges: &[MergedRange],
) -> Result<ConditionalRule, SidecarError> {
    Ok(ConditionalRule {
        ranges: ranges.to_vec(),
        rule_type: attribute_value(reader, element, b"type")?.unwrap_or_else(|| "unknown".into()),
        operator: attribute_value(reader, element, b"operator")?,
        formulas: Vec::new(),
        text: attribute_value(reader, element, b"text")?,
        dxf_index: attribute_value(reader, element, b"dxfId")?
            .and_then(|value| value.parse::<usize>().ok()),
        priority: attribute_value(reader, element, b"priority")?
            .and_then(|value| value.parse::<i64>().ok())
            .unwrap_or(0),
        stop_if_true: attribute_value(reader, element, b"stopIfTrue")?
            .is_some_and(|value| value == "1" || value == "true"),
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
        axis_position: None,
        axis_color: None,
        min_length: None,
        max_length: None,
    })
}

/// Attributes of an x14:dataBar twin, keyed by the x14:id that links it to
/// its 2006 base rule; merged once the sheet has been read.
pub(crate) struct X14DataBar {
    gradient: Option<bool>,
    same_as_positive: bool,
    axis_position: Option<String>,
    min_length: u32,
    max_length: u32,
}

/// dataBar/@minLength or @maxLength as a 0..=100 percentage; malformed or
/// absent values read as the element's schema default.
pub(crate) fn bar_length_attribute<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    name: &[u8],
    default: u32,
) -> Result<u32, SidecarError> {
    Ok(attribute_value(reader, element, name)?
        .and_then(|value| value.trim().parse::<u32>().ok())
        .map(|value| value.min(100))
        .unwrap_or(default))
}

/// The six text children of `<headerFooter>`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum HeaderFooterSection {
    OddHeader,
    OddFooter,
    EvenHeader,
    EvenFooter,
    FirstHeader,
    FirstFooter,
}

impl HeaderFooterSection {
    fn from_tag(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"oddHeader" => Some(Self::OddHeader),
            b"oddFooter" => Some(Self::OddFooter),
            b"evenHeader" => Some(Self::EvenHeader),
            b"evenFooter" => Some(Self::EvenFooter),
            b"firstHeader" => Some(Self::FirstHeader),
            b"firstFooter" => Some(Self::FirstFooter),
            _ => None,
        }
    }

    fn slot(self, page_print: &mut PagePrintInfo) -> &mut Option<String> {
        match self {
            Self::OddHeader => &mut page_print.odd_header,
            Self::OddFooter => &mut page_print.odd_footer,
            Self::EvenHeader => &mut page_print.even_header,
            Self::EvenFooter => &mut page_print.even_footer,
            Self::FirstHeader => &mut page_print.first_header,
            Self::FirstFooter => &mut page_print.first_footer,
        }
    }
}

/// Truncates in place to at most `max_units` UTF-16 code units — the wire
/// caps (zod / preload) measure JavaScript string length, where non-BMP
/// characters count as two.
pub(crate) fn truncate_utf16_units(value: &mut String, max_units: usize) {
    let mut units = 0usize;
    for (byte_index, character) in value.char_indices() {
        units += character.len_utf16();
        if units > max_units {
            value.truncate(byte_index);
            return;
        }
    }
}

pub(crate) fn margin_attribute<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<f64>, SidecarError> {
    Ok(attribute_value(reader, element, name)?
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite() && (0.0..=10.0).contains(value)))
}

pub(crate) fn row_property<R: std::io::BufRead>(
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
                attribute_value(reader, element, b"customFormat")
                    .ok()
                    .flatten()
                    .as_deref(),
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

pub(crate) fn flush_chunk(
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
        index
            .formula_cells
            .extend(pending_formulas.drain(..).take(room));
    }
    index.indexed_through_row = Some(indexed_through_row);
    condition.notify_all();
    Ok(())
}

pub(crate) struct CellBuilder {
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
        rich_image_cells: &HashSet<(usize, usize)>,
    ) -> Result<Option<CellRecord>, SidecarError> {
        let formula = if self.formula.is_empty() {
            None
        } else {
            Some(format!("={}", strip_future_function_markers(&self.formula)))
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
                let mut inline_runs = self.inline_runs;
                for run in &mut inline_runs {
                    normalize_line_endings(&mut run.text);
                }
                rich = qualify_runs(inline_runs);
                let mut inline_text = self.inline_text;
                normalize_line_endings(&mut inline_text);
                Some(CellValue::String(inline_text))
            }
            // An error cell hosting an in-cell picture record renders as the
            // picture, not as its cached #VALUE! placeholder.
            Some("e") if rich_image_cells.contains(&(self.row, self.column)) => None,
            Some("str") | Some("e") => {
                let mut raw_value = self.raw_value;
                normalize_line_endings(&mut raw_value);
                Some(CellValue::String(raw_value))
            }
            Some("b") => Some(CellValue::Boolean(self.raw_value == "1")),
            _ if self.raw_value.trim().is_empty() => None,
            // Real-world exporters write stray non-numeric values without a
            // type attribute (or pad them with whitespace — Excel still
            // reads those as numbers); degrade to text instead of failing
            // the sheet.
            _ => match self.raw_value.trim().parse::<f64>() {
                Ok(number) => Some(CellValue::Number(number)),
                Err(_) => {
                    let mut raw_value = self.raw_value;
                    normalize_line_endings(&mut raw_value);
                    Some(CellValue::String(raw_value))
                }
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
            array_ref: if formula.is_some() {
                self.array_ref
            } else {
                None
            },
            formula,
            style_index: self.style_index,
            rich,
        }))
    }
}

/// Excel stores post-2007 functions as `_xlfn.NAME` (worksheet-scope ones as
/// `_xlfn._xlws.NAME`); its UI never shows the markers. Strip them outside
/// string literals so the formula bar and engine see the plain names — the
/// save side restores the markers on edited cells.
pub(crate) fn strip_future_function_markers(formula: &str) -> String {
    if !formula.contains("_xlfn.") && !formula.contains("_xlws.") {
        return formula.to_owned();
    }
    formula
        .split('"')
        .enumerate()
        .map(|(index, segment)| {
            if index % 2 == 1 {
                segment.to_owned()
            } else {
                segment.replace("_xlfn.", "").replace("_xlws.", "")
            }
        })
        .collect::<Vec<_>>()
        .join("\"")
}

/// `ref` of a `<f t="array">` element; None for ordinary formulas.
pub(crate) fn array_formula_ref<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<Option<String>, SidecarError> {
    if attribute_value(reader, element, b"t")?.as_deref() != Some("array") {
        return Ok(None);
    }
    attribute_value(reader, element, b"ref")
}

/// `si` of a `<f t="shared">` element; None for ordinary formulas.
pub(crate) fn shared_formula_si<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<Option<u32>, SidecarError> {
    if attribute_value(reader, element, b"t")?.as_deref() != Some("shared") {
        return Ok(None);
    }
    Ok(attribute_value(reader, element, b"si")?.and_then(|value| value.parse::<u32>().ok()))
}

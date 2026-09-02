use super::*;

#[test]
fn normalizes_crlf_and_stray_cr_to_lf() {
    let mut crlf = "Hyderabad Day Shift\r\n(500 Seats)".to_owned();
    normalize_line_endings(&mut crlf);
    assert_eq!(crlf, "Hyderabad Day Shift\n(500 Seats)");

    let mut stray_cr = "a\rb\r\nc".to_owned();
    normalize_line_endings(&mut stray_cr);
    assert_eq!(stray_cr, "a\nb\nc");

    let mut untouched = "plain\ntext".to_owned();
    normalize_line_endings(&mut untouched);
    assert_eq!(untouched, "plain\ntext");
}

#[test]
fn strips_future_function_markers_outside_strings() {
    assert_eq!(
        strip_future_function_markers("_xlfn.MINIFS(C7:C10,C7:C10,\">0\")"),
        "MINIFS(C7:C10,C7:C10,\">0\")"
    );
    assert_eq!(
        strip_future_function_markers("_xlfn._xlws.SORT(A:A)"),
        "SORT(A:A)"
    );
    assert_eq!(
        strip_future_function_markers("CONCATENATE(\"_xlfn.MINIFS(\",A1)"),
        "CONCATENATE(\"_xlfn.MINIFS(\",A1)"
    );
    assert_eq!(strip_future_function_markers("SUM(A1:A3)"), "SUM(A1:A3)");
}

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
fn normal_view_zoom_follows_the_view_type() {
    // Normal view: zoomScale is the one to restore.
    assert_eq!(normal_view_zoom(None, Some(85), Some(85)), Some(85));
    assert_eq!(normal_view_zoom(Some("normal"), Some(85), None), Some(85));
    // Sheet saved in another view: only zoomScaleNormal applies.
    assert_eq!(
        normal_view_zoom(Some("pageBreakPreview"), Some(55), Some(70)),
        Some(70)
    );
    assert_eq!(
        normal_view_zoom(Some("pageBreakPreview"), Some(55), None),
        None
    );
    assert_eq!(
        normal_view_zoom(Some("pageLayout"), Some(55), Some(120)),
        Some(120)
    );
    // The 100% default serializes as nothing; out-of-range values clamp.
    assert_eq!(normal_view_zoom(None, Some(100), None), None);
    assert_eq!(normal_view_zoom(None, None, None), None);
    assert_eq!(normal_view_zoom(None, Some(500), None), Some(400));
    assert_eq!(normal_view_zoom(None, Some(5), None), Some(10));
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
        source_xml_bytes: 1024,
        column_widths: Vec::new(),
        default_row_height: None,
        default_row_height_fixed: false,
        default_column_width: None,
        base_column_width: None,
        freeze: None,
        hidden: false,
        tab_color: None,
        show_grid_lines: true,
        show_formulas: false,
        show_row_col_headers: true,
        right_to_left: false,
        zoom_scale: None,
        tables: Vec::new(),
        comments: Vec::new(),
        pivot_ranges: Vec::new(),
        pivot_tables: Vec::new(),
        sparklines: Vec::new(),
        print_area: None,
        print_titles: None,
        has_scoped_defined_names: false,
        cell_images: Vec::new(),
    };
    let range = CellRange {
        start_row: 0,
        end_row: 1_999,
        start_column: 0,
        end_column: 99,
    };
    assert!(range.validate(&sheet).is_err());
}

fn sparklines_from(xml: &str) -> Vec<SparklineGroupInfo> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    parse_sparkline_groups(&mut reader).unwrap()
}

#[test]
fn reads_cdata_wrapped_sparkline_refs() {
    let xml = sparkline_ext(
        r#"<x14:sparklineGroup>
<x14:sparklines><x14:sparkline>
<xm:f><![CDATA[Sheet1!A2:C2]]></xm:f><xm:sqref><![CDATA[D2]]></xm:sqref>
</x14:sparkline></x14:sparklines>
</x14:sparklineGroup>"#,
    );
    let groups = sparklines_from(&xml);
    assert_eq!(groups.len(), 1);
    assert_eq!(groups[0].cells[0].source_ref, "Sheet1!A2:C2");
    assert_eq!(groups[0].cells[0].cell, "D2");
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
    let cells: Vec<_> = groups[0]
        .cells
        .iter()
        .map(|entry| entry.cell.as_str())
        .collect();
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
        source_xml_bytes: 1024,
        column_widths: Vec::new(),
        default_row_height: None,
        default_row_height_fixed: false,
        default_column_width: None,
        base_column_width: None,
        freeze: None,
        hidden: false,
        tab_color: None,
        show_grid_lines: true,
        show_formulas: false,
        show_row_col_headers: true,
        right_to_left: false,
        zoom_scale: None,
        tables: Vec::new(),
        comments: Vec::new(),
        pivot_ranges: Vec::new(),
        pivot_tables: Vec::new(),
        sparklines: Vec::new(),
        print_area: None,
        print_titles: None,
        has_scoped_defined_names: false,
        cell_images: Vec::new(),
    };
    let json = serde_json::to_string(&sheet).unwrap();
    assert!(!json.contains("sparklines"));
}

fn rich_data_fixture_entries() -> Vec<(&'static str, &'static str)> {
    vec![
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/><Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/><Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="2"><c r="B2" t="e" vm="1"><v>#VALUE!</v></c><c r="C2" t="e"><v>#DIV/0!</v></c></row></sheetData>
</worksheet>"#,
        ),
        (
            "xl/metadata.xml",
            r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>"#,
        ),
        (
            "xl/richData/richValueRel.xml",
            r#"<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="rId1"/></richValueRels>"#,
        ),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#,
        ),
        (
            "xl/richData/rdrichvalue.xml",
            r#"<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>"#,
        ),
        (
            "xl/richData/rdrichvaluestructure.xml",
            r#"<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>"#,
        ),
        ("xl/media/image1.png", "fake-png-bytes"),
    ]
}

#[test]
fn resolves_in_cell_rich_value_pictures() {
    let (_dir, path) = open_fixture(&rich_data_fixture_entries());
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let images = &metadata.sheets[0].cell_images;
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].row, 1);
    assert_eq!(images[0].column, 1);
    assert_eq!(images[0].media_path, "xl/media/image1.png");

    // The media lookup accepts the cell-image id like a visual id.
    let media = sessions
        .read_media(&metadata.session_id, &images[0].id)
        .unwrap();
    assert_eq!(media.media_type, "image/png");

    // The cached #VALUE! stays out of the grid; unrelated error cells keep
    // their values.
    let range = CellRange {
        start_row: 0,
        end_row: 1,
        start_column: 0,
        end_column: 2,
    };
    let result = sessions
        .read_range(&metadata.session_id, "sheet-1", &range)
        .unwrap();
    assert!(
        !result
            .cells
            .iter()
            .any(|cell| cell.row == 1 && cell.column == 1 && cell.value.is_some())
    );
    assert!(result.cells.iter().any(|cell| {
        cell.row == 1
            && cell.column == 2
            && matches!(&cell.value, Some(CellValue::String(text)) if text == "#DIV/0!")
    }));

    // The wire payload carries only the lookup fields.
    let json = serde_json::to_string(&metadata.sheets[0]).unwrap();
    assert!(json.contains("cellImages"));
    assert!(!json.contains("mediaPath"));
}

/// A valueMetadata bk may carry one rc per metadata type; XLRICHVALUE
/// still resolves when it is not the first record.
#[test]
fn resolves_rich_value_from_a_non_first_metadata_record() {
    let metadata = r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="2"><metadataType name="XLDAPR" minSupportedVersion="120000"/><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="7"/><rc t="2" v="0"/></bk></valueMetadata></metadata>"#;
    let entries: Vec<(&str, &str)> = rich_data_fixture_entries()
        .into_iter()
        .map(|(name, content)| {
            if name == "xl/metadata.xml" {
                (name, metadata)
            } else {
                (name, content)
            }
        })
        .collect();
    let (_dir, path) = open_fixture(&entries);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let images = &metadata.sheets[0].cell_images;
    assert_eq!(images.len(), 1);
    assert_eq!((images[0].row, images[0].column), (1, 1));
    assert_eq!(images[0].media_path, "xl/media/image1.png");
}

/// Past the per-sheet overlay cap the cell keeps its cached error: a
/// picture cell must render as the image or the error, never blank.
#[test]
fn over_cap_picture_cells_keep_the_cached_error() {
    let cells: String = (0..=MAX_CELL_IMAGES)
        .map(|index| {
            format!(
                r#"<c r="{}1" t="e" vm="1"><v>#VALUE!</v></c>"#,
                column_label(index)
            )
        })
        .collect();
    let worksheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">{cells}</row></sheetData>
</worksheet>"#
    );
    let entries: Vec<(&str, &str)> = rich_data_fixture_entries()
        .into_iter()
        .map(|(name, content)| {
            if name == "xl/worksheets/sheet1.xml" {
                (name, worksheet.as_str())
            } else {
                (name, content)
            }
        })
        .collect();
    let (_dir, path) = open_fixture(&entries);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.sheets[0].cell_images.len(), MAX_CELL_IMAGES);
    let range = CellRange {
        start_row: 0,
        end_row: 0,
        start_column: 0,
        end_column: MAX_CELL_IMAGES,
    };
    let result = sessions
        .read_range(&metadata.session_id, "sheet-1", &range)
        .unwrap();
    // Within the cap: suppressed. The one past it: error preserved.
    assert!(
        !result
            .cells
            .iter()
            .any(|cell| cell.column < MAX_CELL_IMAGES && cell.value.is_some())
    );
    assert!(result.cells.iter().any(|cell| {
        cell.column == MAX_CELL_IMAGES
            && matches!(&cell.value, Some(CellValue::String(text)) if text == "#VALUE!")
    }));
}

fn column_label(mut index: usize) -> String {
    let mut label = String::new();
    loop {
        label.insert(0, (b'A' + (index % 26) as u8) as char);
        if index < 26 {
            return label;
        }
        index = index / 26 - 1;
    }
}

#[test]
fn vm_without_rich_data_parts_keeps_the_cached_error() {
    let entries = rich_data_fixture_entries();
    let kept: Vec<(&str, &str)> = entries
        .into_iter()
        .filter(|(name, _)| !name.contains("richData") && *name != "xl/metadata.xml")
        .collect();
    let (_dir, path) = open_fixture(&kept);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert!(metadata.sheets[0].cell_images.is_empty());
    let range = CellRange {
        start_row: 0,
        end_row: 1,
        start_column: 0,
        end_column: 2,
    };
    let result = sessions
        .read_range(&metadata.session_id, "sheet-1", &range)
        .unwrap();
    assert!(result.cells.iter().any(|cell| {
        cell.row == 1
            && cell.column == 1
            && matches!(&cell.value, Some(CellValue::String(text)) if text == "#VALUE!")
    }));
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

#[test]
fn parses_sheet_view_right_to_left() {
    let (_dir, path) = open_fixture(&[
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="R" sheetId="1" r:id="rId1"/><sheet name="L" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert!(metadata.sheets[0].right_to_left);
    assert!(!metadata.sheets[1].right_to_left);
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

/// Small multi-row dimensions can be stale too (tdf113271 declares
/// A1:F5 over 462 rows); they are cheap to verify by scanning.
#[test]
fn stale_small_dimension_is_remeasured() {
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
<dimension ref="A1:F5"/>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="40"><c r="C40"><v>2</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.sheets[0].row_count, 40);
    assert_eq!(metadata.sheets[0].column_count, 6);
}

/// Non-conformant packages (tdf131575, written by old .NET tooling) use
/// '\' entry separators, case-drifted names, and leading-'/' rel targets;
/// Excel opens them, so entry lookup must tolerate all three.
#[test]
fn opens_backslash_and_case_drifted_package() {
    let (_dir, path) = open_fixture(&[
        (
            r"xl\workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            r"xl\_rels\workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="/xl/sharedStrings.xml"/></Relationships>"#,
        ),
        (
            r"xl\sharedstrings.xml",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>hello</t></si></sst>"#,
        ),
        (
            r"xl\worksheets\sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>7</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.sheets.len(), 1);
    assert_eq!(metadata.sheets[0].row_count, 1);
    assert_eq!(metadata.sheets[0].column_count, 2);
}

/// Theme substitution keeps rendering on the theme latin face, but the
/// literal cached Normal-font name and the theme's minor <a:ea> face
/// still reach the wire — the renderer needs them for the column MDW.
#[test]
fn normal_font_name_survives_theme_substitution() {
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
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
        ),
        (
            "xl/theme/theme1.xml",
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Yu Gothic"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#,
        ),
        (
            "xl/styles.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><scheme val="minor"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.styles[0].font_family.as_deref(), Some("Calibri"));
    assert_eq!(
        metadata.normal_font_name.as_deref(),
        Some("ＭＳ Ｐゴシック")
    );
    let theme_fonts = metadata.theme_fonts.as_ref().unwrap();
    assert_eq!(theme_fonts.minor_ea.as_deref(), Some("Yu Gothic"));
    let json = serde_json::to_string(&metadata).unwrap();
    assert!(json.contains("\"normalFontName\""));
    assert!(json.contains("\"minorEa\""));
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
    assert!(
        matches!(&result.cells[0].value, Some(CellValue::String(text)) if text == "Partner ID")
    );
}

#[test]
fn reads_saved_print_settings_and_print_names() {
    let (_dir, path) = open_fixture(&[
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'S'!$A$1:$C$4</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">'S'!$1:$2</definedName><definedName name="Visible">S!$A$1</definedName></definedNames></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<customSheetViews><customSheetView guid="{1}"><pageSetup paperSize="1" orientation="landscape"/><headerFooter><oddHeader>decoy</oddHeader></headerFooter></customSheetView></customSheetViews>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<printOptions gridLines="1" headings="true"/>
<pageMargins left="0.25" right="0.3" top="0.75" bottom="0.8" header="0.3" footer="0.4"/>
<pageSetup paperSize="9" scale="65" fitToWidth="2" fitToHeight="0" orientation="landscape"/>
<headerFooter><oddHeader>&amp;CBudget &amp;A</oddHeader><oddFooter>&amp;CSeite &amp;P von &amp;N</oddFooter></headerFooter>
</worksheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(
        metadata.sheets[0].print_area.as_deref(),
        Some("'S'!$A$1:$C$4")
    );
    assert_eq!(
        metadata.sheets[0].print_titles.as_deref(),
        Some("'S'!$1:$2")
    );
    assert_eq!(metadata.defined_names.len(), 1);
    assert_eq!(metadata.defined_names[0].name, "Visible");
    // The _xlnm print names above are localSheetId-scoped even though the
    // modeled defined_names omit them — the duplicate gate keys off this.
    assert!(metadata.sheets[0].has_scoped_defined_names);
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
    let setup = result.page_setup.expect("page setup");
    assert_eq!(setup.orientation.as_deref(), Some("landscape"));
    assert_eq!(setup.paper_size, Some(9));
    assert_eq!(setup.scale, Some(65));
    assert_eq!(setup.fit_to_width, Some(2));
    assert_eq!(setup.fit_to_height, Some(0));
    assert!(setup.fit_to_page);
    assert!(setup.print_gridlines);
    assert!(setup.print_headings);
    let margins = setup.margins.expect("margins");
    assert_eq!(margins.left, 0.25);
    assert_eq!(margins.footer, 0.4);
    assert_eq!(setup.odd_header.as_deref(), Some("&CBudget &A"));
    assert_eq!(setup.odd_footer.as_deref(), Some("&CSeite &P von &N"));
}

/// differentOddEven/differentFirst plus all six section texts come
/// through, and `<legacyDrawingHF>` resolves its VML shapes to media
/// parts fetchable through read_media (a comment shape in the same VML
/// part and an oversized picture are ignored).
#[test]
fn reads_header_footer_variants_and_pictures() {
    let big_picture = "x".repeat(2 * 1024 * 1024 + 1);
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
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<headerFooter differentOddEven="1" differentFirst="true" scaleWithDoc="0"><oddHeader>&amp;L&amp;G&amp;COdd</oddHeader><oddFooter>&amp;P</oddFooter><evenHeader>&amp;CEven</evenHeader><evenFooter>&amp;R&amp;G</evenFooter><firstHeader>&amp;CFirst</firstHeader><firstFooter>&amp;L&amp;D</firstFooter></headerFooter>
<legacyDrawingHF r:id="rId7"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing2.vml"/></Relationships>"#,
        ),
        (
            "xl/drawings/vmlDrawing2.vml",
            r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" path="m@4@5l@4@11@9@11@9@5xe"/>
<v:shape id="LH" o:spid="_x0000_s3073" type="#_x0000_t75" style='position:absolute;margin-left:0;margin-top:0;width:442.5pt;height:43.5pt;z-index:1'><v:imagedata o:relid="rId1" o:title="logo"/></v:shape>
<v:shape id="RFEVEN" type="#_x0000_t75" style='width:1in;height:0.5in'><v:imagedata o:relid="rId2"/></v:shape>
<v:shape id="CH" type="#_x0000_t75" style='width:100pt;height:50pt'><v:imagedata o:relid="rId3"/></v:shape>
<v:shape id="Comment_x0020_1" type="#_x0000_t202" style='width:100pt;height:50pt'><x:ClientData ObjectType="Note"/></v:shape>
</xml>"##,
        ),
        (
            "xl/drawings/_rels/vmlDrawing2.vml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpeg"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.png"/></Relationships>"#,
        ),
        ("xl/media/image1.png", "PNG-BYTES"),
        ("xl/media/image2.jpeg", "JPEG-BYTES"),
        ("xl/media/image3.png", big_picture.as_str()),
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
    let setup = result.page_setup.expect("page setup");
    assert!(setup.different_odd_even);
    assert!(setup.different_first);
    assert!(setup.header_footer_fixed_size);
    assert_eq!(setup.odd_header.as_deref(), Some("&L&G&COdd"));
    assert_eq!(setup.odd_footer.as_deref(), Some("&P"));
    assert_eq!(setup.even_header.as_deref(), Some("&CEven"));
    assert_eq!(setup.even_footer.as_deref(), Some("&R&G"));
    assert_eq!(setup.first_header.as_deref(), Some("&CFirst"));
    assert_eq!(setup.first_footer.as_deref(), Some("&L&D"));
    let pictures = &setup.header_footer_pictures;
    assert_eq!(pictures.len(), 2, "{pictures:?}");
    assert_eq!(pictures[0].position, "LH");
    assert_eq!(pictures[0].width_pt, 442.5);
    assert_eq!(pictures[0].height_pt, 43.5);
    assert_eq!(pictures[0].media_type, "image/png");
    assert_eq!(pictures[1].position, "RFEVEN");
    assert_eq!(pictures[1].width_pt, 72.0);
    assert_eq!(pictures[1].height_pt, 36.0);
    assert_eq!(pictures[1].media_type, "image/jpeg");
    let media = sessions
        .read_media(&metadata.session_id, &pictures[0].id)
        .unwrap();
    assert_eq!(media.media_type, "image/png");
    assert_eq!(media.base64, {
        use base64::Engine;
        base64::engine::general_purpose::STANDARD.encode("PNG-BYTES")
    });
    let json = serde_json::to_string(&setup).unwrap();
    assert!(json.contains("\"differentOddEven\":true"));
    assert!(json.contains("\"headerFooterFixedSize\":true"));
    assert!(json.contains("\"headerFooterPictures\""));
    assert!(!json.contains("mediaPath"));
    assert!(
        sessions
            .read_media(&metadata.session_id, "hf-picture-0-ch")
            .is_err()
    );
}

/// The wire caps header/footer text at 500 UTF-16 code units (JS string
/// length); an emoji-laden header must come out within that bound or the
/// schema rejects the whole range read.
#[test]
fn caps_header_text_by_utf16_length() {
    let emoji_header = "😀".repeat(300);
    let worksheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<headerFooter><oddHeader>{emoji_header}</oddHeader></headerFooter>
</worksheet>"#
    );
    let (_dir, path) = open_fixture(&[
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        ("xl/worksheets/sheet1.xml", worksheet.as_str()),
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
    let header = result.page_setup.unwrap().odd_header.unwrap();
    assert_eq!(header.encode_utf16().count(), 500);
    assert!(header.chars().all(|character| character == '😀'));
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
    // Medium 1-7 stripe against a white body: no second band.
    assert!(table.second_row_stripe_fill.is_none());
}

/// Medium 8-14 fill both bands (Excel Medium9: accent tint 0.6
/// alternating with 0.8); dropping the second band left even data rows
/// white (50867_with_table).
#[test]
fn medium_full_color_styles_fill_both_row_bands() {
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
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="str"><v>a</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId2"/></tableParts></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
        ),
        (
            "xl/tables/table1.xml",
            r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C3"><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns><tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/></table>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let table = &metadata.sheets[0].tables[0];
    let base = DEFAULT_ACCENTS[0];
    assert_eq!(
        table.stripe_fill.as_deref(),
        Some(visuals::tint_to_hex(base, 0.6).as_str())
    );
    assert_eq!(
        table.second_row_stripe_fill.as_deref(),
        Some(visuals::tint_to_hex(base, 0.8).as_str())
    );
    assert_eq!(
        table.total_row_fill.as_deref(),
        Some(rgb_hex(base).as_str())
    );
    assert_eq!(table.total_row_font_color.as_deref(), Some("#FFFFFF"));
}

/// Negative tints must reproduce Excel's HSL shading: the totals band of
/// the Dark block and the Light-block text were pixel-measured on the
/// classic theme's accent1 (#4F81BD). Excel quantizes HSL to 0-240, so
/// allow one 1/255 step per channel.
#[test]
fn tint_matches_excel_shades() {
    fn assert_close(actual: &str, expected: &str) {
        let channel =
            |hex: &str, at: usize| i32::from_str_radix(&hex[at..at + 2], 16).expect("hex channel");
        for at in [1, 3, 5] {
            assert!(
                (channel(actual, at) - channel(expected, at)).abs() <= 1,
                "{actual} != {expected}"
            );
        }
    }
    let accent1 = (0x4F, 0x81, 0xBD);
    assert_close(&visuals::tint_to_hex(accent1, -0.25), "#366092");
    assert_close(&visuals::tint_to_hex(accent1, -0.5), "#244062");
    assert_close(&visuals::tint_to_hex(accent1, 0.6), "#B8CCE4");
    assert_close(&visuals::tint_to_hex(accent1, 0.8), "#DCE6F1");
    assert_close(&visuals::tint_to_hex((0, 0, 0), 0.85), "#D9D9D9");
}

/// Built-in family rules, per the Excel calibration workbook
/// (genoffice-sample/sheets/calib): totals bands, dk1-gray Medium 15-21
/// stripes, tinted Medium 22-28, solid Dark bodies, paired Dark 8-11.
#[test]
fn builtin_palette_matches_calibration() {
    let colors = ColorContext::default();
    let accent1 = DEFAULT_ACCENTS[0];

    let light2 = builtin_table_palette(Some("TableStyleLight2"), &colors);
    assert!(light2.header_fill.is_none());
    assert_eq!(
        light2.header_font_color,
        Some(visuals::tint_to_hex(accent1, -0.25))
    );
    assert_eq!(light2.stripe_fill, Some(visuals::tint_to_hex(accent1, 0.8)));
    assert_eq!(light2.total_row_border_style.as_deref(), Some("thin"));
    assert_eq!(light2.total_row_border_color, Some(rgb_hex(accent1)));
    assert!(light2.total_row_fill.is_none());

    // Light 15-21 headers are plain dk1, not accent-colored, and the
    // whole table carries a thin accent grid (outline + inner rules).
    let light16 = builtin_table_palette(Some("TableStyleLight16"), &colors);
    assert_eq!(light16.header_font_color.as_deref(), Some("#000000"));
    assert_eq!(light16.total_row_border_style.as_deref(), Some("double"));
    assert_eq!(light16.whole_table_border_color, Some(rgb_hex(accent1)));
    assert_eq!(light16.whole_table_border_style.as_deref(), Some("thin"));
    assert_eq!(
        light16.inner_horizontal_border_color,
        Some(rgb_hex(accent1))
    );
    assert_eq!(light16.inner_vertical_border_color, Some(rgb_hex(accent1)));

    let medium2 = builtin_table_palette(Some("TableStyleMedium2"), &colors);
    assert_eq!(medium2.header_fill, Some(rgb_hex(accent1)));
    assert!(medium2.total_row_fill.is_none());
    assert_eq!(medium2.total_row_border_style.as_deref(), Some("double"));
    assert_eq!(medium2.total_row_border_color, Some(rgb_hex(accent1)));

    // Medium 15-21 stripe in dk1 gray and rule the totals band in dk1.
    let medium16 = builtin_table_palette(Some("TableStyleMedium16"), &colors);
    assert_eq!(medium16.header_fill, Some(rgb_hex(accent1)));
    assert_eq!(medium16.stripe_fill.as_deref(), Some("#D9D9D9"));
    assert_eq!(medium16.total_row_border_color.as_deref(), Some("#000000"));

    let medium23 = builtin_table_palette(Some("TableStyleMedium23"), &colors);
    assert_eq!(
        medium23.header_fill,
        Some(visuals::tint_to_hex(accent1, 0.8))
    );
    assert_eq!(medium23.header_font_color.as_deref(), Some("#000000"));
    assert_eq!(
        medium23.whole_table_fill,
        Some(visuals::tint_to_hex(accent1, 0.8))
    );
    assert_eq!(
        medium23.total_row_fill,
        Some(visuals::tint_to_hex(accent1, 0.8))
    );
    assert_eq!(medium23.total_row_border_style.as_deref(), Some("medium"));

    let dark2 = builtin_table_palette(Some("TableStyleDark2"), &colors);
    assert_eq!(dark2.header_fill.as_deref(), Some("#000000"));
    assert_eq!(dark2.whole_table_fill, Some(rgb_hex(accent1)));
    assert_eq!(
        dark2.stripe_fill,
        Some(visuals::tint_to_hex(accent1, -0.25))
    );
    assert_eq!(
        dark2.total_row_fill,
        Some(visuals::tint_to_hex(accent1, -0.5))
    );
    assert_eq!(dark2.body_font_color.as_deref(), Some("#FFFFFF"));
    assert_eq!(dark2.total_row_font_color.as_deref(), Some("#FFFFFF"));

    // Dark 10 pairs an accent4 header with accent3 bands.
    let dark10 = builtin_table_palette(Some("TableStyleDark10"), &colors);
    assert_eq!(dark10.header_fill, Some(rgb_hex(DEFAULT_ACCENTS[3])));
    assert_eq!(
        dark10.stripe_fill,
        Some(visuals::tint_to_hex(DEFAULT_ACCENTS[2], 0.6))
    );
    assert_eq!(dark10.total_row_border_color.as_deref(), Some("#000000"));

    // dk1 members tint 0.05 lighter than the accent members.
    let light1 = builtin_table_palette(Some("TableStyleLight1"), &colors);
    assert_eq!(light1.stripe_fill.as_deref(), Some("#D9D9D9"));

    // Out-of-range family numbers must not panic on accent lookup.
    let _ = builtin_table_palette(Some("TableStyleDark12"), &colors);
    assert!(builtin_table_palette(None, &colors).header_fill.is_none());
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

/// rPr <vertAlign val="subscript|superscript"> must survive into the
/// wire; other values are dropped, and vertAlign alone qualifies a run
/// as formatted.
#[test]
fn rich_run_vert_align_reaches_the_wire() {
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
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><t>Cr</t></r><r><rPr><vertAlign val="subscript"/></rPr><t>2</t></r><r><rPr><vertAlign val="superscript"/></rPr><t>3</t></r><r><rPr><vertAlign val="baseline"/></rPr><t>%</t></r></si></sst>"#,
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
    let runs = result.cells[0].rich.as_ref().expect("rich runs");
    assert_eq!(runs[0].vert_align, None);
    assert_eq!(runs[1].vert_align.as_deref(), Some("subscript"));
    assert_eq!(runs[2].vert_align.as_deref(), Some("superscript"));
    assert_eq!(runs[3].vert_align, None);
    let json = serde_json::to_value(&runs[1]).unwrap();
    assert_eq!(json["vertAlign"], "subscript");
}

/// A oneCellAnchor sizes by xdr:ext (encoded as offsets in the from
/// cell); dashes, caps, and flips ride along the shape metadata.
#[test]
fn parses_one_cell_anchor_ext_and_line_style() {
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
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>342900</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>133349</xdr:rowOff></xdr:from><xdr:ext cx="1647825" cy="790576"/><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Line 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="127000" cap="rnd"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let shape = &metadata.visuals[0];
    assert_eq!(shape.kind, "shape");
    assert_eq!(shape.anchor.to_row, 8);
    assert_eq!(shape.anchor.to_column, 6);
    assert_eq!(shape.anchor.to_row_offset, 133349 + 790576);
    assert_eq!(shape.anchor.to_column_offset, 342900 + 1647825);
    assert_eq!(shape.line_dash.as_deref(), Some("sysDot"));
    assert_eq!(shape.line_cap.as_deref(), Some("rnd"));
    assert!(shape.flip_v);
    assert!(!shape.flip_h);
    assert_eq!(shape.line_width, Some(10.0));
}

/// A workbook whose first sheet is a chartsheet keeps that sheet in the
/// tab list (name and order preserved) and its absoluteAnchor-ed chart
/// flows through the visuals pipeline like a worksheet drawing.
#[test]
fn chartsheet_keeps_its_tab_and_renders_its_chart() {
    let (_dir, path) = open_fixture(&[
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart1" sheetId="9" r:id="rId1"/><sheet name="Sheet1" sheetId="1" r:id="rId2"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/chartsheets/sheet1.xml",
            r#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr/><sheetViews><sheetView tabSelected="1" workbookViewId="0" zoomToFit="1"/></sheetViews><drawing r:id="rId1"/></chartsheet>"#,
        ),
        (
            "xl/chartsheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9304587" cy="6077107"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#,
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
        ),
        (
            "xl/charts/chart1.xml",
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Bears</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>8</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>8</v></c></row></sheetData></worksheet>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let names: Vec<&str> = metadata
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect();
    assert_eq!(names, ["Chart1", "Sheet1"]);
    assert_eq!(metadata.sheets[0].id, "sheet-9");
    let chart = metadata
        .visuals
        .iter()
        .find(|visual| visual.sheet_id == "sheet-9")
        .expect("chartsheet chart visual");
    assert_eq!(chart.kind, "chart");
    assert_eq!(chart.chart_path.as_deref(), Some("xl/charts/chart1.xml"));
    let chart_types = &chart.chart.as_ref().expect("parsed chart").chart_types;
    assert_eq!(chart_types, &["barChart"]);
    // absoluteAnchor: top-left cell markers with the xdr:ext as offsets.
    assert_eq!(chart.anchor.from_row, 0);
    assert_eq!(chart.anchor.from_column, 0);
    assert_eq!(chart.anchor.from_row_offset, 0);
    assert_eq!(chart.anchor.from_column_offset, 0);
    assert_eq!(chart.anchor.to_row, 0);
    assert_eq!(chart.anchor.to_column, 0);
    assert_eq!(chart.anchor.to_row_offset, 6077107);
    assert_eq!(chart.anchor.to_column_offset, 9304587);
}

/// A bare `<oleObject>` with neither objectPr nor a VML shape still
/// surfaces as an OLE visual: the hidden drawing fallback lends its
/// anchor and is itself dropped so the object is not drawn twice.
#[test]
fn ole_object_without_anchor_borrows_the_drawing_fallback() {
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
    assert!(metadata.visuals.iter().all(|visual| visual.kind != "shape"));
    let ole = metadata
        .visuals
        .iter()
        .find(|visual| visual.kind == "ole")
        .expect("ole visual");
    assert_eq!(ole.prog_id.as_deref(), Some("Acrobat Document"));
    assert!(ole.media_path.is_none());
    assert_eq!(
        (
            ole.anchor.from_row,
            ole.anchor.from_column,
            ole.anchor.to_row,
            ole.anchor.to_column
        ),
        (1, 1, 4, 3)
    );
    // Read-only: no drawing edit locator.
    assert!(ole.drawing_path.is_none() && ole.drawing_index.is_none());
}

/// The OLE visual takes its fallback shape's place in the drawing so a
/// shape drawn after the object in Excel still paints on top of it.
#[test]
fn ole_object_keeps_the_fallback_shapes_z_order_slot() {
    let shape = |id: u32, name: &str, hidden: &str, row: u32| {
        format!(
            r#"<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="{name}"{hidden}/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"#,
            row + 3
        )
    };
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{}{}{}</xdr:wsDr>"#,
        shape(2, "Below", "", 0),
        shape(1025, "Object 1", r#" hidden="1""#, 1),
        shape(3, "Above", "", 2),
    );
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
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId2"/><oleObjects><oleObject progId="Acrobat Document" shapeId="1025" r:id="rId3"/></oleObjects></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
        ),
        ("xl/drawings/drawing1.xml", drawing.as_str()),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let order: Vec<(&str, Option<&str>)> = metadata
        .visuals
        .iter()
        .map(|visual| (visual.kind.as_str(), visual.name.as_deref()))
        .collect();
    assert_eq!(
        order,
        vec![
            ("shape", Some("Below")),
            ("ole", None),
            ("shape", Some("Above"))
        ]
    );
    // The shapes keep the ids read_drawing gave them by anchor position.
    assert_eq!(metadata.visuals[0].id, "visual-1");
    assert_eq!(metadata.visuals[1].id, "ole-1");
    assert_eq!(metadata.visuals[2].id, "visual-3");
}

/// Excel 2010+ form: `mc:AlternateContent` repeats the object (x14 Choice
/// with objectPr, bare Fallback). The visual takes the objectPr anchor,
/// the preview picture behind `objectPr/@r:id`, and the hidden compat
/// fallback shape in the drawing part disappears.
#[test]
fn ole_object_with_object_pr_uses_its_anchor_and_preview() {
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
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/><drawing r:id="rId2"/><legacyDrawing r:id="rId3"/><oleObjects><mc:AlternateContent><mc:Choice Requires="x14"><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId4"><objectPr defaultSize="0" r:id="rId5"><anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>4</xdr:col><xdr:colOff>384810</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>156210</xdr:rowOff></to></anchor></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId4"/></mc:Fallback></mc:AlternateContent></oleObjects></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/></Relationships>"#,
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>384810</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>156210</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1025" name="Object 1" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
        ),
        (
            "xl/drawings/vmlDrawing1.vml",
            r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" filled="f" stroked="f"/><v:shape id="_x0000_s1025" type="#_x0000_t75" style="position:absolute" filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]"><v:imagedata o:relid="rId1" o:title=""/><x:ClientData ObjectType="Pict"><x:Anchor>1, 0, 1, 0, 4, 101, 3, 41</x:Anchor></x:ClientData></v:shape></xml>"##,
        ),
        (
            "xl/drawings/_rels/vmlDrawing1.vml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.emf"/></Relationships>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.visuals.len(), 1, "fallback shape must be dropped");
    let ole = &metadata.visuals[0];
    assert_eq!(ole.kind, "ole");
    assert_eq!(ole.prog_id.as_deref(), Some("Word.Document.12"));
    // objectPr wins over the VML picture for the preview.
    assert_eq!(ole.media_path.as_deref(), Some("xl/media/image1.emf"));
    assert_eq!(ole.media_type.as_deref(), Some("image/x-emf"));
    assert_eq!(ole.anchor.from_column, 1);
    assert_eq!(ole.anchor.to_column, 4);
    assert_eq!(ole.anchor.to_column_offset, 384_810);
    assert_eq!(ole.anchor.to_row_offset, 156_210);
    assert!(ole.anchor.explicit_to);
    // VML stroked/filled with system colours → Excel's black frame on white.
    assert_eq!(ole.line_color.as_deref(), Some("#000000"));
    assert_eq!(ole.fill_color.as_deref(), Some("#FFFFFF"));
}

/// Pre-2010 bare form with no drawing part at all: the legacy VML shape
/// (`_x0000_s<shapeId>`) supplies the anchor in pixel offsets and the
/// preview picture through the VML part's own relationships.
#[test]
fn bare_ole_object_falls_back_to_the_vml_shape() {
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
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rId2"/><oleObjects><oleObject progId="Worksheet" shapeId="1025" r:id="rId3"/><oleObject progId="Package" shapeId="1026" r:id="rId4"/></oleObjects></worksheet>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/Microsoft_Office_Excel_97-2003_Worksheet1.xls"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/></Relationships>"#,
        ),
        (
            "xl/drawings/vmlDrawing1.vml",
            r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" filled="f" stroked="f"/><v:shape id="_x0000_s1025" type="#_x0000_t75" style="position:absolute"><v:imagedata o:relid="rId1" o:title=""/><x:ClientData ObjectType="Pict"><x:SizeWithCells/><x:Anchor>
    0, 0, 4, 0, 5, 1, 10, 1</x:Anchor><x:CF>Pict</x:CF></x:ClientData></v:shape><v:shape id="_x0000_s1026" type="#_x0000_t75" style="position:absolute" stroked="t" strokecolor="#1F4E79"><v:imagedata o:relid="rId2" o:title=""/><x:ClientData ObjectType="Pict"><x:Anchor>0, 0, 11, 0, 0, 42, 13, 3</x:Anchor></x:ClientData></v:shape></xml>"##,
        ),
        (
            "xl/drawings/_rels/vmlDrawing1.vml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.emf"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/></Relationships>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let ole: Vec<_> = metadata
        .visuals
        .iter()
        .filter(|visual| visual.kind == "ole")
        .collect();
    assert_eq!(ole.len(), 2);
    assert_eq!(ole[0].prog_id.as_deref(), Some("Worksheet"));
    assert_eq!(ole[0].media_path.as_deref(), Some("xl/media/image1.emf"));
    // "0, 0, 4, 0, 5, 1, 10, 1": columns 0..5, rows 4..10, px offsets → EMU.
    assert_eq!(
        (
            ole[0].anchor.from_column,
            ole[0].anchor.from_row,
            ole[0].anchor.to_column,
            ole[0].anchor.to_row
        ),
        (0, 4, 5, 10)
    );
    assert_eq!(ole[0].anchor.to_column_offset, 9525);
    assert_eq!(ole[0].anchor.to_row_offset, 9525);
    assert_eq!(ole[1].prog_id.as_deref(), Some("Package"));
    assert_eq!(ole[1].media_path.as_deref(), Some("xl/media/image2.emf"));
    assert_eq!(ole[1].anchor.to_column_offset, 42 * 9525);
    assert_eq!(ole[1].anchor.to_row_offset, 3 * 9525);
    assert_ne!(ole[0].id, ole[1].id);
    // Frame/fill inherit from the shapetype (both off) unless the shape
    // overrides; an explicit hex stroke colour passes through.
    assert_eq!(ole[0].line_color, None);
    assert_eq!(ole[0].fill_color, None);
    assert_eq!(ole[1].line_color.as_deref(), Some("#1F4E79"));
    assert_eq!(ole[1].fill_color, None);
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
    assert!(pivot.palette.header_fill.is_some());
    assert!(pivot.styled);
}

/// Light 1-7 with stripes off resolve no fills but stay styled, so the
/// renderer keeps the header/grand-total bands bold.
#[test]
fn pivot_light_low_stripes_off_stays_styled() {
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
            r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1" rowGrandTotals="1"><location ref="A8:B11" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotTableStyleInfo name="PivotStyleLight2" showRowHeaders="1" showRowStripes="0"/></pivotTableDefinition>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let pivot = &metadata.sheets[0].pivot_tables[0];
    assert!(pivot.styled);
    assert!(pivot.palette.header_fill.is_none());
    assert!(pivot.palette.stripe_fill.is_none());
    assert!(pivot.palette.whole_table_fill.is_none());
    let json = serde_json::to_string(pivot).unwrap();
    assert!(json.contains("\"styled\":true"));
}

/// Excel-calibrated pivot palettes (genoffice-sample/sheets/calib/
/// pivot-style-truths.json, Office 2007 theme): one representative per
/// block of seven, exact RGB.
#[test]
fn pivot_style_palette_calibrated_bands() {
    let colors = calib_theme();
    let palette = |name: &str| pivot_style_palette(Some(name), &colors);
    let white = Some("#FFFFFF");

    // Light 1-7: no fills, accent-1 tint 0.8 first stripe, white total.
    let light2 = palette("PivotStyleLight2");
    assert!(light2.header_fill.is_none());
    assert!(light2.header_font_color.is_none());
    assert_eq!(light2.stripe_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(light2.column_stripe_fill.as_deref(), Some("#DCE6F2"));
    assert!(light2.second_row_stripe_fill.is_none());
    assert_eq!(light2.total_row_fill.as_deref(), white);
    assert_eq!(light2.subheading2_font_color.as_deref(), Some("#4F81BD"));
    assert_eq!(light2.subheading2_bold, Some(false));
    assert_eq!(palette("PivotStyleLight3").subheading2_bold, Some(true));

    // Light 8-14: shaded body text, tint-0.8 subheading / subtotal.
    let light9 = palette("PivotStyleLight9");
    assert!(light9.header_fill.is_none());
    assert_eq!(light9.whole_table_font_color.as_deref(), Some("#376092"));
    assert_eq!(light9.header_font_color.as_deref(), Some("#000000"));
    assert_eq!(light9.subheading_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(light9.subtotal_fill.as_deref(), Some("#DCE6F2"));
    assert!(light9.stripe_fill.is_none());

    // Light 15-21: tinted header / total, dk1-gray stripes.
    let light16 = palette("PivotStyleLight16");
    assert_eq!(light16.header_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(light16.total_row_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(light16.stripe_fill.as_deref(), Some("#D9D9D9"));
    assert_eq!(light16.header_bold, Some(true));

    // Medium 1-7: shaded header (white, regular), tint-0.4 subheading
    // and subtotal, bold first header cell.
    let medium2 = palette("PivotStyleMedium2");
    assert_eq!(medium2.header_fill.as_deref(), Some("#376092"));
    assert_eq!(medium2.header_font_color.as_deref(), white);
    assert_eq!(medium2.header_bold, Some(false));
    assert_eq!(medium2.first_header_cell_bold, Some(true));
    assert_eq!(medium2.subheading_fill.as_deref(), Some("#95B3D7"));
    assert_eq!(medium2.subheading2_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(medium2.subtotal_fill.as_deref(), Some("#95B3D7"));
    assert_eq!(medium2.subtotal_font_color.as_deref(), white);
    assert!(medium2.whole_table_fill.is_none());
    assert!(medium2.total_row_fill.is_none());

    // Medium 8-14: solid accent header with white bold text.
    let medium9 = palette("PivotStyleMedium9");
    assert_eq!(medium9.header_fill.as_deref(), Some("#4F81BD"));
    assert_eq!(medium9.header_font_color.as_deref(), white);
    assert_eq!(medium9.header_bold, Some(true));
    assert_eq!(medium9.subheading_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(medium9.subtotal_fill.as_deref(), Some("#B9CDE5"));
    assert!(medium9.stripe_fill.is_none());
    assert!(medium9.whole_table_fill.is_none());
    assert!(medium9.total_row_fill.is_none());

    // Medium 15-21: dk1 header and total over a tint-0.8 body.
    let medium16 = palette("PivotStyleMedium16");
    assert_eq!(medium16.header_fill.as_deref(), Some("#000000"));
    assert_eq!(medium16.header_font_color.as_deref(), white);
    assert_eq!(medium16.header_bold, Some(false));
    assert_eq!(medium16.whole_table_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(medium16.stripe_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(medium16.total_row_fill.as_deref(), Some("#000000"));
    assert_eq!(medium16.subheading2_font_color.as_deref(), Some("#808080"));
    // Neutral member: body 0.95, stripe 0.85.
    let medium15 = palette("PivotStyleMedium15");
    assert_eq!(medium15.whole_table_fill.as_deref(), Some("#F2F2F2"));
    assert_eq!(medium15.stripe_fill.as_deref(), Some("#D9D9D9"));

    // Medium 22-28: unfilled header over a tint-0.8 body, tint-0.6
    // row-label column and second stripes, shaded text.
    let medium23 = palette("PivotStyleMedium23");
    assert!(medium23.header_fill.is_none());
    assert_eq!(medium23.whole_table_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(medium23.whole_table_font_color.as_deref(), Some("#376092"));
    assert_eq!(medium23.second_row_stripe_fill.as_deref(), Some("#B9CDE5"));
    assert!(medium23.stripe_fill.is_none());
    assert_eq!(medium23.first_column_fill.as_deref(), Some("#B9CDE5"));
    assert_eq!(medium23.first_column_bold, Some(true));
    assert_eq!(medium23.subheading_font_color.as_deref(), Some("#000000"));

    // Dark 1-7: shade -0.5 header / total, 0.6 body, 0.4 second stripe.
    let dark2 = palette("PivotStyleDark2");
    assert_eq!(dark2.header_fill.as_deref(), Some("#254061"));
    assert_eq!(dark2.header_font_color.as_deref(), white);
    assert_eq!(dark2.whole_table_fill.as_deref(), Some("#B9CDE5"));
    assert_eq!(dark2.second_row_stripe_fill.as_deref(), Some("#95B3D7"));
    assert!(dark2.stripe_fill.is_none());
    assert_eq!(dark2.subheading_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(dark2.total_row_fill.as_deref(), Some("#254061"));
    // The neutral Dark1 (LO pivot_dark1): dk1 0.5 / 0.75 / 0.65 / 0.85.
    let dark1 = palette("PivotStyleDark1");
    assert_eq!(dark1.header_fill.as_deref(), Some("#808080"));
    assert_eq!(dark1.whole_table_fill.as_deref(), Some("#BFBFBF"));
    assert_eq!(dark1.second_row_stripe_fill.as_deref(), Some("#A6A6A6"));
    assert_eq!(dark1.subheading_fill.as_deref(), Some("#D9D9D9"));

    // Dark 8-14: dk1 tint 0.25 header / total, 0.8 body, 0.6 bands.
    let dark9 = palette("PivotStyleDark9");
    assert_eq!(dark9.header_fill.as_deref(), Some("#404040"));
    assert_eq!(dark9.whole_table_fill.as_deref(), Some("#DCE6F2"));
    assert_eq!(dark9.subheading_fill.as_deref(), Some("#B9CDE5"));
    assert_eq!(dark9.subtotal_fill.as_deref(), Some("#B9CDE5"));
    assert_eq!(dark9.total_row_fill.as_deref(), Some("#404040"));
    assert!(dark9.stripe_fill.is_none());

    // Dark 15-21: dk1 header / total, solid body with tint-0.8 text.
    let dark16 = palette("PivotStyleDark16");
    assert_eq!(dark16.header_fill.as_deref(), Some("#000000"));
    assert_eq!(dark16.whole_table_fill.as_deref(), Some("#4F81BD"));
    assert_eq!(dark16.whole_table_font_color.as_deref(), Some("#DCE6F2"));
    assert_eq!(dark16.subheading_fill.as_deref(), Some("#376092"));
    assert_eq!(dark16.subheading_font_color.as_deref(), white);
    assert_eq!(dark16.total_row_fill.as_deref(), Some("#000000"));
    assert_eq!(
        palette("PivotStyleDark15").whole_table_fill.as_deref(),
        Some("#8C8C8C")
    );

    // Dark 22-28: shaded header / row-label column over the solid body,
    // 0.4 second stripes, no grand-total fill.
    let dark23 = palette("PivotStyleDark23");
    assert_eq!(dark23.header_fill.as_deref(), Some("#376092"));
    assert_eq!(dark23.header_bold, Some(false));
    assert_eq!(dark23.first_header_cell_font_color.as_deref(), white);
    assert_eq!(dark23.first_header_cell_bold, Some(true));
    assert_eq!(dark23.whole_table_fill.as_deref(), Some("#4F81BD"));
    assert_eq!(dark23.whole_table_font_color.as_deref(), Some("#DCE6F2"));
    assert_eq!(dark23.second_row_stripe_fill.as_deref(), Some("#95B3D7"));
    assert_eq!(dark23.second_column_stripe_fill.as_deref(), Some("#95B3D7"));
    assert_eq!(dark23.first_column_fill.as_deref(), Some("#376092"));
    assert!(dark23.total_row_fill.is_none());
    assert_eq!(dark23.total_row_font_color.as_deref(), white);
    assert_eq!(dark23.subtotal_bold, Some(false));
    assert_eq!(palette("PivotStyleDark22").subtotal_bold, Some(true));
    assert_eq!(
        palette("PivotStyleDark22")
            .second_row_stripe_fill
            .as_deref(),
        Some("#8C8C8C")
    );

    // Unknown and custom names resolve to an empty palette.
    assert!(palette("PivotStyleCustom").header_fill.is_none());
    assert!(palette("PivotStyleDark29").header_fill.is_none());
}

/// The wire carries the flattened palette, row kinds and firstDataCol;
/// stripe fills are gated by showRowStripes / showColStripes.
#[test]
fn pivot_wire_carries_bands_and_row_kinds() {
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
            r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1"><location ref="A1:D11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields><pivotField axis="axisRow"/><pivotField axis="axisRow"/><pivotField dataField="1"/></pivotFields><rowFields><field x="0"/><field x="1"/></rowFields><rowItems><i><x/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i><i t="grand"><x/></i></rowItems><pivotTableStyleInfo name="PivotStyleDark2" showRowStripes="1" showColStripes="0"/></pivotTableDefinition>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let pivot = &metadata.sheets[0].pivot_tables[0];
    assert_eq!(pivot.first_data_row, 2);
    assert_eq!(pivot.first_data_col, 1);
    assert_eq!(pivot.row_kinds, "sddtg");
    // Default theme accent1 (#4472C4) tint 0.4 second stripe; column
    // stripes stay off.
    assert_eq!(
        pivot.palette.second_row_stripe_fill.as_deref(),
        Some("#8FAADC")
    );
    assert!(pivot.palette.second_column_stripe_fill.is_none());
    let json = serde_json::to_string(pivot).unwrap();
    assert!(json.contains("\"rowKinds\":\"sddtg\""));
    assert!(json.contains("\"firstDataCol\":1"));
    assert!(json.contains("\"secondRowStripeFill\":\"#8FAADC\""));
    assert!(json.contains("\"headerFontColor\":\"#FFFFFF\""));
    assert!(!json.contains("stripeFill\":null"));
}

/// Office 2007 theme used by the pivot calibration workbook.
fn calib_theme() -> ColorContext {
    ColorContext::with_theme(vec![
        (0xFF, 0xFF, 0xFF),
        (0x00, 0x00, 0x00),
        (0xEE, 0xEC, 0xE1),
        (0x1F, 0x49, 0x7D),
        (0x4F, 0x81, 0xBD),
        (0xC0, 0x50, 0x4D),
        (0x9B, 0xBB, 0x59),
        (0x80, 0x64, 0xA2),
        (0x4B, 0xAC, 0xC6),
        (0xF7, 0x96, 0x46),
        (0x00, 0x00, 0xFF),
        (0x80, 0x00, 0x80),
    ])
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
    assert!(metadata.sheets[0].default_row_height_fixed);
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

/// Opens a one-sheet fixture and returns its conditional rules once the
/// sidecar has indexed the worksheet.
fn conditional_rules_of(sheet_xml: &str) -> Vec<ConditionalRule> {
    let (_dir, path) = open_fixture(&[
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        (
            // Office theme: dk1 black, lt1 white, dk2 #1F497D, lt2 #EEECE1.
            "xl/theme/theme1.xml",
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#,
        ),
    ]);
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    let sheet_id = metadata.sheets[0].id.clone();
    // Fixtures only populate A1; the rules themselves are sheet-level.
    let range = CellRange {
        start_row: 0,
        end_row: 0,
        start_column: 0,
        end_column: 0,
    };
    loop {
        let result = sessions
            .read_range(&metadata.session_id, &sheet_id, &range)
            .unwrap();
        if result.indexing_complete {
            break result.conditional_rules;
        }
        std::thread::sleep(std::time::Duration::from_millis(5));
    }
}

/// Formula-typed thresholds keep their text verbatim (the renderer
/// evaluates them); a 2006-only dataBar reports the ECMA defaults 10/90.
#[test]
fn formula_cfvos_and_legacy_databar_extents() {
    let rules = conditional_rules_of(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<conditionalFormatting sqref="B2:B9"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="num" val="0"/><cfvo type="formula" val="AVERAGE($B$2:$B$9)*1.5"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>
<conditionalFormatting sqref="C2:C9"><cfRule type="dataBar" priority="2"><dataBar minLength="20" maxLength="80" showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF63C384"/></dataBar></cfRule></conditionalFormatting>
<conditionalFormatting sqref="D2:D9"><cfRule type="iconSet" priority="3"><iconSet iconSet="3TrafficLights2"><cfvo type="percent" val="0"/><cfvo type="num" val="0"/><cfvo type="formula" val="$T$5"/></iconSet></cfRule></conditionalFormatting>
</worksheet>"#,
    );
    assert_eq!(rules.len(), 3);
    let bar = &rules[0];
    assert_eq!(bar.cfvos[1].kind, "formula");
    assert_eq!(
        bar.cfvos[1].value.as_deref(),
        Some("AVERAGE($B$2:$B$9)*1.5")
    );
    assert_eq!(bar.min_length, Some(10));
    assert_eq!(bar.max_length, Some(90));
    assert_eq!(bar.axis_position, None);
    assert_eq!(bar.axis_color, None);
    assert_eq!(bar.negative_same_as_positive, None);
    let explicit = &rules[1];
    assert_eq!(explicit.min_length, Some(20));
    assert_eq!(explicit.max_length, Some(80));
    assert!(!explicit.show_value);
    let icons = &rules[2];
    assert_eq!(icons.cfvos[2].kind, "formula");
    assert_eq!(icons.cfvos[2].value.as_deref(), Some("$T$5"));
    // Only dataBar rules carry extents.
    assert_eq!(icons.min_length, None);
    assert_eq!(icons.max_length, None);
}

/// The x14 twin overrides the 2006 extents (its defaults are 0/100) and
/// contributes the axis position and axis colour (rgb or theme+tint).
#[test]
fn x14_databar_twin_merges_axis_and_extents() {
    let rules = conditional_rules_of(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<conditionalFormatting sqref="B2:B5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{AAAA}</x14:id></ext></extLst></cfRule></conditionalFormatting>
<conditionalFormatting sqref="D2:D5"><cfRule type="dataBar" priority="2"><dataBar minLength="10" maxLength="90"><cfvo type="min"/><cfvo type="max"/><color rgb="FF00FF00"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{BBBB}</x14:id></ext></extLst></cfRule></conditionalFormatting>
<conditionalFormatting sqref="F2:F5"><cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{CCCC}</x14:id></ext></extLst></cfRule></conditionalFormatting>
<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>
<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{AAAA}"><x14:dataBar axisPosition="middle" minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFFFF00"/><x14:axisColor rgb="FF000080"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting>
<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{BBBB}"><x14:dataBar gradient="0"><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:axisColor theme="3" tint="0.5"/></x14:dataBar></x14:cfRule><xm:sqref>D2:D5</xm:sqref></x14:conditionalFormatting>
<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{CCCC}"><x14:dataBar minLength="5" maxLength="95" axisPosition="sideways"><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>F2:F5</xm:sqref></x14:conditionalFormatting>
</x14:conditionalFormattings></ext></extLst>
</worksheet>"#,
    );
    assert_eq!(rules.len(), 3);
    let middle = &rules[0];
    assert_eq!(middle.axis_position.as_deref(), Some("middle"));
    assert_eq!(middle.axis_color.as_deref(), Some("#000080"));
    assert_eq!(middle.negative_color.as_deref(), Some("#FFFF00"));
    assert_eq!(middle.min_length, Some(0));
    assert_eq!(middle.max_length, Some(100));
    assert_eq!(middle.cfvos[0].kind, "autoMin");
    assert_eq!(middle.cfvos[1].kind, "autoMax");
    // Twin without explicit extents: x14 defaults (0/100) replace the
    // 2006 element's explicit 10/90; a themed axis colour resolves.
    let themed = &rules[1];
    assert_eq!(themed.min_length, Some(0));
    assert_eq!(themed.max_length, Some(100));
    assert_eq!(themed.axis_position, None);
    assert_eq!(themed.gradient, Some(false));
    // dk2 #1F497D lightened by tint 0.5 (Excel's HLS tint).
    assert_eq!(themed.axis_color.as_deref(), Some("#71A1DC"));
    // Explicit x14 extents win; an unknown axisPosition reads as default.
    let explicit = &rules[2];
    assert_eq!(explicit.min_length, Some(5));
    assert_eq!(explicit.max_length, Some(95));
    assert_eq!(explicit.axis_position, None);
    assert_eq!(explicit.axis_color, None);
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
    // No sheetFormatPr customHeight: the sheet default stays auto-fit.
    assert!(!metadata.sheets[0].default_row_height_fixed);
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
        .map(|cell| {
            (
                cell.row,
                cell.column,
                cell.value.is_some(),
                cell.style_index,
            )
        })
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
    assert!(
        widths
            .iter()
            .any(|width| width.width == Some(12.0) && width.style_index.is_none())
    );
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
<cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="58" applyNumberFormat="1"/><xf numFmtId="44" applyNumberFormat="1"/><xf numFmtId="57" applyNumberFormat="1"/><xf numFmtId="27" applyNumberFormat="1"/><xf numFmtId="30" applyNumberFormat="1"/><xf numFmtId="55" applyNumberFormat="1"/><xf numFmtId="20" applyNumberFormat="1"/><xf numFmtId="21" applyNumberFormat="1"/></cellXfs>
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
    // Id 55 is a date in Excel (ja: yyyy-mm), not the ECMA zh time
    // pattern — a ja month header rendered "0\u{65f6}00\u{5206}" (#049).
    assert_eq!(format(6), Some("yyyy/m/d"));
    // Time builtins carry a leading zero on the hour in Excel (09:30),
    // not the ECMA h:mm text.
    assert_eq!(format(7), Some("hh:mm"));
    assert_eq!(format(8), Some("hh:mm:ss"));

    // Locale-reserved ids must follow the UI locale. In Portuguese (and
    // other day-first locales), id 58 is a full date rather than the
    // zh-CN month/day pattern that previously leaked into every workbook.
    let mut portuguese_sessions = WorkbookSessions::new();
    let portuguese = portuguese_sessions
        .open_with_locale(&path, "pt", None)
        .unwrap();
    let portuguese_format = |index: usize| portuguese.styles[index].number_format.as_deref();
    assert_eq!(portuguese_format(1), Some("d/m/yyyy"));
    assert_eq!(portuguese_format(3), Some("yyyy/m/d"));
    assert_eq!(portuguese_format(4), Some("d/m/yyyy"));
    assert_eq!(portuguese_format(5), Some("d/m/yyyy"));
    assert_eq!(portuguese_format(6), Some("d/m/yyyy"));
}

/// Builtin 14/22 are Excel's OS-region short-date formats; the host
/// passes the system pattern, explicit formatCode entries still win.
#[test]
fn applies_system_short_date_to_builtin_date_formats() {
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
<numFmts count="1"><numFmt numFmtId="165" formatCode="yyyy/mm/dd"/></numFmts>
<fonts count="1"><font/></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" applyNumberFormat="1"/><xf numFmtId="22" applyNumberFormat="1"/><xf numFmtId="165" applyNumberFormat="1"/><xf numFmtId="55" applyNumberFormat="1"/></cellXfs>
</styleSheet>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="1"><v>42438</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);

    let mut sessions = WorkbookSessions::new();
    let metadata = sessions
        .open_with_locale(&path, "en", Some("yyyy/m/d"))
        .unwrap();
    let format = |index: usize| metadata.styles[index].number_format.as_deref();
    assert_eq!(format(1), Some("yyyy/m/d"));
    assert_eq!(format(2), Some("yyyy/m/d hh:mm"));
    assert_eq!(format(3), Some("yyyy/mm/dd"));
    assert_eq!(format(4), Some("yyyy/m/d"));
    assert_eq!(metadata.short_date_format.as_deref(), Some("yyyy/m/d"));

    let mut default_sessions = WorkbookSessions::new();
    let default_metadata = default_sessions.open(&path).unwrap();
    let default_format = |index: usize| default_metadata.styles[index].number_format.as_deref();
    assert_eq!(default_format(1), Some("m/d/yyyy"));
    assert_eq!(default_format(2), Some("m/d/yy hh:mm"));
    assert!(default_metadata.short_date_format.is_none());
}

/// Hancom exports wrap individual styles.xml entries in
/// mc:AlternateContent; a core consumer resolves mc:Fallback (or a lone
/// mc:Choice) instead of dropping the entry and shifting every later id.
#[test]
fn resolves_alternate_content_style_entries() {
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
            // font 1 and xf 1 are Fallback-wrapped, numFmt 165 and dxf 0
            // Choice-only: dropping any of them shifts the later indexes.
            "xl/styles.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:hs="http://schemas.haansoft.com/office/spreadsheet/8.0" mc:Ignorable="hs">
<numFmts count="1"><mc:AlternateContent><mc:Choice Requires="hs"><numFmt numFmtId="165" formatCode="0.000"/></mc:Choice></mc:AlternateContent></numFmts>
<fonts count="3"><font><sz val="11"/></font><mc:AlternateContent><mc:Choice Requires="hs"><font><b/><sz val="20"/></font></mc:Choice><mc:Fallback><font><b/><sz val="18"/></font></mc:Fallback></mc:AlternateContent><font><sz val="9"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><mc:AlternateContent><mc:Fallback><xf fontId="1" applyFont="1"><alignment horizontal="center"/></xf></mc:Fallback></mc:AlternateContent><xf numFmtId="165" fontId="2" applyNumberFormat="1" applyFont="1"/></cellXfs>
<dxfs count="2"><mc:AlternateContent><mc:Fallback><dxf><font><b/></font></dxf></mc:Fallback></mc:AlternateContent><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>
</styleSheet>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2.5</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);

    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert_eq!(metadata.styles.len(), 3);
    // Fallback wins over Choice for the wrapped font (18, not 20).
    assert!(metadata.styles[1].bold);
    assert_eq!(metadata.styles[1].font_size, Some(18.0));
    assert_eq!(
        metadata.styles[1].horizontal_alignment.as_deref(),
        Some("center")
    );
    // Entries after the wrappers keep their declared ids.
    assert_eq!(metadata.styles[2].font_size, Some(9.0));
    assert_eq!(metadata.styles[2].number_format.as_deref(), Some("0.000"));
    assert_eq!(metadata.dxf_styles.len(), 2);
    assert!(metadata.dxf_styles[0].bold);
    assert_eq!(
        metadata.dxf_styles[1].fill_color.as_deref(),
        Some("#FF0000")
    );
}

#[test]
fn parses_shrink_to_fit_alignment() {
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
<fonts count="1"><font><sz val="11"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellXfs count="3"><xf/><xf applyAlignment="1"><alignment horizontal="center" shrinkToFit="1"/></xf><xf applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>
</styleSheet>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="1" t="str"><v>text</v></c></row></sheetData>
</worksheet>"#,
        ),
    ]);

    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(&path).unwrap();
    assert!(!metadata.styles[0].shrink_to_fit);
    assert!(metadata.styles[1].shrink_to_fit);
    assert!(!metadata.styles[2].shrink_to_fit);
    assert!(metadata.styles[2].wrap_text);
}

#[test]
fn blank_cell_keeps_shrink_to_fit_only_style() {
    let default = CellStyle::default();
    let mut style = CellStyle::default();
    style.shrink_to_fit = true;
    assert!(style.styles_blank_cell(&default));
    assert!(!default.styles_blank_cell(&default));
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

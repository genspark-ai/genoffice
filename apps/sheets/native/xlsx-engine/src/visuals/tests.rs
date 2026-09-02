use super::*;

fn row_kinds(pivot_fields: &str, row_fields: &str, row_items: &str) -> String {
    let xml = format!(
        r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1"><location ref="A1:D11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields>{pivot_fields}</pivotFields><rowFields>{row_fields}</rowFields><rowItems>{row_items}</rowItems></pivotTableDefinition>"#
    );
    pivot_row_kinds(&Document::parse(&xml).unwrap())
}

/// Compact form, subtotals at the bottom (the calibration workbook): the
/// outer items are level-1 subheadings, `t="default"` rows subtotals.
#[test]
fn pivot_row_kinds_compact_bottom_subtotals() {
    let kinds = row_kinds(
        r#"<pivotField axis="axisRow" subtotalTop="0"/><pivotField axis="axisRow"/><pivotField dataField="1"/>"#,
        r#"<field x="0"/><field x="1"/>"#,
        r#"<i><x/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i><i><x v="1"/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i t="default"><x v="1"/></i><i t="grand"><x/></i>"#,
    );
    assert_eq!(kinds, "sddtsddtg");
}

/// Three row fields: depth-1 outer items are deeper subheadings, and a
/// blank spacer row keeps its own kind.
#[test]
fn pivot_row_kinds_three_levels_and_blank() {
    let kinds = row_kinds(
        r#"<pivotField axis="axisRow"/><pivotField axis="axisRow"/><pivotField axis="axisRow"/><pivotField dataField="1"/>"#,
        r#"<field x="0"/><field x="1"/><field x="2"/>"#,
        r#"<i><x/></i><i r="1"><x/></i><i r="2"><x/></i><i r="2"><x v="1"/></i><i t="blank"><x/></i><i t="grand"><x/></i>"#,
    );
    assert_eq!(kinds, "sSddbg");
}

/// Tabular fields (outline="0") put the inner item on the outer item's
/// row, so no row is a subheading; a single row field has none either.
#[test]
fn pivot_row_kinds_tabular_and_single_field_are_data() {
    let tabular = row_kinds(
        r#"<pivotField axis="axisRow" outline="0"/><pivotField axis="axisRow"/><pivotField dataField="1"/>"#,
        r#"<field x="0"/><field x="1"/>"#,
        r#"<i><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i><i t="grand"><x/></i>"#,
    );
    assert_eq!(tabular, "ddtg");
    let single = row_kinds(
        r#"<pivotField axis="axisRow"/><pivotField dataField="1"/>"#,
        r#"<field x="0"/>"#,
        r#"<i><x/></i><i><x v="1"/></i><i t="grand"><x/></i>"#,
    );
    assert_eq!(single, "ddg");
    assert_eq!(row_kinds("", "", ""), "");
}

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
fn header_footer_slot_accepts_only_the_six_positions_and_variants() {
    assert_eq!(header_footer_slot("LH").as_deref(), Some("LH"));
    assert_eq!(header_footer_slot(" rf ").as_deref(), Some("RF"));
    assert_eq!(header_footer_slot("CHFIRST").as_deref(), Some("CHFIRST"));
    assert_eq!(header_footer_slot("lfEven").as_deref(), Some("LFEVEN"));
    assert_eq!(header_footer_slot("Comment_x0020_1"), None);
    assert_eq!(header_footer_slot("CHODD"), None);
    assert_eq!(header_footer_slot("XH"), None);
    assert_eq!(header_footer_slot("L"), None);
}

#[test]
fn vml_size_reads_points_and_converts_other_units() {
    assert_eq!(
        vml_size_pt(
            "position:absolute;margin-left:0;margin-top:0;width:442.5pt;height:43.5pt;z-index:1"
        ),
        Some((442.5, 43.5))
    );
    assert_eq!(vml_size_pt("width:2in;height:96px"), Some((144.0, 72.0)));
    assert_eq!(vml_size_pt("width:10pt"), None);
    assert_eq!(vml_size_pt("width:0pt;height:5pt"), None);
    assert_eq!(vml_size_pt("width:50%;height:5pt"), None);
    assert_eq!(vml_size_pt("width:9999pt;height:5pt"), None);
}

#[test]
fn locale_reserved_builtins_follow_viewer_locale() {
    // Long-date id 31: y/m/d order everywhere, kanji/hangul literals on
    // the matching locales (prod refs pin the en rendering to 2025/3/29).
    assert_eq!(builtin_number_format(31, "en"), Some("yyyy/m/d"));
    assert_eq!(
        builtin_number_format(31, "ja"),
        Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\"")
    );
    assert_eq!(
        builtin_number_format(31, "ko"),
        Some("yyyy\"\u{b144}\" mm\"\u{c6d4}\" dd\"\u{c77c}\"")
    );
    assert_eq!(
        builtin_number_format(31, "zh"),
        Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\"")
    );
    // ja 34/35 are dates, not the zh AM/PM times (#049 class).
    assert_eq!(
        builtin_number_format(34, "ja"),
        Some("yyyy\"\u{5e74}\"m\"\u{6708}\"")
    );
    assert_eq!(
        builtin_number_format(33, "ko"),
        Some("h\"\u{c2dc}\" mm\"\u{bd84}\" ss\"\u{cd08}\"")
    );
    // Era-year ja dates render Gregorian.
    assert_eq!(builtin_number_format(27, "ja"), Some("yyyy/m/d"));
    // Non-CJK locales keep their short date outside id 31.
    assert_eq!(builtin_number_format(30, "en"), Some("m/d/yyyy"));
    assert_eq!(builtin_number_format(58, "de"), Some("d.m.yyyy"));
    // 55/56 no-host fallbacks: plain date (CJK) or the short date.
    assert_eq!(builtin_number_format(56, "ja"), Some("yyyy/m/d"));
    assert_eq!(builtin_number_format(55, "pt"), Some("d/m/yyyy"));
    // Builtin 47 renders with a colon (Excel), not ECMA's "mmss.0".
    assert_eq!(builtin_number_format(47, "en"), Some("mm:ss.0"));
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

/// VML `x:Anchor` offsets are pixels; Excel pads the text with newlines.
#[test]
fn vml_anchor_parses_pixel_offsets_into_emu() {
    let anchor = parse_vml_anchor("\n    1, 176, 2, 57, 1, 272, 2, 129").expect("anchor");
    assert_eq!((anchor.from_column, anchor.from_row), (1, 2));
    assert_eq!((anchor.to_column, anchor.to_row), (1, 2));
    assert_eq!(anchor.from_column_offset, 176 * 9525);
    assert_eq!(anchor.from_row_offset, 57 * 9525);
    assert_eq!(anchor.to_column_offset, 272 * 9525);
    assert_eq!(anchor.to_row_offset, 129 * 9525);
    assert!(anchor.explicit_to);
    assert!(parse_vml_anchor("1, 2, 3").is_none());
    assert!(parse_vml_anchor("a, 0, 0, 0, 0, 0, 0, 0").is_none());
    assert!(parse_vml_anchor("-1, 0, 0, 0, 0, 0, 0, 0").is_none());
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
    let absolute = parsed_anchor(r#"<xdr:pos x="100" y="100"/><xdr:ext cx="914400" cy="914400"/>"#);
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
    let chart =
        metadata(r#"<c:title/><c:plotArea><c:barChart><c:ser/><c:ser/></c:barChart></c:plotArea>"#);
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
    let chart = metadata(r#"<c:title/><c:plotArea><c:barChart><c:ser/></c:barChart></c:plotArea>"#);
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

    let formatted = plot("<c:dLbls><c:numFmt formatCode=\"#,##0\" sourceLinked=\"0\"/></c:dLbls>");
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
    let unknown = "<c:plotArea><c:areaChart><c:grouping val=\"weird\"/></c:areaChart></c:plotArea>";
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
fn category_group_spans_follow_the_expanded_idx_slots() {
    // Innermost idx 2 is a blank slot: the expanded categories are
    // [Q1, Q2, "", Q4, Q1] and the groups span idx space directly.
    let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="5"/>
            <c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt><c:pt idx="4"><c:v>Q1</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="4"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
    let chart = metadata(body);
    assert_eq!(chart.series[0].categories, vec!["Q1", "Q2", "", "Q4", "Q1"]);
    let groups = chart.series[0].category_groups.clone().unwrap();
    let spans: Vec<(&str, usize, usize)> = groups
        .iter()
        .map(|group| (group.label.as_str(), group.start, group.end))
        .collect();
    assert_eq!(spans, vec![("2008", 0, 4), ("2009", 4, 5)]);
    // An outer pt without idx cannot be placed: no groups at all.
    let no_idx = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="2"/>
            <c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:lvl>
            <c:lvl><c:pt><c:v>2008</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
    assert!(metadata(no_idx).series[0].category_groups.is_none());
}

#[test]
fn category_groups_and_categories_emit_in_idx_order() {
    // Innermost pts arrive out of idx order: expansion places each at
    // its idx slot, so document order stops mattering.
    let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val="3"/>
            <c:lvl><c:pt idx="2"><c:v>Q1 09</c:v></c:pt><c:pt idx="0"><c:v>Q1 08</c:v></c:pt><c:pt idx="1"><c:v>Q2 08</c:v></c:pt></c:lvl>
            <c:lvl><c:pt idx="0"><c:v>2008</c:v></c:pt><c:pt idx="2"><c:v>2009</c:v></c:pt></c:lvl>
        </c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
    let chart = metadata(body);
    assert_eq!(chart.series[0].categories, vec!["Q1 08", "Q2 08", "Q1 09"]);
    let groups = chart.series[0].category_groups.clone().unwrap();
    let spans: Vec<(&str, usize, usize)> = groups
        .iter()
        .map(|group| (group.label.as_str(), group.start, group.end))
        .collect();
    assert_eq!(spans, vec![("2008", 0, 2), ("2009", 2, 3)]);
}

#[test]
fn sparse_string_caches_expand_into_idx_slots() {
    let body = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="3"/>
            <c:pt idx="2"><c:v>c</c:v></c:pt><c:pt idx="0"><c:v>a</c:v></c:pt>
        </c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
    assert_eq!(metadata(body).series[0].categories, vec!["a", "", "c"]);
    // An idx-less point cannot be placed: document order remains.
    let no_idx = r#"<c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="3"/>
            <c:pt idx="2"><c:v>c</c:v></c:pt><c:pt><c:v>a</c:v></c:pt>
        </c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea>"#;
    assert_eq!(metadata(no_idx).series[0].categories, vec!["c", "a"]);
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
    let with =
        "<c:plotArea><c:barChart/><c:catAx/><c:valAx><c:majorGridlines/></c:valAx></c:plotArea>";
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

/// A template's pre-sized tail (blank categories, cached zeros) must not
/// become ghost axis slots, and sparse value caches align by idx instead
/// of compacting leftwards (prod combo charts).
#[test]
fn trims_blank_tails_and_aligns_sparse_value_caches() {
    let combo = r#"<c:plotArea><c:barChart><c:ser>
            <c:cat><c:numRef><c:numCache><c:ptCount val="6"/><c:pt idx="0"><c:v>42339</c:v></c:pt><c:pt idx="1"><c:v>42370</c:v></c:pt><c:pt idx="2"><c:v>42401</c:v></c:pt></c:numCache></c:numRef></c:cat>
            <c:val><c:numRef><c:numCache><c:ptCount val="6"/><c:pt idx="0"><c:v>1000</c:v></c:pt><c:pt idx="1"><c:v>1150</c:v></c:pt><c:pt idx="2"><c:v>1300</c:v></c:pt><c:pt idx="3"><c:v>0</c:v></c:pt><c:pt idx="4"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart><c:lineChart><c:ser>
            <c:cat><c:numRef><c:numCache><c:ptCount val="6"/><c:pt idx="0"><c:v>42339</c:v></c:pt><c:pt idx="1"><c:v>42370</c:v></c:pt><c:pt idx="2"><c:v>42401</c:v></c:pt></c:numCache></c:numRef></c:cat>
            <c:val><c:numRef><c:numCache><c:ptCount val="6"/><c:pt idx="0"><c:v>1000</c:v></c:pt><c:pt idx="1"><c:v>1160</c:v></c:pt><c:pt idx="2"><c:v>1298.4</c:v></c:pt><c:pt idx="5"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:lineChart></c:plotArea><c:dispBlanksAs val="gap"/>"#;
    let chart = metadata(combo);
    assert_eq!(chart.disp_blanks_as.as_deref(), Some("gap"));
    let bar = &chart.series[0];
    assert_eq!(bar.categories, vec!["42339", "42370", "42401"]);
    assert_eq!(bar.values, vec![1000.0, 1150.0, 1300.0]);
    assert_eq!(bar.blanks, None);
    // The idx 5 zero must not compact onto slot 3; after the trim the
    // line ends with its last real point.
    let line = &chart.series[1];
    assert_eq!(line.values, vec![1000.0, 1160.0, 1298.4]);
    assert_eq!(line.blanks, None);
}

/// Interior blank cells surface as blank markers (values keep a 0
/// filler) so the renderer can apply c:dispBlanksAs.
#[test]
fn marks_interior_blank_values() {
    let body = r#"<c:plotArea><c:lineChart><c:ser>
            <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>
            <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:lineChart></c:plotArea>"#;
    let chart = metadata(body);
    assert!(chart.disp_blanks_as.is_none());
    let series = &chart.series[0];
    assert_eq!(series.values, vec![1.0, 0.0, 3.0]);
    assert_eq!(series.blanks, Some(vec![1]));
}

/// Labelled zero tails are real data; label-less charts keep their
/// numeric fallback ticks — neither may trim.
#[test]
fn keeps_labelled_zero_tails_and_unlabelled_series() {
    let labelled = r#"<c:plotArea><c:barChart><c:ser>
            <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>
            <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart></c:plotArea>"#;
    assert_eq!(metadata(labelled).series[0].values.len(), 3);

    let unlabelled = r#"<c:plotArea><c:barChart><c:ser>
            <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart></c:plotArea>"#;
    assert_eq!(metadata(unlabelled).series[0].values.len(), 3);
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

#[test]
fn sys_clr_solid_fill_resolves() {
    // prod_041: sysClr window/windowText must beat the style fillRef.
    let sppr = r#"<xdr:spPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>
            <a:ln><a:solidFill><a:sysClr val="windowText"/></a:solidFill></a:ln>
            </xdr:spPr>"#;
    let document = Document::parse(sppr).unwrap();
    let root = document.root_element();
    assert_eq!(
        drawing_fill_color(root, &ColorContext::default()).as_deref(),
        Some("#FFFFFF")
    );
    let line = root
        .children()
        .find(|node| node.has_tag_name("ln"))
        .unwrap();
    assert_eq!(
        drawing_fill_color(line, &ColorContext::default()).as_deref(),
        Some("#000000")
    );
}

#[test]
fn parses_src_rect_crop_fractions() {
    let fill = r#"<xdr:blipFill xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:blip/><a:srcRect l="22481" t="11008" r="20930" b="38139"/>
            </xdr:blipFill>"#;
    let document = Document::parse(fill).unwrap();
    let crop = blip_crop(document.root_element()).expect("crop");
    assert_eq!(crop.left, 0.22481);
    assert_eq!(crop.top, 0.11008);
    assert_eq!(crop.right, 0.20930);
    assert_eq!(crop.bottom, 0.38139);
}

#[test]
fn empty_or_degenerate_src_rect_is_no_crop() {
    for rect in [
        r#"<a:srcRect/>"#,
        r#"<a:srcRect l="0" t="0" r="0" b="0"/>"#,
        r#"<a:srcRect l="60000" r="40000"/>"#,
        r#"<a:srcRect t="200000"/>"#,
    ] {
        let fill = format!(
            r#"<xdr:blipFill xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:blip/>{rect}</xdr:blipFill>"#
        );
        let document = Document::parse(&fill).unwrap();
        assert!(blip_crop(document.root_element()).is_none(), "{rect}");
    }
}

#[test]
fn blip_fill_shape_never_falls_back_to_style_fill_ref() {
    // prod_039: a watermark blipFill shape must not paint the fillRef
    // accent as a solid box; the blip alpha rides along as opacity.
    let xml = format!(
        r#"<xdr:wsDr {XDR} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:sp>
              <xdr:nvSpPr><xdr:cNvPr id="10" name="wm"/></xdr:nvSpPr>
              <xdr:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm>
                <a:prstGeom prst="rect"/>
                <a:blipFill><a:blip r:embed="rId1"><a:alphaModFix amt="25000"/></a:blip><a:srcRect/></a:blipFill>
              </xdr:spPr>
              <xdr:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></xdr:style>
            </xdr:sp></xdr:wsDr>"#
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
        explicit_to: true,
    };
    let relationships = HashMap::from([(
        "rId1".to_owned(),
        Relationship {
            target: "../media/image1.png".to_owned(),
            relationship_type: String::new(),
        },
    )]);
    let visual = shape_visual(
        shape,
        anchor,
        "visual-1".into(),
        "sheet1",
        None,
        &theme_colors(),
        "xl/drawings/drawing1.xml",
        &relationships,
        None,
    );
    assert_eq!(
        visual.fill_media_path.as_deref(),
        Some("xl/media/image1.png")
    );
    assert_eq!(visual.fill_color.as_deref(), Some("none"));
    assert!(visual.fill_gradient.is_none());
    assert_eq!(visual.opacity, Some(0.25));
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
    let gradient = style_fill_gradient(Some(document.root_element()), &colors).expect("gradient");
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

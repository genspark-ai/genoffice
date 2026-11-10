//! Table (`<tableParts>`) and pivot style palettes: built-in and custom
//! table styles calibrated against Excel, plus the per-sheet table reader.

use super::*;

pub(crate) const DEFAULT_ACCENTS: [(u8, u8, u8); 6] = [
    (0x44, 0x72, 0xC4),
    (0xED, 0x7D, 0x31),
    (0xA5, 0xA5, 0xA5),
    (0xFF, 0xC0, 0x00),
    (0x5B, 0x9B, 0xD5),
    (0x70, 0xAD, 0x47),
];

pub(crate) fn rgb_hex((red, green, blue): (u8, u8, u8)) -> String {
    format!("#{red:02X}{green:02X}{blue:02X}")
}

/// Fills/font colors a table band can draw, resolved from a custom
/// `<tableStyle>` in styles.xml (each element references a dxf) or from the
/// built-in family rules (builtin_table_palette).
#[derive(Clone, Debug, Default)]
pub(crate) struct CustomTablePalette {
    pub(crate) header_fill: Option<String>,
    pub(crate) header_font_color: Option<String>,
    pub(crate) stripe_fill: Option<String>,
    pub(crate) second_row_stripe_fill: Option<String>,
    pub(crate) column_stripe_fill: Option<String>,
    pub(crate) second_column_stripe_fill: Option<String>,
    pub(crate) whole_table_fill: Option<String>,
    pub(crate) first_column_fill: Option<String>,
    pub(crate) last_column_fill: Option<String>,
    pub(crate) total_row_fill: Option<String>,
    pub(crate) total_row_font_color: Option<String>,
    pub(crate) total_row_border_color: Option<String>,
    pub(crate) total_row_border_style: Option<String>,
    pub(crate) body_font_color: Option<String>,
    pub(crate) first_header_cell_font_color: Option<String>,
    pub(crate) whole_table_border_color: Option<String>,
    pub(crate) whole_table_border_style: Option<String>,
    pub(crate) inner_horizontal_border_color: Option<String>,
    pub(crate) inner_horizontal_border_style: Option<String>,
    pub(crate) inner_vertical_border_color: Option<String>,
    pub(crate) inner_vertical_border_style: Option<String>,
    pub(crate) header_bottom_border_color: Option<String>,
    pub(crate) header_bottom_border_style: Option<String>,
}

/// styles.xml `<tableStyles>` → per-style band palette.
pub(crate) fn read_custom_table_styles(
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
                        // Outline edges are uniform in practice; take left,
                        // fall back to top.
                        let outline = dxf.border_left.as_ref().or(dxf.border_top.as_ref());
                        palette.whole_table_border_color =
                            outline.and_then(|edge| edge.color.clone());
                        palette.whole_table_border_style = outline.map(|edge| edge.style.clone());
                        palette.inner_horizontal_border_color = dxf
                            .border_inner_horizontal
                            .as_ref()
                            .and_then(|edge| edge.color.clone());
                        palette.inner_horizontal_border_style = dxf
                            .border_inner_horizontal
                            .as_ref()
                            .map(|edge| edge.style.clone());
                        palette.inner_vertical_border_color = dxf
                            .border_inner_vertical
                            .as_ref()
                            .and_then(|edge| edge.color.clone());
                        palette.inner_vertical_border_style = dxf
                            .border_inner_vertical
                            .as_ref()
                            .map(|edge| edge.style.clone());
                    }
                    Some("headerRow") => {
                        palette.header_fill = dxf.fill_color.clone();
                        palette.header_font_color = dxf.font_color.clone();
                        palette.header_bottom_border_color = dxf
                            .border_bottom
                            .as_ref()
                            .and_then(|edge| edge.color.clone());
                        palette.header_bottom_border_style =
                            dxf.border_bottom.as_ref().map(|edge| edge.style.clone());
                    }
                    Some("totalRow") => {
                        palette.total_row_fill = dxf.fill_color.clone();
                        palette.total_row_font_color = dxf.font_color.clone();
                        palette.total_row_border_color =
                            dxf.border_top.as_ref().and_then(|edge| edge.color.clone());
                        palette.total_row_border_style =
                            dxf.border_top.as_ref().map(|edge| edge.style.clone());
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

/// Resolved bands of a built-in pivot style (see `pivot_style_palette`).
/// Serialized flat into `PivotTableInfo`; every field is optional so the
/// renderer inherits down the band precedence for anything unset.
#[derive(Clone, Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotStylePalette {
    /// Column-header rows (row offsets below firstDataRow).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_bold: Option<bool>,
    /// Top-left cell of the output area (Excel's "First Header Cell"), above
    /// the header band.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header_cell_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header_cell_bold: Option<bool>,
    /// Every cell of the output area, lowest layer.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_font_color: Option<String>,
    /// Row stripes: even row offsets from firstDataRow take `stripe_fill`,
    /// odd ones `second_row_stripe_fill`; only emitted with showRowStripes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub second_row_stripe_fill: Option<String>,
    /// Column stripes over the data columns (offsets from firstDataCol);
    /// only emitted with showColStripes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub column_stripe_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub second_column_stripe_fill: Option<String>,
    /// Row-label columns left of firstDataCol (Excel's "Row Headers").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column_bold: Option<bool>,
    /// Level-1 row subheading rows (outer row items, `row_kinds` 's').
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading_bold: Option<bool>,
    /// Deeper row subheadings (`row_kinds` 'S').
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading2_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading2_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subheading2_bold: Option<bool>,
    /// Subtotal rows (`row_kinds` 't').
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subtotal_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subtotal_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subtotal_bold: Option<bool>,
    /// Grand-total row (`row_kinds` 'g').
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_bold: Option<bool>,
}

/// Built-in pivot style bands, vector-calibrated against Excel for Mac
/// (calibration workbook: genoffice-sample/sheets/calib/gen_pivot_calib.py,
/// truth table pivot-style-truths.json). Accent cycle is (n-1) % 7 with
/// 0 = dk1 neutral, variants come in blocks of 7. Neutral members tint
/// their bands lighter than the accent members (0.8 → 0.85, 0.6 → 0.75,
/// 0.4 → 0.65) and use dk1 tint 0.5 where accents use shade -0.25.
///
/// Row stripes run in two different directions: the Light blocks and
/// Medium 15-21 paint the EVEN offsets from firstDataRow (first stripe),
/// Dark 1-7, Medium 22-28 and Dark 22-28 paint the ODD offsets (second
/// stripe) over the whole-table body — the opposite of the table styles.
/// The stripe count includes subheading / subtotal rows, and a band with
/// its own fill (subheading, total) covers the stripe.
pub(crate) fn pivot_style_palette(
    style_name: Option<&str>,
    colors: &ColorContext,
) -> PivotStylePalette {
    const WHITE: &str = "#FFFFFF";
    let mut palette = PivotStylePalette::default();
    let Some(name) = style_name else {
        return palette;
    };
    let (family, number) = if let Some(rest) = name.strip_prefix("PivotStyleLight") {
        ("light", rest.parse::<usize>().ok())
    } else if let Some(rest) = name.strip_prefix("PivotStyleMedium") {
        ("medium", rest.parse::<usize>().ok())
    } else if let Some(rest) = name.strip_prefix("PivotStyleDark") {
        ("dark", rest.parse::<usize>().ok())
    } else {
        ("", None)
    };
    let Some(number) = number.filter(|number| (1..=28).contains(number)) else {
        return palette;
    };
    let dark1 = visuals::theme_dark1(colors).unwrap_or((0x00, 0x00, 0x00));
    let accent_index = (number - 1) % 7;
    let neutral = accent_index == 0;
    let base = if neutral {
        dark1
    } else {
        visuals::theme_accent(colors, accent_index).unwrap_or(DEFAULT_ACCENTS[accent_index - 1])
    };
    let black = rgb_hex(dark1);
    let white = || Some(WHITE.to_owned());
    // Band tint of the accent member / of the dk1 neutral member.
    let band = |accent_tint: f64, neutral_tint: f64| {
        visuals::tint_to_hex(base, if neutral { neutral_tint } else { accent_tint })
    };
    let light = || band(0.8, 0.85);
    let mid = || band(0.6, 0.75);
    let deep = || band(0.4, 0.65);
    // Excel's "darker" step: accent shade -0.25, dk1 tint 0.5 (mid gray).
    let shade = || band(-0.25, 0.5);
    let solid = || band(0.0, 0.5);
    let gray = |tint: f64| visuals::tint_to_hex(dark1, tint);
    let variant = (number - 1) / 7;
    match (family, variant) {
        // Light 1-7: no fills at all (accent rules only), first stripe
        // tint 0.8, deeper subheadings in accent text (Light1: gray bold);
        // the grand total is painted explicit white so stripes stop above it.
        ("light", 0) => {
            palette.header_bold = Some(true);
            palette.stripe_fill = Some(light());
            palette.column_stripe_fill = Some(light());
            palette.subheading_bold = Some(true);
            palette.subheading2_font_color = Some(solid());
            // Excel's Light2 is the one member whose deeper subheadings
            // are not bold.
            palette.subheading2_bold = Some(number != 2);
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = white();
            palette.total_row_bold = Some(true);
        }
        // Light 8-14: unfilled header, body text in the accent shade,
        // tint-0.8 subheading and subtotal bands, no stripes (Excel draws
        // the banding as rules).
        ("light", 1) => {
            palette.header_font_color = Some(black.clone());
            palette.header_bold = Some(true);
            palette.whole_table_font_color = Some(if neutral { black.clone() } else { shade() });
            palette.subheading_fill = Some(light());
            palette.subheading_font_color = Some(black.clone());
            palette.subheading_bold = Some(true);
            palette.subheading2_font_color = Some(black.clone());
            palette.subheading2_bold = Some(true);
            palette.subtotal_fill = Some(light());
            palette.subtotal_font_color = Some(black.clone());
            palette.subtotal_bold = Some(true);
            palette.total_row_font_color = Some(black);
            palette.total_row_bold = Some(true);
        }
        // Light 15-21: tint-0.8 header and grand total, dk1-gray stripes for
        // every member (Light16 stripes #D9D9D9, not accent).
        ("light", 2) => {
            palette.header_fill = Some(light());
            palette.header_bold = Some(true);
            palette.stripe_fill = Some(gray(0.85));
            palette.column_stripe_fill = Some(gray(0.85));
            palette.subheading_bold = Some(true);
            palette.subheading2_bold = Some(true);
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = Some(light());
            palette.total_row_bold = Some(true);
        }
        // Light 22-28: no fills, all text in the accent shade with no bold;
        // the stripe (tint 0.8) also runs over the unfilled grand total.
        // The neutral Light22 keeps black text and bolds its row-label
        // column, subtotals, header and grand total instead (subheading
        // rows inherit the column bold, their values stay regular).
        ("light", _) => {
            palette.header_bold = Some(neutral);
            palette.whole_table_font_color = Some(if neutral { black } else { shade() });
            palette.stripe_fill = Some(light());
            palette.column_stripe_fill = Some(light());
            palette.first_column_bold = Some(neutral);
            palette.subtotal_bold = Some(neutral);
            palette.total_row_bold = Some(neutral);
        }
        // Medium 1-7: shaded header with white regular text, tint-0.4
        // subheading (white text) and subtotal (white bold), tint-0.8 deeper
        // subheadings, unfilled body and grand total.
        ("medium", 0) => {
            palette.header_fill = Some(shade());
            palette.header_font_color = white();
            palette.header_bold = Some(false);
            palette.first_header_cell_bold = Some(true);
            palette.subheading_fill = Some(deep());
            palette.subheading_font_color = white();
            palette.subheading_bold = Some(false);
            palette.subheading2_fill = Some(light());
            palette.subheading2_bold = Some(false);
            palette.subtotal_fill = Some(deep());
            palette.subtotal_font_color = white();
            palette.subtotal_bold = Some(true);
            palette.total_row_bold = Some(true);
        }
        // Medium 8-14: solid accent header with white bold text
        // (aspose_sample1, Medium9), tint-0.8 subheading, tint-0.6 subtotal,
        // unfilled body and grand total; stripes are rules, not fills.
        ("medium", 1) => {
            palette.header_fill = Some(solid());
            palette.header_font_color = white();
            palette.header_bold = Some(true);
            palette.subheading_fill = Some(light());
            palette.subheading_bold = Some(true);
            palette.subheading2_bold = Some(true);
            palette.subtotal_fill = Some(mid());
            palette.subtotal_bold = Some(true);
            palette.total_row_bold = Some(true);
        }
        // Medium 15-21: dk1 header and grand total with white regular text
        // over a tint-0.8 body (neutral 0.95); the first stripe is tint 0.8
        // too, so it only shows on the neutral member (0.85 on 0.95).
        ("medium", 2) => {
            palette.header_fill = Some(black.clone());
            palette.header_font_color = white();
            palette.header_bold = Some(false);
            palette.whole_table_fill = Some(band(0.8, 0.95));
            palette.stripe_fill = Some(light());
            palette.column_stripe_fill = Some(light());
            palette.subheading_bold = Some(true);
            palette.subheading2_font_color = Some(gray(0.5));
            palette.subheading2_bold = Some(true);
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = Some(black);
            palette.total_row_font_color = white();
            palette.total_row_bold = Some(false);
        }
        // Medium 22-28: whole table tint 0.8 with shaded text, the
        // row-label column tint 0.6 in bold, the second stripe tint 0.6;
        // header and grand total are unfilled (they show the body/column
        // fills) with shaded bold text; subheading/subtotal rows black bold.
        ("medium", _) => {
            palette.header_bold = Some(true);
            palette.whole_table_fill = Some(light());
            palette.whole_table_font_color = Some(shade());
            palette.second_row_stripe_fill = Some(mid());
            palette.second_column_stripe_fill = Some(mid());
            palette.first_column_fill = Some(mid());
            palette.first_column_bold = Some(true);
            palette.subheading_font_color = Some(black.clone());
            palette.subheading_bold = Some(true);
            palette.subtotal_font_color = Some(black);
            palette.subtotal_bold = Some(true);
            palette.total_row_bold = Some(true);
        }
        // Dark 1-7: shade -0.5 header and grand total (neutral dk1 tint 0.5)
        // with white bold text, tint-0.6 body, tint-0.4 second stripe,
        // tint-0.8 level-1 subheading (pivot_dark1: 0.5 / 0.75 / 0.65 /
        // 0.85 on dk1).
        ("dark", 0) => {
            let dark = band(-0.5, 0.5);
            palette.header_fill = Some(dark.clone());
            palette.header_font_color = white();
            palette.header_bold = Some(true);
            palette.whole_table_fill = Some(mid());
            palette.second_row_stripe_fill = Some(deep());
            palette.subheading_fill = Some(light());
            palette.subheading_bold = Some(true);
            palette.subheading2_bold = Some(true);
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = Some(dark);
            palette.total_row_font_color = white();
            palette.total_row_bold = Some(true);
        }
        // Dark 8-14: dk1 tint 0.25 header and grand total for every member,
        // tint-0.8 body, tint-0.6 subheading and subtotal, no stripes.
        ("dark", 1) => {
            palette.header_fill = Some(gray(0.25));
            palette.header_font_color = white();
            palette.header_bold = Some(true);
            palette.whole_table_fill = Some(light());
            palette.subheading_fill = Some(mid());
            palette.subheading_bold = Some(true);
            palette.subheading2_bold = Some(true);
            palette.subtotal_fill = Some(mid());
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = Some(gray(0.25));
            palette.total_row_font_color = white();
            palette.total_row_bold = Some(true);
        }
        // Dark 15-21: dk1 header and grand total, solid accent body (neutral
        // dk1 tint 0.55) with tint-0.8 text, shaded level-1 subheading and
        // white bold subheading / subtotal text, no stripes.
        ("dark", 2) => {
            palette.header_fill = Some(black.clone());
            palette.header_font_color = white();
            palette.header_bold = Some(true);
            palette.whole_table_fill = Some(band(0.0, 0.55));
            palette.whole_table_font_color = Some(light());
            palette.subheading_fill = Some(shade());
            palette.subheading_font_color = white();
            palette.subheading_bold = Some(true);
            palette.subheading2_font_color = white();
            palette.subheading2_bold = Some(true);
            palette.subtotal_font_color = white();
            palette.subtotal_bold = Some(true);
            palette.total_row_fill = Some(black);
            palette.total_row_font_color = white();
            palette.total_row_bold = Some(true);
        }
        // Dark 22-28: shaded header and row-label column over a solid accent
        // body (neutral: dk1 tint 0.5 throughout) with tint-0.8 text and a
        // tint-0.4 second stripe (neutral 0.55); the first header cell and
        // subheadings are white bold, subtotals plain (Dark22: bold), and
        // the grand total row has no fill of its own (stripes show through)
        // but white bold text.
        ("dark", _) => {
            palette.header_fill = Some(shade());
            palette.header_bold = Some(false);
            palette.first_header_cell_font_color = white();
            palette.first_header_cell_bold = Some(true);
            palette.whole_table_fill = Some(solid());
            palette.whole_table_font_color = Some(light());
            palette.second_row_stripe_fill = Some(band(0.4, 0.55));
            palette.second_column_stripe_fill = Some(band(0.4, 0.55));
            palette.first_column_fill = Some(shade());
            palette.first_column_bold = Some(false);
            palette.subheading_font_color = white();
            palette.subheading_bold = Some(true);
            palette.subheading2_font_color = white();
            palette.subheading2_bold = Some(true);
            palette.subtotal_bold = Some(neutral);
            palette.total_row_font_color = white();
            palette.total_row_bold = Some(true);
        }
        _ => {}
    }
    palette
}

/// Table style column 1 of the Light block (Light1-7) draws its frame in the
/// base color; the filled families' borders stay unmodelled.
pub(crate) fn table_style_border(
    style_name: Option<&str>,
    colors: &ColorContext,
) -> Option<String> {
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

/// Built-in table style bands, pixel-calibrated against Excel for Mac
/// (calibration workbook: genoffice-sample/sheets/calib). Accent cycle is
/// (n-1) % 7 with 0 = dk1, variants come in blocks of 7; dk1-based members
/// tint their bands 0.05 lighter than the accent members do.
pub(crate) fn builtin_table_palette(
    style_name: Option<&str>,
    colors: &ColorContext,
) -> CustomTablePalette {
    const WHITE: &str = "#FFFFFF";
    let mut palette = CustomTablePalette::default();
    // A tableStyleInfo without a name (or none at all) is Excel's style
    // "None": nothing gets painted.
    let Some(name) = style_name else {
        return palette;
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
    let accent = |index: usize| {
        // Wrap so an out-of-range family number can't index past accent6.
        let index = (index - 1) % 6 + 1;
        visuals::theme_accent(colors, index).unwrap_or(DEFAULT_ACCENTS[index - 1])
    };
    let dark1 = visuals::theme_dark1(colors).unwrap_or((0x00, 0x00, 0x00));
    let accent_index = number.saturating_sub(1) % 7;
    let neutral = accent_index == 0;
    let base = if neutral { dark1 } else { accent(accent_index) };
    let band = |color, from_dark1: bool, tint: f64| {
        visuals::tint_to_hex(color, if from_dark1 { tint + 0.05 } else { tint })
    };
    let variant = number.saturating_sub(1) / 7;
    match (family, variant) {
        // Light 1-7: unfilled shaded-accent header, single accent rule over
        // the totals band.
        ("light", 0) => {
            palette.header_font_color = Some(visuals::tint_to_hex(base, -0.25));
            palette.stripe_fill = Some(band(base, neutral, 0.8));
            palette.total_row_font_color = Some(visuals::tint_to_hex(base, -0.25));
            palette.total_row_border_color = Some(rgb_hex(base));
            palette.total_row_border_style = Some("thin".into());
        }
        // Light 8-14: solid header, banding drawn as row RULES, not fills —
        // no stripe color (tablerefsnamed renders all-white).
        ("light", 1) => {
            palette.header_fill = Some(rgb_hex(base));
            palette.header_font_color = Some(WHITE.into());
            palette.total_row_border_color = Some(rgb_hex(base));
            palette.total_row_border_style = Some("double".into());
        }
        // Light 15-21: unfilled header in plain dk1 (only 1-7 color it), and
        // the whole table gridded — thin base-color outline plus inner
        // horizontal/vertical rules on every cell (ref: Light16 draws a full
        // accent1 grid over the banding).
        ("light", _) => {
            palette.header_font_color = Some(rgb_hex(dark1));
            palette.stripe_fill = Some(band(base, neutral, 0.8));
            palette.total_row_border_color = Some(rgb_hex(base));
            palette.total_row_border_style = Some("double".into());
            palette.whole_table_border_color = Some(rgb_hex(base));
            palette.whole_table_border_style = Some("thin".into());
            palette.inner_horizontal_border_color = Some(rgb_hex(base));
            palette.inner_horizontal_border_style = Some("thin".into());
            palette.inner_vertical_border_color = Some(rgb_hex(base));
            palette.inner_vertical_border_style = Some("thin".into());
        }
        // Medium 8-14 are Excel's "full color" block: both bands filled
        // (Medium9: accent tint 0.6 alternating with 0.8, #B8CCE4 /
        // #DCE6F1) and the totals band solid accent under white text.
        ("medium", 1) => {
            palette.header_fill = Some(rgb_hex(base));
            palette.header_font_color = Some(WHITE.into());
            palette.stripe_fill = Some(band(base, neutral, 0.6));
            palette.second_row_stripe_fill = Some(band(base, neutral, 0.8));
            palette.whole_table_fill = Some(band(base, neutral, 0.8));
            palette.total_row_fill = Some(rgb_hex(base));
            palette.total_row_font_color = Some(WHITE.into());
        }
        // Medium 15-21: accent header over dk1-gray banding, dk1 totals rule.
        ("medium", 2) => {
            palette.header_fill = Some(rgb_hex(base));
            palette.header_font_color = Some(WHITE.into());
            palette.stripe_fill = Some(visuals::tint_to_hex(dark1, 0.85));
            palette.total_row_border_color = Some(rgb_hex(dark1));
            palette.total_row_border_style = Some("double".into());
        }
        // Medium 22-28: whole table tinted (0.8 body / 0.6 stripe), header
        // included, dk1 text throughout, accent rule over the totals band.
        ("medium", 3) => {
            palette.header_fill = Some(band(base, neutral, 0.8));
            palette.header_font_color = Some(rgb_hex(dark1));
            palette.stripe_fill = Some(band(base, neutral, 0.6));
            palette.whole_table_fill = Some(band(base, neutral, 0.8));
            palette.total_row_fill = Some(band(base, neutral, 0.8));
            palette.total_row_border_color = Some(rgb_hex(base));
            palette.total_row_border_style = Some("medium".into());
        }
        // Medium 1-7 (and unknown style names, which resolve like the default
        // Medium 2): solid header, tint-0.8 stripe, double accent rule over
        // an unfilled totals band.
        ("medium", _) => {
            palette.header_fill = Some(rgb_hex(base));
            palette.header_font_color = Some(WHITE.into());
            palette.stripe_fill = Some(band(base, neutral, 0.8));
            palette.total_row_border_color = Some(rgb_hex(base));
            palette.total_row_border_style = Some("double".into());
        }
        // Dark 8-11 pair two accents: header on accent 2/4/6, bands on
        // accent 1/3/5 (Dark 8 runs both on dk1).
        ("dark", 1) => {
            let band_base = if neutral {
                dark1
            } else {
                accent(2 * (number - 8) - 1)
            };
            let header_base = if neutral {
                dark1
            } else {
                accent(2 * (number - 8))
            };
            palette.header_fill = Some(rgb_hex(header_base));
            palette.header_font_color = Some(WHITE.into());
            palette.stripe_fill = Some(band(band_base, neutral, 0.6));
            palette.second_row_stripe_fill = Some(band(band_base, neutral, 0.8));
            palette.whole_table_fill = Some(band(band_base, neutral, 0.8));
            palette.total_row_fill = Some(band(band_base, neutral, 0.8));
            palette.total_row_border_color = Some(rgb_hex(dark1));
            palette.total_row_border_style = Some("double".into());
        }
        // Dark 1-7: dk1 header over a solid accent body, stripe and totals
        // shaded darker (the dk1 member tints toward white instead).
        ("dark", _) => {
            palette.header_fill = Some(rgb_hex(dark1));
            palette.header_font_color = Some(WHITE.into());
            if neutral {
                palette.whole_table_fill = Some(visuals::tint_to_hex(dark1, 0.45));
                palette.stripe_fill = Some(visuals::tint_to_hex(dark1, 0.25));
                palette.total_row_fill = Some(visuals::tint_to_hex(dark1, 0.15));
            } else {
                palette.whole_table_fill = Some(rgb_hex(base));
                palette.stripe_fill = Some(visuals::tint_to_hex(base, -0.25));
                palette.total_row_fill = Some(visuals::tint_to_hex(base, -0.5));
            }
            palette.body_font_color = Some(WHITE.into());
            palette.total_row_font_color = Some(WHITE.into());
        }
        _ => unreachable!("family is one of light/medium/dark"),
    }
    palette
}

pub(crate) fn read_sheet_tables(
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
        let mut filter_active = false;
        loop {
            match reader.read_event()? {
                // Criteria elements only occur under autoFilter/filterColumn;
                // a bare <filterColumn hiddenButton=…/> is not a live filter.
                Event::Start(element) | Event::Empty(element)
                    if matches!(
                        element.local_name().as_ref(),
                        b"filters"
                            | b"customFilters"
                            | b"dynamicFilter"
                            | b"top10"
                            | b"colorFilter"
                            | b"iconFilter"
                    ) =>
                {
                    filter_active = true;
                }
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
                    show_column_stripes = attribute_value(&reader, &element, b"showColumnStripes")?
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
            let palette = match custom {
                Some(palette) => palette.clone(),
                None => builtin_table_palette(style_name.as_deref(), colors),
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
                filter_active,
                name,
                columns,
                style_name,
                header_fill: palette.header_fill,
                header_font_color: palette.header_font_color,
                stripe_fill: palette.stripe_fill,
                second_row_stripe_fill: palette.second_row_stripe_fill,
                column_stripe_fill: palette.column_stripe_fill,
                second_column_stripe_fill: palette.second_column_stripe_fill,
                whole_table_fill: palette.whole_table_fill,
                // First/last column emphasis (and the first-header corner
                // cell) only paint when tableStyleInfo turns them on.
                first_column_fill: palette.first_column_fill.filter(|_| show_first_column),
                last_column_fill: palette.last_column_fill.filter(|_| show_last_column),
                total_row_fill: palette.total_row_fill,
                total_row_font_color: palette.total_row_font_color,
                total_row_border_color: palette.total_row_border_color,
                total_row_border_style: palette.total_row_border_style,
                body_font_color: palette.body_font_color,
                first_header_cell_font_color: palette
                    .first_header_cell_font_color
                    .filter(|_| show_first_column),
                totals_row_count,
                border_color,
                whole_table_border_color: palette.whole_table_border_color,
                whole_table_border_style: palette.whole_table_border_style,
                inner_horizontal_border_color: palette.inner_horizontal_border_color,
                inner_horizontal_border_style: palette.inner_horizontal_border_style,
                inner_vertical_border_color: palette.inner_vertical_border_color,
                inner_vertical_border_style: palette.inner_vertical_border_style,
                header_bottom_border_color: palette.header_bottom_border_color,
                header_bottom_border_style: palette.header_bottom_border_style,
            });
        }
    }
    Ok(tables)
}

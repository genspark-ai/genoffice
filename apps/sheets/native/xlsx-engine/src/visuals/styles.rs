//! `styles.xml` reader: cellXfs / dxf resolution into `CellStyle`, plus the
//! font, fill and border element parsers it is built from.

use super::*;

pub fn read_styles(
    archive: &mut ZipArchive<File>,
    colors: &ColorContext,
    theme_fonts: Option<&ThemeFonts>,
    locale: &str,
    short_date_format: Option<&str>,
) -> Result<(Vec<CellStyle>, Vec<CellStyle>, Option<String>), SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/styles.xml")? else {
        return Ok((vec![CellStyle::default()], Vec::new(), None));
    };
    let document = parse_document(&xml, "styles.xml")?;
    // Only the top-level <numFmts> table: dxf-local <numFmt> entries reuse
    // file-local ids that must not shadow builtin ids for cell xfs.
    let custom_formats = document
        .descendants()
        .filter(|node| node.has_tag_name("numFmts"))
        .flat_map(|node| mc_children(node, "numFmt"))
        .filter_map(|node| {
            Some((
                node.attribute("numFmtId")?.parse::<u32>().ok()?,
                node.attribute("formatCode")?.to_owned(),
            ))
        })
        .collect::<HashMap<_, _>>();
    let fonts = document
        .descendants()
        .find(|node| node.has_tag_name("fonts"))
        .map(|node| {
            mc_children(node, "font")
                .into_iter()
                .map(|font| parse_font(font, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let fills = document
        .descendants()
        .find(|node| node.has_tag_name("fills"))
        .map(|node| {
            mc_children(node, "fill")
                .into_iter()
                .map(|fill| parse_fill(fill, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let borders = document
        .descendants()
        .find(|node| node.has_tag_name("borders"))
        .map(|node| {
            mc_children(node, "border")
                .into_iter()
                .map(|border| parse_border(border, &colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let cell_xfs = document
        .descendants()
        .find(|node| node.has_tag_name("cellXfs"));
    // Literal cached <name val> of the Normal (cellXfs[0]) font: for scheme
    // fonts the theme substitution below erases it, but Excel derives the
    // column-width MDW from this face (a ja workbook caches the locale
    // resolution, e.g. MS PGothic, while the theme latin says Calibri).
    let normal_font_name = cell_xfs
        .and_then(|node| mc_children(node, "xf").into_iter().next())
        .and_then(|xf| numeric_attribute(xf, "fontId"))
        .and_then(|index| fonts.get(index))
        .and_then(|font| font.family.clone());
    let styles = cell_xfs
        .map(|node| {
            mc_children(node, "xf")
                .into_iter()
                .map(|xf| {
                    let font = numeric_attribute(xf, "fontId")
                        .and_then(|index| fonts.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let fill = numeric_attribute(xf, "fillId")
                        .and_then(|index| fills.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let border = numeric_attribute(xf, "borderId")
                        .and_then(|index| borders.get(index))
                        .cloned()
                        .unwrap_or_default();
                    let number_format = numeric_attribute(xf, "numFmtId").and_then(|id| {
                        custom_formats
                            .get(&(id as u32))
                            .cloned()
                            .or_else(|| short_date_number_format(id as u32, short_date_format))
                            .or_else(|| {
                                builtin_number_format(id as u32, locale).map(ToOwned::to_owned)
                            })
                    });
                    let alignment = xf.children().find(|child| child.has_tag_name("alignment"));
                    // Excel resolves scheme fonts against the theme; the
                    // literal <name val> is only a cached copy.
                    let font_family = match (font.scheme.as_deref(), theme_fonts) {
                        (Some("major"), Some(fonts)) => Some(fonts.major.clone()),
                        (Some("minor"), Some(fonts)) => Some(fonts.minor.clone()),
                        _ => font.family,
                    };
                    CellStyle {
                        font_family,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        underline: font.underline,
                        strikethrough: font.strikethrough,
                        wrap_text: alignment
                            .and_then(|node| node.attribute("wrapText"))
                            .is_some_and(|value| value == "1" || value == "true"),
                        shrink_to_fit: alignment
                            .and_then(|node| node.attribute("shrinkToFit"))
                            .is_some_and(|value| value == "1" || value == "true"),
                        font_color: font.color,
                        fill_color: fill.color,
                        font_color_theme: font.color_theme,
                        font_color_tint: font.color_tint,
                        fill_color_theme: fill.theme,
                        fill_color_tint: fill.tint,
                        font_scheme: font.scheme,
                        horizontal_alignment: alignment
                            .and_then(|node| node.attribute("horizontal"))
                            .map(ToOwned::to_owned),
                        vertical_alignment: alignment
                            .and_then(|node| node.attribute("vertical"))
                            .map(ToOwned::to_owned),
                        indent: alignment
                            .and_then(|node| node.attribute("indent"))
                            .and_then(|value| value.parse::<u32>().ok())
                            .filter(|steps| *steps > 0),
                        text_rotation: alignment
                            .and_then(|node| node.attribute("textRotation"))
                            .and_then(|value| value.parse::<u32>().ok())
                            .filter(|degrees| (1..=180).contains(degrees) || *degrees == 255),
                        number_format,
                        border_top: border.top,
                        border_bottom: border.bottom,
                        border_left: border.left,
                        border_right: border.right,
                        border_diagonal: border.diagonal,
                        diagonal_up: border.diagonal_up,
                        diagonal_down: border.diagonal_down,
                        border_inner_horizontal: None,
                        border_inner_vertical: None,
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let dxfs = document
        .descendants()
        .find(|node| node.has_tag_name("dxfs"))
        .map(|node| {
            mc_children(node, "dxf")
                .into_iter()
                .map(|dxf| parse_dxf(dxf, colors))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let styles = if styles.is_empty() {
        vec![CellStyle::default()]
    } else {
        styles
    };
    Ok((styles, dxfs, normal_font_name))
}

/// Differential (dxf) styles referenced by conditional-formatting rules.
/// Solid dxf fills carry the color in bgColor, unlike cell fills.
pub(crate) fn parse_dxf(dxf: Node<'_, '_>, colors: &ColorContext) -> CellStyle {
    let font = dxf
        .children()
        .find(|node| node.has_tag_name("font"))
        .map(|node| parse_font(node, colors))
        .unwrap_or_default();
    let fill_color = dxf
        .children()
        .find(|node| node.has_tag_name("fill"))
        .and_then(|fill| {
            let pattern = fill
                .children()
                .find(|node| node.has_tag_name("patternFill"))?;
            pattern
                .children()
                .find(|node| node.has_tag_name("bgColor"))
                .or_else(|| pattern.children().find(|node| node.has_tag_name("fgColor")))
                .and_then(|node| parse_color(node, colors))
        });
    let border = dxf
        .children()
        .find(|node| node.has_tag_name("border"))
        .map(|node| parse_border(node, colors))
        .unwrap_or_default();
    // Unlike cell xfs, a dxf carries its format code inline.
    let number_format = dxf
        .children()
        .find(|node| node.has_tag_name("numFmt"))
        .and_then(|node| node.attribute("formatCode"))
        .filter(|code| !code.is_empty() && *code != "General")
        .map(ToOwned::to_owned);
    CellStyle {
        font_family: font.family,
        font_size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        wrap_text: false,
        shrink_to_fit: false,
        font_color: font.color,
        fill_color,
        font_color_theme: None,
        font_color_tint: None,
        fill_color_theme: None,
        fill_color_tint: None,
        font_scheme: None,
        horizontal_alignment: None,
        vertical_alignment: None,
        indent: None,
        text_rotation: None,
        number_format,
        border_top: border.top,
        border_bottom: border.bottom,
        border_left: border.left,
        border_right: border.right,
        border_diagonal: border.diagonal,
        diagonal_up: border.diagonal_up,
        diagonal_down: border.diagonal_down,
        border_inner_horizontal: border.horizontal,
        border_inner_vertical: border.vertical,
    }
}

pub(crate) fn parse_font(font: Node<'_, '_>, colors: &ColorContext) -> FontStyle {
    let color_node = font.children().find(|node| node.has_tag_name("color"));
    let color = color_node.and_then(|node| parse_color(node, colors));
    let (color_theme, color_tint) = color_node
        .filter(|_| color.is_some())
        .map(|node| theme_provenance(node, colors))
        .unwrap_or((None, None));
    FontStyle {
        family: font
            .children()
            .find(|node| node.has_tag_name("name"))
            .and_then(|node| node.attribute("val"))
            .map(ToOwned::to_owned),
        size: font
            .children()
            .find(|node| node.has_tag_name("sz"))
            .and_then(|node| node.attribute("val"))
            .and_then(|value| value.parse::<f64>().ok()),
        bold: font
            .children()
            .find(|node| node.has_tag_name("b"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        italic: font
            .children()
            .find(|node| node.has_tag_name("i"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        underline: font
            .children()
            .find(|node| node.has_tag_name("u"))
            .is_some_and(|node| node.attribute("val") != Some("none")),
        strikethrough: font
            .children()
            .find(|node| node.has_tag_name("strike"))
            .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        color,
        color_theme,
        color_tint,
        scheme: font
            .children()
            .find(|node| node.has_tag_name("scheme"))
            .and_then(|node| node.attribute("val"))
            .filter(|value| *value == "major" || *value == "minor")
            .map(ToOwned::to_owned),
    }
}

/// Theme slot + tint of a color node, when the color resolves through the
/// theme palette (mirrors resolve_color: a resolvable theme slot wins over
/// rgb/indexed).
pub(crate) fn theme_provenance(
    node: Node<'_, '_>,
    colors: &ColorContext,
) -> (Option<usize>, Option<f64>) {
    let theme = node
        .attribute("theme")
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|index| colors.theme.get(*index).is_some());
    if theme.is_none() {
        return (None, None);
    }
    // Clamped to the wire schema's [-1, 1]; apply_tint saturates the same
    // way, so a clamped tint reproduces the baked color exactly.
    let tint = node
        .attribute("tint")
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .map(|value| value.clamp(-1.0, 1.0))
        .filter(|value| *value != 0.0);
    (theme, tint)
}

#[derive(Clone, Default)]
pub(crate) struct FillInfo {
    pub(crate) color: Option<String>,
    pub(crate) theme: Option<usize>,
    pub(crate) tint: Option<f64>,
}

pub(crate) fn parse_fill(fill: Node<'_, '_>, colors: &ColorContext) -> FillInfo {
    let Some(pattern) = fill
        .children()
        .find(|node| node.has_tag_name("patternFill"))
    else {
        return parse_gradient_fill(fill, colors);
    };
    let pattern_type = pattern.attribute("patternType");
    if pattern_type == Some("none") {
        return FillInfo::default();
    }
    let fg_node = pattern.children().find(|node| node.has_tag_name("fgColor"));
    let bg_node = pattern.children().find(|node| node.has_tag_name("bgColor"));
    let foreground = fg_node.and_then(|node| parse_color(node, colors));
    let background = bg_node.and_then(|node| parse_color(node, colors));
    // Textured patterns (gray125, stripes, …) render as the per-channel blend
    // of both colors — the closest flat-color approximation of the texture.
    // Blends carry no single theme provenance.
    if pattern_type.is_some_and(|value| value != "solid") {
        if let (Some(fg), Some(bg)) = (foreground.as_deref(), background.as_deref()) {
            if let Some(mixed) = mix_hex(fg, bg) {
                return FillInfo {
                    color: Some(mixed),
                    ..FillInfo::default()
                };
            }
        }
    }
    let (color, node) = if foreground.is_some() {
        (foreground, fg_node)
    } else {
        (background, bg_node)
    };
    let (theme, tint) = node
        .filter(|_| color.is_some())
        .map(|node| theme_provenance(node, colors))
        .unwrap_or((None, None));
    FillInfo { color, theme, tint }
}

/// The cell style model carries one flat color, so a gradientFill is
/// approximated by the mid-gradient blend of its outermost stops. The blend
/// carries no single theme provenance.
pub(crate) fn parse_gradient_fill(fill: Node<'_, '_>, colors: &ColorContext) -> FillInfo {
    let Some(gradient) = fill
        .children()
        .find(|node| node.has_tag_name("gradientFill"))
    else {
        return FillInfo::default();
    };
    let mut stops = gradient
        .children()
        .filter(|node| node.has_tag_name("stop"))
        .filter_map(|stop| {
            let position = stop.attribute("position")?.parse::<f64>().ok()?;
            let color = stop
                .children()
                .find(|node| node.has_tag_name("color"))
                .and_then(|node| parse_color(node, colors))?;
            Some((position, color))
        })
        .collect::<Vec<_>>();
    // Stops are not guaranteed document-ordered.
    stops.sort_by(|a, b| a.0.total_cmp(&b.0));
    let (Some((_, first)), Some((_, last))) = (stops.first(), stops.last()) else {
        return FillInfo::default();
    };
    FillInfo {
        color: mix_hex(first, last),
        ..FillInfo::default()
    }
}

pub(crate) fn mix_hex(first: &str, second: &str) -> Option<String> {
    let parse = |hex: &str| -> Option<(u8, u8, u8)> {
        let value = hex.strip_prefix('#')?;
        Some((
            u8::from_str_radix(value.get(0..2)?, 16).ok()?,
            u8::from_str_radix(value.get(2..4)?, 16).ok()?,
            u8::from_str_radix(value.get(4..6)?, 16).ok()?,
        ))
    };
    let (r1, g1, b1) = parse(first)?;
    let (r2, g2, b2) = parse(second)?;
    Some(format!(
        "#{:02X}{:02X}{:02X}",
        (u16::from(r1) + u16::from(r2)) / 2,
        (u16::from(g1) + u16::from(g2)) / 2,
        (u16::from(b1) + u16::from(b2)) / 2,
    ))
}

pub(crate) fn parse_border(border: Node<'_, '_>, colors: &ColorContext) -> BorderSet {
    let edge = |name: &str| -> Option<BorderEdge> {
        let node = border.children().find(|child| child.has_tag_name(name))?;
        let style = node.attribute("style")?;
        if style == "none" {
            return None;
        }
        Some(BorderEdge {
            style: style.to_owned(),
            color: node
                .children()
                .find(|child| child.has_tag_name("color"))
                .and_then(|child| parse_color(child, colors)),
        })
    };
    BorderSet {
        top: edge("top"),
        bottom: edge("bottom"),
        left: edge("left"),
        right: edge("right"),
        diagonal: edge("diagonal"),
        vertical: edge("vertical"),
        horizontal: edge("horizontal"),
        diagonal_up: border
            .attribute("diagonalUp")
            .is_some_and(|value| value == "1" || value == "true"),
        diagonal_down: border
            .attribute("diagonalDown")
            .is_some_and(|value| value == "1" || value == "true"),
    }
}

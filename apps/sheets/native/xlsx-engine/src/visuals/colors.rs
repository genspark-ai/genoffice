//! Color resolution: indexed palette, theme palette and tints, scheme /
//! modifier chains on drawing fills, and the RGB <-> HSL helpers.

use super::*;

/// DrawingML solid fill: srgbClr, or schemeClr resolved via the theme palette.
pub(crate) fn drawing_fill_color(node: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    // An a:ln below `node` carries the outline color, not the fill.
    let outside_outline = |child: &Node<'_, '_>| {
        !child
            .ancestors()
            .take_while(|ancestor| *ancestor != node)
            .any(|ancestor| ancestor.has_tag_name("ln"))
    };
    let fill = node
        .descendants()
        .find(|child| child.has_tag_name("solidFill") && outside_outline(child))
        .or_else(|| {
            // Gradient fills approximate to their first stop color.
            node.descendants()
                .find(|child| child.has_tag_name("gradFill") && outside_outline(child))
                .and_then(|grad| grad.descendants().find(|child| child.has_tag_name("gs")))
        })?;
    if let Some(srgb) = fill
        .descendants()
        .find(|child| child.has_tag_name("srgbClr"))
    {
        return srgb.attribute("val").map(|value| format!("#{value}"));
    }
    if let Some(sys) = fill
        .descendants()
        .find(|child| child.has_tag_name("sysClr"))
    {
        // lastClr is Excel's resolved cache; window/windowText carry fixed
        // meanings when it is missing.
        let base = sys
            .attribute("lastClr")
            .and_then(parse_hex_rgb)
            .or_else(|| match sys.attribute("val") {
                Some("window") => Some((0xFF, 0xFF, 0xFF)),
                Some("windowText") => Some((0, 0, 0)),
                _ => None,
            })?;
        return Some(apply_color_modifiers(sys, base));
    }
    let scheme = fill
        .descendants()
        .find(|child| child.has_tag_name("schemeClr"))?;
    let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
    Some(apply_color_modifiers(scheme, base))
}

/// DrawingML child color transforms (a:lumMod/a:lumOff/a:tint/a:shade), the
/// usual RGB approximations — enough for Excel's accent "40% lighter" dPt
/// and legend variants.
pub(crate) fn apply_color_modifiers(color_node: Node<'_, '_>, base: (u8, u8, u8)) -> String {
    let modifiers = color_node
        .children()
        .filter(|child| child.is_element())
        .filter_map(|child| {
            Some((
                child.tag_name().name().to_owned(),
                child.attribute("val")?.parse::<f64>().ok()? / 100_000.0,
            ))
        })
        .collect::<Vec<_>>();
    apply_modifier_values(base, &modifiers)
}

/// xdr:style fillRef pointing at a fmtScheme gradient entry: each phClr
/// stop resolves against the reference's scheme color.
pub(crate) fn style_fill_gradient(
    style_node: Option<Node<'_, '_>>,
    colors: &ColorContext,
) -> Option<FillGradient> {
    let reference = style_node?
        .children()
        .find(|node| node.has_tag_name("fillRef"))?;
    let idx = reference.attribute("idx")?.parse::<usize>().ok()?;
    let gradient = colors.fill_styles.get(idx.checked_sub(1)?)?.as_ref()?;
    let scheme = reference
        .children()
        .find(|node| node.has_tag_name("schemeClr"))?;
    let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
    let base = parse_hex_rgb(apply_color_modifiers(scheme, base).trim_start_matches('#'))?;
    let stops = gradient
        .stops
        .iter()
        .map(|stop| FillGradientStop {
            position: stop.position,
            color: apply_modifier_values(base, &stop.modifiers),
        })
        .collect();
    Some(FillGradient {
        angle: gradient.angle,
        stops,
    })
}

/// Data-driven twin of apply_color_modifiers, for theme gradient stops whose
/// source document is gone by the time the fillRef color is known.
pub(crate) fn apply_modifier_values(base: (u8, u8, u8), modifiers: &[(String, f64)]) -> String {
    let mut channels = [f64::from(base.0), f64::from(base.1), f64::from(base.2)];
    for (name, value) in modifiers {
        let value = *value;
        match name.as_str() {
            "lumMod" | "shade" => {
                for channel in &mut channels {
                    *channel *= value;
                }
            }
            "lumOff" => {
                for channel in &mut channels {
                    *channel += 255.0 * value;
                }
            }
            "tint" => {
                for channel in &mut channels {
                    *channel = *channel * value + 255.0 * (1.0 - value);
                }
            }
            "satMod" => {
                let clamp = |v: f64| v.round().clamp(0.0, 255.0) as u8;
                let (hue, saturation, luminance) =
                    rgb_to_hsl((clamp(channels[0]), clamp(channels[1]), clamp(channels[2])));
                let (red, green, blue) = hsl_to_rgb(hue, (saturation * value).min(1.0), luminance);
                channels = [f64::from(red), f64::from(green), f64::from(blue)];
            }
            _ => {}
        }
    }
    let clamp = |value: f64| value.round().clamp(0.0, 255.0) as u8;
    format!(
        "#{:02X}{:02X}{:02X}",
        clamp(channels[0]),
        clamp(channels[1]),
        clamp(channels[2])
    )
}

pub(crate) fn scheme_color_rgb(name: &str, colors: &ColorContext) -> Option<(u8, u8, u8)> {
    if let Some(rest) = name.strip_prefix("accent") {
        return theme_accent(colors, rest.parse::<usize>().ok()?);
    }
    let index = match name {
        "lt1" | "bg1" => 0,
        "dk1" | "tx1" => 1,
        "lt2" | "bg2" => 2,
        "dk2" | "tx2" => 3,
        _ => return None,
    };
    colors.theme.get(index).copied()
}

/// Legacy indexed palette, ECMA-376 §18.8.27. Indexes 64/65 are the system
/// window text/background colors.
pub(crate) const INDEXED_COLORS: [&str; 66] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
    "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
    "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
    "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
    "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
    "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
    "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333", "000000", "FFFFFF",
];

pub(crate) fn parse_color(node: Node<'_, '_>, colors: &ColorContext) -> Option<String> {
    resolve_color(
        node.attribute("rgb"),
        node.attribute("indexed"),
        node.attribute("theme"),
        node.attribute("tint"),
        colors,
    )
}

pub fn resolve_color(
    rgb: Option<&str>,
    indexed: Option<&str>,
    theme: Option<&str>,
    tint: Option<&str>,
    colors: &ColorContext,
) -> Option<String> {
    // A resolvable theme slot wins over rgb: Excel treats rgb as a cached
    // copy of the theme color, and some producers bake a wrong cache
    // (tdf113271 writes theme="1" rgb="FFFFFF" for black dk1 text).
    if let Some(base) = theme
        .and_then(|value| value.parse::<usize>().ok())
        .and_then(|index| colors.theme.get(index).copied())
    {
        let tint = tint
            .and_then(|value| value.parse::<f64>().ok())
            .unwrap_or(0.0);
        let (red, green, blue) = apply_tint(base, tint);
        return Some(format!("#{red:02X}{green:02X}{blue:02X}"));
    }
    if let Some(rgb) = rgb {
        let value = if rgb.len() == 8 { &rgb[2..] } else { rgb };
        return Some(format!("#{value}"));
    }
    let index = indexed?.parse::<usize>().ok()?;
    // 64/65 are the fixed system window text/background slots — producers
    // that override the palette still expect the system colors there.
    if index < 64 {
        if let Some(value) = colors.indexed.get(index) {
            return Some(format!("#{value}"));
        }
    }
    INDEXED_COLORS.get(index).map(|value| format!("#{value}"))
}

/// styles.xml `<colors><indexedColors>` — a legacy-palette override written
/// by workbooks converted from .xls. Entries are ARGB ("00RRGGBB").
pub fn read_indexed_palette(
    archive: &mut ZipArchive<File>,
    colors: &mut ColorContext,
) -> Result<(), SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/styles.xml")? else {
        return Ok(());
    };
    let document = parse_document(&xml, "styles.xml")?;
    let Some(list) = document
        .descendants()
        .find(|node| node.has_tag_name("indexedColors"))
    else {
        return Ok(());
    };
    colors.indexed = list
        .children()
        .filter(|node| node.has_tag_name("rgbColor"))
        .filter_map(|node| node.attribute("rgb"))
        .map(|value| {
            let hex = if value.len() == 8 { &value[2..] } else { value };
            hex.to_owned()
        })
        .collect();
    Ok(())
}

/// Theme accent color (1-6) as rgb, if the palette was loaded.
pub fn theme_accent(colors: &ColorContext, accent: usize) -> Option<(u8, u8, u8)> {
    // Effective palette order: [lt1, dk1, lt2, dk2, accent1-6, ...]
    colors.theme.get(3 + accent).copied()
}

/// Theme dk1 (neutral text) color as rgb, if the palette was loaded.
pub fn theme_dark1(colors: &ColorContext) -> Option<(u8, u8, u8)> {
    colors.theme.get(1).copied()
}

pub fn tint_to_hex(base: (u8, u8, u8), tint: f64) -> String {
    let (red, green, blue) = apply_tint(base, tint);
    format!("#{red:02X}{green:02X}{blue:02X}")
}

pub fn read_theme_palette(archive: &mut ZipArchive<File>) -> Result<ColorContext, SidecarError> {
    let Some(xml) = read_optional_xml(archive, "xl/theme/theme1.xml")? else {
        return Ok(ColorContext::default());
    };
    let document = parse_document(&xml, "theme1.xml")?;
    let Some(scheme) = document
        .descendants()
        .find(|node| node.has_tag_name("clrScheme"))
    else {
        return Ok(ColorContext::default());
    };
    let slot = |name: &str| -> Option<(u8, u8, u8)> {
        let node = scheme.children().find(|child| child.has_tag_name(name))?;
        let hex = node
            .children()
            .find(|child| child.has_tag_name("srgbClr"))
            .and_then(|child| child.attribute("val"))
            .or_else(|| {
                node.children()
                    .find(|child| child.has_tag_name("sysClr"))
                    .and_then(|child| child.attribute("lastClr"))
            })?;
        parse_hex_rgb(hex)
    };
    // The `theme` attribute indexes [lt1, dk1, lt2, dk2, accent1-6, hlink,
    // folHlink] — light/dark pairs swapped versus clrScheme document order.
    let order = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    let mut theme = Vec::with_capacity(order.len());
    for name in order {
        match slot(name) {
            Some(color) => theme.push(color),
            None => return Ok(ColorContext::default()),
        }
    }
    let fill_styles = document
        .descendants()
        .find(|node| node.has_tag_name("fillStyleLst"))
        .map(|list| {
            list.children()
                .filter(|node| node.is_element())
                .map(|node| parse_theme_gradient(node))
                .collect()
        })
        .unwrap_or_default();
    Ok(ColorContext {
        theme,
        fill_styles,
        indexed: Vec::new(),
    })
}

/// A fillStyleLst gradFill entry as data (the theme document does not
/// outlive the palette); non-gradient entries map to None.
pub(crate) fn parse_theme_gradient(node: Node<'_, '_>) -> Option<ThemeGradient> {
    if !node.has_tag_name("gradFill") {
        return None;
    }
    let stops = node
        .descendants()
        .filter(|child| child.has_tag_name("gs"))
        .filter_map(|gs| {
            let position = gs.attribute("pos")?.parse::<f64>().ok()? / 100_000.0;
            let color = gs
                .children()
                .find(|child| child.has_tag_name("schemeClr") || child.has_tag_name("srgbClr"))?;
            let modifiers = color
                .children()
                .filter(|child| child.is_element())
                .filter_map(|child| {
                    Some((
                        child.tag_name().name().to_owned(),
                        child.attribute("val")?.parse::<f64>().ok()? / 100_000.0,
                    ))
                })
                .collect();
            Some(ThemeGradientStop {
                position,
                modifiers,
            })
        })
        .collect::<Vec<_>>();
    if stops.len() < 2 {
        return None;
    }
    let angle = node
        .descendants()
        .find(|child| child.has_tag_name("lin"))
        .and_then(|lin| lin.attribute("ang"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| value / 60_000.0)
        .unwrap_or(90.0);
    Some(ThemeGradient { stops, angle })
}

pub(crate) fn parse_hex_rgb(hex: &str) -> Option<(u8, u8, u8)> {
    let value = if hex.len() == 8 { &hex[2..] } else { hex };
    if value.len() != 6 {
        return None;
    }
    Some((
        u8::from_str_radix(&value[0..2], 16).ok()?,
        u8::from_str_radix(&value[2..4], 16).ok()?,
        u8::from_str_radix(&value[4..6], 16).ok()?,
    ))
}

/// Excel's tint transform: scale HSL luminance toward black (tint < 0) or
/// white (tint > 0).
pub(crate) fn apply_tint(rgb: (u8, u8, u8), tint: f64) -> (u8, u8, u8) {
    if tint == 0.0 {
        return rgb;
    }
    let (hue, saturation, luminance) = rgb_to_hsl(rgb);
    let luminance = if tint < 0.0 {
        luminance * (1.0 + tint)
    } else {
        luminance * (1.0 - tint) + tint
    };
    hsl_to_rgb(hue, saturation, luminance.clamp(0.0, 1.0))
}

pub(crate) fn rgb_to_hsl((red, green, blue): (u8, u8, u8)) -> (f64, f64, f64) {
    let red = f64::from(red) / 255.0;
    let green = f64::from(green) / 255.0;
    let blue = f64::from(blue) / 255.0;
    let maximum = red.max(green).max(blue);
    let minimum = red.min(green).min(blue);
    let luminance = (maximum + minimum) / 2.0;
    if maximum == minimum {
        return (0.0, 0.0, luminance);
    }
    let delta = maximum - minimum;
    let saturation = if luminance > 0.5 {
        delta / (2.0 - maximum - minimum)
    } else {
        delta / (maximum + minimum)
    };
    let hue = if maximum == red {
        (green - blue) / delta + if green < blue { 6.0 } else { 0.0 }
    } else if maximum == green {
        (blue - red) / delta + 2.0
    } else {
        (red - green) / delta + 4.0
    } / 6.0;
    (hue, saturation, luminance)
}

pub(crate) fn hsl_to_rgb(hue: f64, saturation: f64, luminance: f64) -> (u8, u8, u8) {
    if saturation == 0.0 {
        let value = (luminance * 255.0).round() as u8;
        return (value, value, value);
    }
    let q = if luminance < 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let p = 2.0 * luminance - q;
    let channel = |mut t: f64| -> u8 {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        let value = if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 1.0 / 2.0 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        };
        (value * 255.0).round() as u8
    };
    (
        channel(hue + 1.0 / 3.0),
        channel(hue),
        channel(hue - 1.0 / 3.0),
    )
}

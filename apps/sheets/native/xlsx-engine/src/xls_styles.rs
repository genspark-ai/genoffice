//! Cell formatting of a BIFF8 workbook — FONT, FORMAT, XF, PALETTE and XFEXT
//! records from the globals substream — rebuilt as an OOXML `styles.xml`. Each XF
//! record becomes one `cellXfs` entry at the same index, so a cell's `ixfe`
//! maps to `s` by a fixed offset and nothing is deduplicated or reordered.

use std::collections::HashMap;

use crate::visuals::apply_tint;
use crate::xlsx_package::escape_xml;

/// Cell XF indices start after the converter's own three defaults (general,
/// short date, date+time) so files without a binary layer keep working.
pub(crate) const XF_BASE: u16 = 3;

#[derive(Default, Clone)]
pub(crate) struct XlsFont {
    pub(crate) twips: u16,
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) strike: bool,
    pub(crate) underline: u8,
    pub(crate) escapement: u16,
    pub(crate) color: u16,
    pub(crate) name: String,
}

#[derive(Default, Clone)]
pub(crate) struct XlsXf {
    pub(crate) font: u16,
    pub(crate) format: u16,
    pub(crate) locked: bool,
    pub(crate) hidden: bool,
    /// A style XF (the Normal/Heading… definitions) rather than a cell XF.
    pub(crate) is_style: bool,
    pub(crate) h_align: u8,
    pub(crate) v_align: u8,
    pub(crate) wrap: bool,
    pub(crate) rotation: u8,
    pub(crate) indent: u8,
    pub(crate) shrink: bool,
    /// (style, color index) for left, right, top, bottom, diagonal.
    pub(crate) borders: [(u8, u16); 5],
    pub(crate) diag_down: bool,
    pub(crate) diag_up: bool,
    pub(crate) pattern: u8,
    pub(crate) fore: u16,
    pub(crate) back: u16,
}

/// A FullColorExt: the color itself and the tint (-1..1) laid over it.
#[derive(Clone, Copy)]
pub(crate) enum XlsExtColor {
    Indexed(u16, f64),
    Rgb([u8; 3], f64),
    Theme(usize, f64),
}

/// The colors an XFEXT record adds to one XF. Excel 2007+ keeps the real
/// color here and leaves the nearest palette index in the XF for older readers.
#[derive(Default, Clone)]
pub(crate) struct XlsXfExt {
    pub(crate) fore: Option<XlsExtColor>,
    pub(crate) back: Option<XlsExtColor>,
    /// left, right, top, bottom, diagonal — the XF's border order
    pub(crate) borders: [Option<XlsExtColor>; 5],
    pub(crate) text: Option<XlsExtColor>,
}

#[derive(Default)]
pub(crate) struct XlsStyles {
    pub(crate) fonts: Vec<XlsFont>,
    pub(crate) formats: HashMap<u16, String>,
    pub(crate) xfs: Vec<XlsXf>,
    pub(crate) palette: Option<Vec<[u8; 3]>>,
    pub(crate) xf_exts: HashMap<u16, XlsXfExt>,
    /// XFCRC's XF count: the extensions describe the XF table as it was then
    pub(crate) xf_ext_count: Option<u16>,
    /// The THEME record's colors in `theme` attribute order (lt1, dk1, lt2, dk2, accents…).
    pub(crate) theme: Option<Vec<[u8; 3]>>,
}

impl XlsStyles {
    pub(crate) fn is_empty(&self) -> bool {
        self.xfs.is_empty()
    }

    pub(crate) fn cell_style(&self, ixfe: u16) -> Option<u16> {
        (usize::from(ixfe) < self.xfs.len()).then_some(XF_BASE + ixfe)
    }

    pub(crate) fn push_font(&mut self, bytes: &[u8]) {
        if bytes.len() < 16 {
            return;
        }
        let u16_at = |at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]);
        let grbit = u16_at(2);
        self.fonts.push(XlsFont {
            twips: u16_at(0),
            italic: grbit & 0x0002 != 0,
            strike: grbit & 0x0008 != 0,
            color: u16_at(4),
            bold: u16_at(6) >= 0x02BC,
            escapement: u16_at(8),
            underline: bytes[10],
            name: short_unicode_string(&bytes[14..]),
        });
    }

    pub(crate) fn push_format(&mut self, bytes: &[u8]) {
        if bytes.len() < 5 {
            return;
        }
        let id = u16::from_le_bytes([bytes[0], bytes[1]]);
        let length = usize::from(u16::from_le_bytes([bytes[2], bytes[3]]));
        let code = unicode_string(&bytes[4..], length);
        self.formats.insert(id, code);
    }

    pub(crate) fn push_xf(&mut self, bytes: &[u8]) {
        if bytes.len() < 20 {
            return;
        }
        let u16_at = |at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]);
        let u32_at = |at: usize| {
            u32::from_le_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]])
        };
        let flags = u16_at(4);
        let border1 = u32_at(10);
        let border2 = u32_at(14);
        let fill = u16_at(18);
        self.xfs.push(XlsXf {
            font: u16_at(0),
            format: u16_at(2),
            locked: flags & 0x0001 != 0,
            hidden: flags & 0x0002 != 0,
            is_style: flags & 0x0004 != 0,
            h_align: bytes[6] & 0x07,
            wrap: bytes[6] & 0x08 != 0,
            v_align: (bytes[6] >> 4) & 0x07,
            rotation: bytes[7],
            indent: bytes[8] & 0x0F,
            shrink: bytes[8] & 0x10 != 0,
            borders: [
                ((border1 & 0xF) as u8, ((border1 >> 16) & 0x7F) as u16),
                (
                    ((border1 >> 4) & 0xF) as u8,
                    ((border1 >> 23) & 0x7F) as u16,
                ),
                (((border1 >> 8) & 0xF) as u8, (border2 & 0x7F) as u16),
                (
                    ((border1 >> 12) & 0xF) as u8,
                    ((border2 >> 7) & 0x7F) as u16,
                ),
                (
                    ((border2 >> 21) & 0xF) as u8,
                    ((border2 >> 14) & 0x7F) as u16,
                ),
            ],
            diag_down: border1 & 0x4000_0000 != 0,
            diag_up: border1 & 0x8000_0000 != 0,
            pattern: ((border2 >> 26) & 0x3F) as u8,
            fore: fill & 0x7F,
            back: (fill >> 7) & 0x7F,
        });
    }

    pub(crate) fn set_palette(&mut self, bytes: &[u8]) {
        if bytes.len() < 2 {
            return;
        }
        let count = usize::from(u16::from_le_bytes([bytes[0], bytes[1]]));
        let colors: Vec<[u8; 3]> = bytes[2..]
            .chunks_exact(4)
            .take(count)
            .map(|c| [c[0], c[1], c[2]])
            .collect();
        if !colors.is_empty() {
            self.palette = Some(colors);
        }
    }

    pub(crate) fn set_xf_crc(&mut self, bytes: &[u8]) {
        if bytes.len() >= 16 {
            self.xf_ext_count = Some(u16::from_le_bytes([bytes[14], bytes[15]]));
        }
    }

    /// XFEXT: FrtHeader (12), reserved, ixfe, reserved, cexts, then ExtProps
    /// of (type, size including this header, data).
    pub(crate) fn push_xf_ext(&mut self, bytes: &[u8]) {
        if bytes.len() < 20 {
            return;
        }
        let u16_at = |at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]);
        let ixfe = u16_at(14);
        let count = usize::from(u16_at(18));
        let mut ext = XlsXfExt::default();
        let mut at = 20usize;
        for _ in 0..count {
            if at + 4 > bytes.len() {
                break;
            }
            let kind = u16_at(at);
            let size = usize::from(u16_at(at + 2));
            if size < 4 || at + size > bytes.len() {
                break;
            }
            let color = Self::full_color_ext(&bytes[at + 4..at + size]);
            match kind {
                0x0004 => ext.fore = color,
                0x0005 => ext.back = color,
                0x0009 => ext.borders[0] = color,
                0x000A => ext.borders[1] = color,
                0x0007 => ext.borders[2] = color,
                0x0008 => ext.borders[3] = color,
                0x000B => ext.borders[4] = color,
                0x000D => ext.text = color,
                _ => {}
            }
            at += size;
        }
        self.xf_exts.insert(ixfe, ext);
    }

    /// THEME: FrtHeader (12), version, then the theme package when the
    /// workbook does not use the default theme of that version.
    pub(crate) fn set_theme(&mut self, bytes: &[u8]) {
        let Some(package) = bytes.get(16..).filter(|package| !package.is_empty()) else {
            return;
        };
        let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(package)) else {
            return;
        };
        let mut xml = String::new();
        // a crafted package must not inflate without bound
        for index in 0..archive.len().min(THEME_MAX_ENTRIES) {
            let Ok(entry) = archive.by_index(index) else {
                continue;
            };
            if entry.name().is_ok_and(|name| name.ends_with("theme1.xml")) {
                let mut bounded = std::io::Read::take(entry, THEME_MAX_BYTES + 1);
                let read = std::io::Read::read_to_string(&mut bounded, &mut xml);
                if read.map_or(true, |size| size as u64 > THEME_MAX_BYTES) {
                    return;
                }
                break;
            }
        }
        let colors: Vec<[u8; 3]> = THEME_SLOTS
            .iter()
            .map_while(|slot| theme_slot_rgb(&xml, slot))
            .collect();
        if colors.len() == THEME_SLOTS.len() {
            self.theme = Some(colors);
        }
    }

    fn full_color_ext(bytes: &[u8]) -> Option<XlsExtColor> {
        if bytes.len() < 8 {
            return None;
        }
        let tint = f64::from(i16::from_le_bytes([bytes[2], bytes[3]])) / 32767.0;
        let value = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        match u16::from_le_bytes([bytes[0], bytes[1]]) {
            1 => Some(XlsExtColor::Indexed(u16::try_from(value).ok()?, tint)),
            2 => Some(XlsExtColor::Rgb([bytes[4], bytes[5], bytes[6]], tint)),
            3 => Some(XlsExtColor::Theme(usize::try_from(value).ok()?, tint)),
            // automatic / not set: the XF's own palette index stands
            _ => None,
        }
    }

    /// An extension color as `FFRRGGBB`, theme and tint resolved so the
    /// converted workbook needs no theme part.
    fn ext_rgb(&self, color: XlsExtColor) -> Option<String> {
        let (rgb, tint) = match color {
            XlsExtColor::Indexed(icv, tint) => return self.rgb(icv).filter(|_| tint == 0.0),
            XlsExtColor::Rgb(rgb, tint) => (rgb, tint),
            XlsExtColor::Theme(index, tint) => {
                let theme = self.theme.as_deref().unwrap_or(&DEFAULT_THEME);
                (*theme.get(index)?, tint)
            }
        };
        let (r, g, b) = apply_tint((rgb[0], rgb[1], rgb[2]), tint);
        Some(format!("FF{r:02X}{g:02X}{b:02X}"))
    }

    fn ext_of(&self, index: usize) -> Option<&XlsXfExt> {
        // a writer that rebuilt the XF table without the extensions leaves them stale
        if self
            .xf_ext_count
            .is_some_and(|count| usize::from(count) != self.xfs.len())
        {
            return None;
        }
        self.xf_exts.get(&u16::try_from(index).ok()?)
    }

    /// `FFRRGGBB` for a palette index, unless the extension names the real color.
    fn color_rgb(&self, icv: u16, ext: Option<XlsExtColor>) -> Option<String> {
        ext.and_then(|color| self.ext_rgb(color))
            .or_else(|| self.rgb(icv))
    }

    /// `rgb` attribute value for a palette index; `None` for the automatic
    /// and system colors, which OOXML expresses by leaving the color out.
    fn rgb(&self, icv: u16) -> Option<String> {
        // 0–7 are the fixed built-in colors, the same eight the palette opens with
        let [r, g, b] = if icv < 8 {
            DEFAULT_PALETTE[usize::from(icv)]
        } else {
            let slot = usize::from(icv - 8);
            match &self.palette {
                Some(palette) => *palette.get(slot)?,
                None => *DEFAULT_PALETTE.get(slot)?,
            }
        };
        Some(format!("FF{r:02X}{g:02X}{b:02X}"))
    }

    fn color_tag(&self, icv: u16, ext: Option<XlsExtColor>) -> String {
        match self.color_rgb(icv, ext) {
            Some(rgb) => format!(r#"<color rgb="{rgb}"/>"#),
            None => String::new(),
        }
    }

    fn font_xml(&self, font: &XlsFont, color: Option<XlsExtColor>) -> String {
        let mut out = String::from("<font>");
        if font.bold {
            out.push_str("<b/>");
        }
        if font.italic {
            out.push_str("<i/>");
        }
        if font.strike {
            out.push_str("<strike/>");
        }
        match font.underline {
            0x01 => out.push_str("<u/>"),
            0x02 => out.push_str(r#"<u val="double"/>"#),
            0x21 => out.push_str(r#"<u val="singleAccounting"/>"#),
            0x22 => out.push_str(r#"<u val="doubleAccounting"/>"#),
            _ => {}
        }
        match font.escapement {
            1 => out.push_str(r#"<vertAlign val="superscript"/>"#),
            2 => out.push_str(r#"<vertAlign val="subscript"/>"#),
            _ => {}
        }
        let points = f64::from(if font.twips == 0 { 220 } else { font.twips }) / 20.0;
        out.push_str(&format!(r#"<sz val="{points}"/>"#));
        out.push_str(&self.color_tag(font.color, color));
        let name = if font.name.is_empty() {
            "Calibri"
        } else {
            &font.name
        };
        out.push_str(&format!(r#"<name val="{}"/>"#, escape_xml(name)));
        out.push_str("</font>");
        out
    }

    fn fill_xml(&self, xf: &XlsXf, ext: Option<&XlsXfExt>) -> String {
        let Some(pattern) = fill_pattern(xf.pattern) else {
            return r#"<fill><patternFill patternType="none"/></fill>"#.to_string();
        };
        let fg = self
            .color_rgb(xf.fore, ext.and_then(|e| e.fore))
            .map(|rgb| format!(r#"<fgColor rgb="{rgb}"/>"#))
            .unwrap_or_default();
        // a solid fill shows only its foreground; the background a writer
        // left behind would only confuse readers that look at both
        let bg = if pattern == "solid" {
            String::new()
        } else {
            self.color_rgb(xf.back, ext.and_then(|e| e.back))
                .map(|rgb| format!(r#"<bgColor rgb="{rgb}"/>"#))
                .unwrap_or_default()
        };
        format!(r#"<fill><patternFill patternType="{pattern}">{fg}{bg}</patternFill></fill>"#)
    }

    fn border_xml(&self, xf: &XlsXf, ext: Option<&XlsXfExt>) -> String {
        let side = |tag: &str, index: usize| {
            let (style, icv) = xf.borders[index];
            match border_style(style) {
                Some(style) => format!(
                    r#"<{tag} style="{style}">{}</{tag}>"#,
                    self.color_tag(icv, ext.and_then(|e| e.borders[index]))
                ),
                None => format!("<{tag}/>"),
            }
        };
        let mut attrs = String::new();
        if xf.diag_up {
            attrs.push_str(r#" diagonalUp="1""#);
        }
        if xf.diag_down {
            attrs.push_str(r#" diagonalDown="1""#);
        }
        format!(
            "<border{attrs}>{}{}{}{}{}</border>",
            side("left", 0),
            side("right", 1),
            side("top", 2),
            side("bottom", 3),
            side("diagonal", 4),
        )
    }

    fn alignment_xml(xf: &XlsXf) -> String {
        let mut attrs = String::new();
        if let Some(h) = h_align(xf.h_align) {
            attrs.push_str(&format!(r#" horizontal="{h}""#));
        }
        // BIFF's default vertical alignment is bottom, which is also OOXML's
        if let Some(v) = v_align(xf.v_align) {
            attrs.push_str(&format!(r#" vertical="{v}""#));
        }
        if xf.rotation != 0 {
            attrs.push_str(&format!(r#" textRotation="{}""#, xf.rotation));
        }
        if xf.wrap {
            attrs.push_str(r#" wrapText="1""#);
        }
        if xf.indent != 0 {
            attrs.push_str(&format!(r#" indent="{}""#, xf.indent));
        }
        if xf.shrink {
            attrs.push_str(r#" shrinkToFit="1""#);
        }
        if attrs.is_empty() {
            String::new()
        } else {
            format!("<alignment{attrs}/>")
        }
    }

    /// The font list position for a BIFF font index: index 4 does not exist
    /// in the file, so every index past it sits one slot lower.
    fn font_id(&self, ifnt: u16) -> usize {
        let index = if ifnt >= 4 { ifnt - 1 } else { ifnt };
        usize::from(index).min(self.fonts.len().saturating_sub(1))
    }

    fn num_fmt_id(&self, ifmt: u16) -> u16 {
        if ifmt < 164 || self.formats.contains_key(&ifmt) {
            ifmt
        } else {
            0
        }
    }

    pub(crate) fn styles_xml(&self) -> String {
        let mut fonts: Vec<String> = if self.fonts.is_empty() {
            vec![r#"<font><sz val="11"/><name val="Calibri"/></font>"#.to_string()]
        } else {
            self.fonts.iter().map(|f| self.font_xml(f, None)).collect()
        };
        let mut fills = vec![
            r#"<fill><patternFill patternType="none"/></fill>"#.to_string(),
            r#"<fill><patternFill patternType="gray125"/></fill>"#.to_string(),
        ];
        let mut borders =
            vec!["<border><left/><right/><top/><bottom/><diagonal/></border>".to_string()];
        let intern = |list: &mut Vec<String>, xml: String| {
            if let Some(index) = list.iter().position(|x| *x == xml) {
                return index;
            }
            list.push(xml);
            list.len() - 1
        };

        let mut xfs = vec![
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#.to_string(),
            r#"<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#.to_string(),
            r#"<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#.to_string(),
        ];
        for (index, xf) in self.xfs.iter().enumerate() {
            let ext = self.ext_of(index);
            let fill_id = intern(&mut fills, self.fill_xml(xf, ext));
            let border_id = intern(&mut borders, self.border_xml(xf, ext));
            let mut font_id = self.font_id(xf.font);
            // the text color belongs to the XF, not to the FONT other XFs share
            if let (Some(color), Some(font)) = (ext.and_then(|e| e.text), self.fonts.get(font_id)) {
                font_id = intern(&mut fonts, self.font_xml(font, Some(color)));
            }
            let alignment = Self::alignment_xml(xf);
            let protection = match (xf.locked, xf.hidden) {
                (true, false) => String::new(),
                (locked, hidden) => format!(
                    r#"<protection locked="{}" hidden="{}"/>"#,
                    u8::from(locked),
                    u8::from(hidden)
                ),
            };
            let mut attrs = format!(
                r#"numFmtId="{}" fontId="{font_id}" fillId="{fill_id}" borderId="{border_id}" xfId="0""#,
                self.num_fmt_id(xf.format),
            );
            if xf.format != 0 {
                attrs.push_str(r#" applyNumberFormat="1""#);
            }
            attrs.push_str(r#" applyFont="1""#);
            if fill_id != 0 {
                attrs.push_str(r#" applyFill="1""#);
            }
            if border_id != 0 {
                attrs.push_str(r#" applyBorder="1""#);
            }
            if !alignment.is_empty() {
                attrs.push_str(r#" applyAlignment="1""#);
            }
            if !protection.is_empty() {
                attrs.push_str(r#" applyProtection="1""#);
            }
            xfs.push(if alignment.is_empty() && protection.is_empty() {
                format!("<xf {attrs}/>")
            } else {
                format!("<xf {attrs}>{alignment}{protection}</xf>")
            });
        }

        let mut format_ids: Vec<&u16> = self.formats.keys().filter(|id| **id >= 164).collect();
        format_ids.sort_unstable();
        let num_fmts = if format_ids.is_empty() {
            String::new()
        } else {
            let entries: String = format_ids
                .iter()
                .map(|id| {
                    format!(
                        r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                        escape_xml(&self.formats[id])
                    )
                })
                .collect();
            format!(
                r#"<numFmts count="{}">{entries}</numFmts>"#,
                format_ids.len()
            )
        };

        format!(
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                "\n",
                r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
                "{num_fmts}",
                r#"<fonts count="{}">{}</fonts>"#,
                r#"<fills count="{}">{}</fills>"#,
                r#"<borders count="{}">{}</borders>"#,
                r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                r#"<cellXfs count="{}">{}</cellXfs>"#,
                r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
                "</styleSheet>"
            ),
            fonts.len(),
            fonts.concat(),
            fills.len(),
            fills.concat(),
            borders.len(),
            borders.concat(),
            xfs.len(),
            xfs.concat(),
            num_fmts = num_fmts,
        )
    }
}

/// An XLUnicodeString whose character count is one byte (FONT names).
fn short_unicode_string(bytes: &[u8]) -> String {
    match bytes.first() {
        Some(&length) => unicode_string(&bytes[1..], usize::from(length)),
        None => String::new(),
    }
}

/// The character data of an XLUnicodeString after its count: a flags byte,
/// then either compressed (one byte per char, Latin-1) or UTF-16LE text.
/// Rich-text and phonetic extras announced in the flags are skipped over.
fn unicode_string(bytes: &[u8], length: usize) -> String {
    let Some(&flags) = bytes.first() else {
        return String::new();
    };
    let mut at = 1usize;
    if flags & 0x08 != 0 {
        at += 2;
    }
    if flags & 0x04 != 0 {
        at += 4;
    }
    if flags & 0x01 == 0 {
        bytes
            .get(at..at + length)
            .map(|chars| chars.iter().map(|&b| char::from(b)).collect())
            .unwrap_or_default()
    } else {
        let units: Vec<u16> = bytes
            .get(at..at + length * 2)
            .map(|chars| {
                chars
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect()
            })
            .unwrap_or_default();
        String::from_utf16_lossy(&units)
    }
}

fn border_style(dg: u8) -> Option<&'static str> {
    Some(match dg {
        1 => "thin",
        2 => "medium",
        3 => "dashed",
        4 => "dotted",
        5 => "thick",
        6 => "double",
        7 => "hair",
        8 => "mediumDashed",
        9 => "dashDot",
        10 => "mediumDashDot",
        11 => "dashDotDot",
        12 => "mediumDashDotDot",
        13 => "slantDashDot",
        _ => return None,
    })
}

fn fill_pattern(fls: u8) -> Option<&'static str> {
    Some(match fls {
        1 => "solid",
        2 => "mediumGray",
        3 => "darkGray",
        4 => "lightGray",
        5 => "darkHorizontal",
        6 => "darkVertical",
        7 => "darkDown",
        8 => "darkUp",
        9 => "darkGrid",
        10 => "darkTrellis",
        11 => "lightHorizontal",
        12 => "lightVertical",
        13 => "lightDown",
        14 => "lightUp",
        15 => "lightGrid",
        16 => "lightTrellis",
        17 => "gray125",
        18 => "gray0625",
        _ => return None,
    })
}

fn h_align(alc: u8) -> Option<&'static str> {
    Some(match alc {
        1 => "left",
        2 => "center",
        3 => "right",
        4 => "fill",
        5 => "justify",
        6 => "centerContinuous",
        7 => "distributed",
        _ => return None,
    })
}

fn v_align(alc: u8) -> Option<&'static str> {
    Some(match alc {
        0 => "top",
        1 => "center",
        3 => "justify",
        4 => "distributed",
        _ => return None,
    })
}

const THEME_MAX_ENTRIES: usize = 64;
const THEME_MAX_BYTES: u64 = 4 * 1024 * 1024;

/// clrScheme children in `theme` attribute order: the light/dark pairs swap.
/// Excel writes XFEXT theme indices in this same order (0 = lt1, 1 = dk1),
/// whatever MS-XLS's ColorTheme table says: default black text is index 1.
const THEME_SLOTS: [&str; 12] = [
    "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// The Office 2007 theme, which a THEME record without a package stands for.
const DEFAULT_THEME: [[u8; 3]; 12] = [
    [0xFF, 0xFF, 0xFF],
    [0x00, 0x00, 0x00],
    [0xEE, 0xEC, 0xE1],
    [0x1F, 0x49, 0x7D],
    [0x4F, 0x81, 0xBD],
    [0xC0, 0x50, 0x4D],
    [0x9B, 0xBB, 0x59],
    [0x80, 0x64, 0xA2],
    [0x4B, 0xAC, 0xC6],
    [0xF7, 0x96, 0x46],
    [0x00, 0x00, 0xFF],
    [0x80, 0x00, 0x80],
];

/// `<a:lt1><a:sysClr … lastClr="FFFFFF"/></a:lt1>` or `<a:srgbClr val="…"/>` → rgb
fn theme_slot_rgb(xml: &str, slot: &str) -> Option<[u8; 3]> {
    let open = xml.find(&format!("<a:{slot}>"))?;
    let body = &xml[open..open + xml[open..].find(&format!("</a:{slot}>"))?];
    let attr = if body.contains("<a:srgbClr") {
        "val=\""
    } else {
        "lastClr=\""
    };
    let at = body.find(attr)? + attr.len();
    let hex = body.get(at..at + 6)?;
    let channel = |i: usize| u8::from_str_radix(hex.get(i..i + 2)?, 16).ok();
    Some([channel(0)?, channel(2)?, channel(4)?])
}

/// The 56 colors Excel ships at palette indices 8–63 (MS-XLS 2.5.161).
const DEFAULT_PALETTE: [[u8; 3]; 56] = [
    [0x00, 0x00, 0x00],
    [0xFF, 0xFF, 0xFF],
    [0xFF, 0x00, 0x00],
    [0x00, 0xFF, 0x00],
    [0x00, 0x00, 0xFF],
    [0xFF, 0xFF, 0x00],
    [0xFF, 0x00, 0xFF],
    [0x00, 0xFF, 0xFF],
    [0x80, 0x00, 0x00],
    [0x00, 0x80, 0x00],
    [0x00, 0x00, 0x80],
    [0x80, 0x80, 0x00],
    [0x80, 0x00, 0x80],
    [0x00, 0x80, 0x80],
    [0xC0, 0xC0, 0xC0],
    [0x80, 0x80, 0x80],
    [0x99, 0x99, 0xFF],
    [0x99, 0x33, 0x66],
    [0xFF, 0xFF, 0xCC],
    [0xCC, 0xFF, 0xFF],
    [0x66, 0x00, 0x66],
    [0xFF, 0x80, 0x80],
    [0x00, 0x66, 0xCC],
    [0xCC, 0xCC, 0xFF],
    [0x00, 0x00, 0x80],
    [0xFF, 0x00, 0xFF],
    [0xFF, 0xFF, 0x00],
    [0x00, 0xFF, 0xFF],
    [0x80, 0x00, 0x80],
    [0x80, 0x00, 0x00],
    [0x00, 0x80, 0x80],
    [0x00, 0x00, 0xFF],
    [0x00, 0xCC, 0xFF],
    [0xCC, 0xFF, 0xFF],
    [0xCC, 0xFF, 0xCC],
    [0xFF, 0xFF, 0x99],
    [0x99, 0xCC, 0xFF],
    [0xFF, 0x99, 0xCC],
    [0xCC, 0x99, 0xFF],
    [0xFF, 0xCC, 0x99],
    [0x33, 0x66, 0xFF],
    [0x33, 0xCC, 0xCC],
    [0x99, 0xCC, 0x00],
    [0xFF, 0xCC, 0x00],
    [0xFF, 0x99, 0x00],
    [0xFF, 0x66, 0x00],
    [0x66, 0x66, 0x99],
    [0x96, 0x96, 0x96],
    [0x00, 0x33, 0x66],
    [0x33, 0x99, 0x66],
    [0x00, 0x33, 0x00],
    [0x33, 0x33, 0x00],
    [0x99, 0x33, 0x00],
    [0x99, 0x33, 0x66],
    [0x33, 0x33, 0x99],
    [0x33, 0x33, 0x33],
];

#[cfg(test)]
mod tests {
    use super::*;

    /// A solid-fill XF on font 0 whose palette indices are the stand-ins Excel leaves behind.
    fn xf(fore_icv: u16) -> Vec<u8> {
        let mut bytes = vec![0u8; 20];
        bytes[14..18].copy_from_slice(&(1u32 << 26).to_le_bytes());
        bytes[18..20].copy_from_slice(&fore_icv.to_le_bytes());
        bytes
    }

    fn font(color_icv: u16) -> Vec<u8> {
        let mut bytes = vec![0u8; 14];
        bytes[0..2].copy_from_slice(&220u16.to_le_bytes());
        bytes[4..6].copy_from_slice(&color_icv.to_le_bytes());
        bytes.extend_from_slice(&[5, 0]);
        bytes.extend_from_slice(b"Arial");
        bytes
    }

    fn ext_prop(kind: u16, color_type: u16, tint: i16, value: [u8; 4]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&20u16.to_le_bytes());
        bytes.extend_from_slice(&color_type.to_le_bytes());
        bytes.extend_from_slice(&tint.to_le_bytes());
        bytes.extend_from_slice(&value);
        bytes.extend_from_slice(&[0; 8]);
        bytes
    }

    fn xf_ext(ixfe: u16, props: &[Vec<u8>]) -> Vec<u8> {
        let mut bytes = vec![0u8; 14];
        bytes.extend_from_slice(&ixfe.to_le_bytes());
        bytes.extend_from_slice(&[0, 0]);
        bytes.extend_from_slice(&(props.len() as u16).to_le_bytes());
        bytes.extend(props.iter().flatten());
        bytes
    }

    fn xf_crc(count: u16) -> Vec<u8> {
        let mut bytes = vec![0u8; 14];
        bytes.extend_from_slice(&count.to_le_bytes());
        bytes.extend_from_slice(&[0; 4]);
        bytes
    }

    #[test]
    fn an_xf_extension_color_replaces_the_palette_stand_in() {
        let mut styles = XlsStyles::default();
        styles.push_font(&font(8));
        // icv 62 (333399) is where Excel parks a navy the palette lacks
        styles.push_xf(&xf(62));
        styles.push_xf(&xf(62));
        styles.push_xf(&xf(62));
        styles.set_xf_crc(&xf_crc(3));
        styles.push_xf_ext(&xf_ext(
            0,
            &[
                ext_prop(0x0004, 2, 0, [0x1F, 0x3A, 0x5F, 0xFF]),
                ext_prop(0x000D, 2, 0, [0xE0, 0x7B, 0x39, 0xFF]),
            ],
        ));
        // accent 1 of the default theme, 40% lighter
        styles.push_xf_ext(&xf_ext(1, &[ext_prop(0x0004, 3, 13106, [4, 0, 0, 0])]));

        let xml = styles.styles_xml();
        assert!(xml.contains(r#"<fgColor rgb="FF1F3A5F"/>"#));
        assert!(xml.contains(r#"<fgColor rgb="FF95B3D7"/>"#));
        // the XF without an extension keeps the palette color
        assert!(xml.contains(r#"<fgColor rgb="FF333399"/>"#));
        // the text color lands on a font of its own; the shared FONT keeps its color
        assert!(xml.contains(r#"<fonts count="2">"#));
        assert!(xml.contains(r#"<color rgb="FFE07B39"/><name val="Arial"/>"#));
        assert!(xml.contains(r#"<color rgb="FF000000"/><name val="Arial"/>"#));
        assert!(xml.contains(r#"fontId="1" fillId="2""#));
    }

    #[test]
    fn theme_indices_follow_excel_light_before_dark() {
        let mut styles = XlsStyles::default();
        styles.push_font(&font(8));
        for index in 0..4u8 {
            styles.push_xf(&xf(62));
            styles.push_xf_ext(&xf_ext(
                u16::from(index),
                &[ext_prop(0x0004, 3, 0, [index, 0, 0, 0])],
            ));
        }
        let xml = styles.styles_xml();
        let fills: Vec<&str> = xml
            .split(r#"<fgColor rgb=""#)
            .skip(1)
            .map(|f| &f[..8])
            .collect();
        assert_eq!(fills, ["FFFFFFFF", "FF000000", "FFEEECE1", "FF1F497D"]);
    }

    #[test]
    fn an_oversized_theme_entry_is_ignored() {
        let mut package = std::io::Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut package);
            writer
                .start_file(
                    "theme/theme/theme1.xml",
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            let chunk = vec![b' '; 1024 * 1024];
            for _ in 0..5 {
                std::io::Write::write_all(&mut writer, &chunk).unwrap();
            }
            writer.finish().unwrap();
        }
        let mut record = vec![0u8; 16];
        record.extend_from_slice(package.get_ref());
        let mut styles = XlsStyles::default();
        styles.set_theme(&record);
        assert!(styles.theme.is_none());
    }

    #[test]
    fn extensions_written_for_another_xf_table_are_ignored() {
        let mut styles = XlsStyles::default();
        styles.push_font(&font(8));
        styles.push_xf(&xf(62));
        styles.set_xf_crc(&xf_crc(5));
        styles.push_xf_ext(&xf_ext(
            0,
            &[ext_prop(0x0004, 2, 0, [0x1F, 0x3A, 0x5F, 0xFF])],
        ));
        let xml = styles.styles_xml();
        assert!(xml.contains(r#"<fgColor rgb="FF333399"/>"#));
        assert!(!xml.contains("1F3A5F"));
    }

    #[test]
    fn a_theme_package_supplies_the_theme_colors() {
        let xml = concat!(
            "<a:clrScheme>",
            r#"<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>"#,
            r#"<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>"#,
            r#"<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>"#,
            r#"<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>"#,
            r#"<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>"#,
            r#"<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>"#,
            r#"<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>"#,
            "</a:clrScheme>"
        );
        let mut package = std::io::Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut package);
            writer
                .start_file(
                    "theme/theme/theme1.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            std::io::Write::write_all(&mut writer, xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        let mut record = vec![0u8; 16];
        record.extend_from_slice(package.get_ref());

        let mut styles = XlsStyles::default();
        styles.push_font(&font(8));
        styles.push_xf(&xf(62));
        styles.push_xf_ext(&xf_ext(0, &[ext_prop(0x0004, 3, 0, [5, 0, 0, 0])]));
        styles.set_theme(&record);
        assert!(styles.styles_xml().contains(r#"<fgColor rgb="FFED7D31"/>"#));
    }
}

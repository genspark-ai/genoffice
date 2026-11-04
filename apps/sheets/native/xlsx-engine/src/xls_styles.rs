//! Cell formatting of a BIFF8 workbook — FONT, FORMAT, XF and PALETTE records
//! from the globals substream — rebuilt as an OOXML `styles.xml`. Each XF
//! record becomes one `cellXfs` entry at the same index, so a cell's `ixfe`
//! maps to `s` by a fixed offset and nothing is deduplicated or reordered.

use std::collections::HashMap;

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

#[derive(Default)]
pub(crate) struct XlsStyles {
    pub(crate) fonts: Vec<XlsFont>,
    pub(crate) formats: HashMap<u16, String>,
    pub(crate) xfs: Vec<XlsXf>,
    pub(crate) palette: Option<Vec<[u8; 3]>>,
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

    fn color_tag(&self, icv: u16) -> String {
        match self.rgb(icv) {
            Some(rgb) => format!(r#"<color rgb="{rgb}"/>"#),
            None => String::new(),
        }
    }

    fn font_xml(&self, font: &XlsFont) -> String {
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
        out.push_str(&self.color_tag(font.color));
        let name = if font.name.is_empty() {
            "Calibri"
        } else {
            &font.name
        };
        out.push_str(&format!(r#"<name val="{}"/>"#, escape_xml(name)));
        out.push_str("</font>");
        out
    }

    fn fill_xml(&self, xf: &XlsXf) -> String {
        let Some(pattern) = fill_pattern(xf.pattern) else {
            return r#"<fill><patternFill patternType="none"/></fill>"#.to_string();
        };
        let fg = self
            .rgb(xf.fore)
            .map(|rgb| format!(r#"<fgColor rgb="{rgb}"/>"#))
            .unwrap_or_default();
        // a solid fill shows only its foreground; the background a writer
        // left behind would only confuse readers that look at both
        let bg = if pattern == "solid" {
            String::new()
        } else {
            self.rgb(xf.back)
                .map(|rgb| format!(r#"<bgColor rgb="{rgb}"/>"#))
                .unwrap_or_default()
        };
        format!(r#"<fill><patternFill patternType="{pattern}">{fg}{bg}</patternFill></fill>"#)
    }

    fn border_xml(&self, xf: &XlsXf) -> String {
        let side = |tag: &str, (style, icv): (u8, u16)| match border_style(style) {
            Some(style) => format!(r#"<{tag} style="{style}">{}</{tag}>"#, self.color_tag(icv)),
            None => format!("<{tag}/>"),
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
            side("left", xf.borders[0]),
            side("right", xf.borders[1]),
            side("top", xf.borders[2]),
            side("bottom", xf.borders[3]),
            side("diagonal", xf.borders[4]),
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
        let fonts: Vec<String> = if self.fonts.is_empty() {
            vec![r#"<font><sz val="11"/><name val="Calibri"/></font>"#.to_string()]
        } else {
            self.fonts.iter().map(|f| self.font_xml(f)).collect()
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
        for xf in &self.xfs {
            let fill_id = intern(&mut fills, self.fill_xml(xf));
            let border_id = intern(&mut borders, self.border_xml(xf));
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
                r#"numFmtId="{}" fontId="{}" fillId="{fill_id}" borderId="{border_id}" xfId="0""#,
                self.num_fmt_id(xf.format),
                self.font_id(xf.font),
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

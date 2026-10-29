//! The Escher drawing format .xls embeds inside BIFF records (MS-ODRAW): the
//! container tree, the pictures of the workbook drawing group, and the
//! per-sheet picture anchors. Only what the conversion needs is parsed —
//! containers are walked recursively, leaf records are opaque bytes.

/// One embedded picture from the drawing group's BStore container.
pub(crate) struct Blip {
    pub(crate) kind: BlipKind,
    /// The image file bytes, ready for an OOXML media part.
    pub(crate) bytes: Vec<u8>,
}

pub(crate) enum BlipKind {
    Png,
    Jpeg,
    /// WMF/EMF/DIB and friends: the conversion skips these for now.
    Other,
}

impl BlipKind {
    /// The media extension the converted package writes for this kind.
    pub(crate) fn extension(&self) -> &'static str {
        match self {
            BlipKind::Png => "png",
            BlipKind::Jpeg => "jpg",
            BlipKind::Other => "bin",
        }
    }
}

/// One picture shape on a sheet.
pub(crate) struct Picture {
    /// `[colL, dxL, rowT, dyT, colR, dxR, rowB, dyB]`, all zero-based;
    /// the dx/dy offsets are 1/1024 of the column width / row height, and
    /// 0xFFFF in colR/rowB means the object runs to the sheet's edge.
    pub(crate) anchor: [u16; 8],
    /// 1-based index into the drawing group's BSE list.
    pub(crate) blip: u32,
}

/// One record in an Escher stream. Container records carry their children in
/// `data`; leaf records carry opaque bytes.
struct Node<'a> {
    ver_inst: u16,
    kind: u16,
    data: &'a [u8],
}

fn u16le(bytes: &[u8], at: usize) -> u16 {
    u16::from_le_bytes([bytes[at], bytes[at + 1]])
}

fn u32le(bytes: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]])
}

/// Parses one level of the Escher tree. A stream fragment that ends mid-record
/// ends the level; everything parsed so far stays usable.
fn children(data: &[u8]) -> Vec<Node<'_>> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 8 <= data.len() {
        let ver_inst = u16le(data, pos);
        let kind = u16le(data, pos + 2);
        let len = u32le(data, pos + 4) as usize;
        if kind == 0 {
            break;
        }
        let end = (pos + 8 + len).min(data.len());
        out.push(Node {
            ver_inst,
            kind,
            data: &data[pos + 8..end],
        });
        pos += 8 + len;
    }
    out
}

fn is_container(ver_inst: u16) -> bool {
    ver_inst & 0xF == 0xF
}

/// Container and record kinds this parser cares about (MS-ODRAW).
const BSTORE_CONTAINER: u16 = 0xF001;
const SP_CONTAINER: u16 = 0xF004;
const BSE: u16 = 0xF007;
const OPT: u16 = 0xF00B;
const CLIENT_ANCHOR: u16 = 0xF010;
const BLIP_PNG: u16 = 0xF01E;
const BLIP_JPEG: u16 = 0xF01D;
/// OPT property `blipPib`: which BSE the shape's picture comes from.
const PROP_BLIP_PIB: u16 = 0x0104;

/// Extracts the pictures of the drawing group: BStoreContainer → BSE records,
/// each carrying one image after a fixed header and an optional name.
pub(crate) fn parse_blips(dgg: &[u8]) -> Vec<Blip> {
    let mut out = Vec::new();
    for store in containers_of(dgg, BSTORE_CONTAINER) {
        for entry in children(store.data) {
            if entry.kind != BSE || entry.data.len() < 36 {
                continue;
            }
            let bytes = entry.data;
            // BSE: btWin32, btMacOS, rgbUid(16), tag(2), cRef(4), foDelay(4),
            // cUsage, cbName, two unused bytes — then the named image.
            let name_len = bytes[29] as usize;
            let Some(blip) = bytes.get(36 + name_len..) else {
                continue;
            };
            if blip.len() < 8 {
                continue;
            }
            let kind = u16le(blip, 2);
            let kind = match kind {
                BLIP_PNG => BlipKind::Png,
                BLIP_JPEG => BlipKind::Jpeg,
                _ => BlipKind::Other,
            };
            // after the blip's own 8-byte record header: rgbUid(16) — doubled
            // when the instance's low bit is set — then a 0xFF tag byte, then
            // the image bytes.
            let skip = 8 + 16 + if u16le(blip, 0) & 0x1 == 1 { 16 } else { 0 } + 1;
            let image = blip.get(skip..).unwrap_or(&[]).to_vec();
            out.push(Blip { kind, bytes: image });
        }
    }
    out
}

/// Every container of `kind` anywhere in the tree, outermost first.
fn containers_of<'a>(data: &'a [u8], kind: u16) -> Vec<Node<'a>> {
    let mut out = Vec::new();
    for node in children(data) {
        if node.kind == kind {
            out.push(node);
            continue;
        }
        if is_container(node.ver_inst) {
            out.extend(containers_of(node.data, kind));
        }
    }
    out
}

/// Collects the picture anchors of one sheet's drawing: every shape container
/// that pairs a client anchor with a blip reference.
pub(crate) fn parse_pictures(drawing: &[u8]) -> Vec<Picture> {
    let mut out = Vec::new();
    for node in children(drawing) {
        visit(&node, &mut out);
    }
    out
}

fn visit(node: &Node<'_>, out: &mut Vec<Picture>) {
    if node.kind == SP_CONTAINER {
        let mut anchor: Option<[u16; 8]> = None;
        let mut blip: Option<u32> = None;
        for child in children(node.data) {
            match child.kind {
                CLIENT_ANCHOR if child.data.len() >= 18 => {
                    // flags(2), then the eight anchor fields
                    let bytes = child.data;
                    anchor = Some([
                        u16le(bytes, 2),
                        u16le(bytes, 4),
                        u16le(bytes, 6),
                        u16le(bytes, 8),
                        u16le(bytes, 10),
                        u16le(bytes, 12),
                        u16le(bytes, 14),
                        u16le(bytes, 16),
                    ]);
                }
                OPT => blip = blip_pib(child.data),
                _ => {}
            }
        }
        if let (Some(anchor), Some(blip)) = (anchor, blip) {
            out.push(Picture { anchor, blip });
        }
        return;
    }
    if is_container(node.ver_inst) {
        for child in children(node.data) {
            visit(&child, out);
        }
    }
}

/// Reads the `blipPib` property out of an OPT record: a run of 6-byte
/// properties written in ascending id order. Excel and LibreOffice both store
/// the picture reference as a plain value ahead of the complex geometry
/// properties, so the run is read front to back and a complex pib — which
/// would point through the pool stream — is skipped for now.
fn blip_pib(data: &[u8]) -> Option<u32> {
    let mut at = 0usize;
    while at + 6 <= data.len() {
        let attr = u16le(data, at);
        if attr & 0x3FFF == PROP_BLIP_PIB {
            return if attr & 0x8000 == 0 {
                Some(u32le(data, at + 2))
            } else {
                None
            };
        }
        at += 6;
    }
    None
}

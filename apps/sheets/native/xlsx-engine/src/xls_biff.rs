//! The binary layer of an .xls workbook (BIFF8 inside an OLE2 container),
//! walked record by record for the parts calamine does not expose: the
//! workbook drawing group (embedded pictures), per-sheet drawing shapes,
//! column widths (COLINFO) and hand-set row heights (ROW). Everything here is
//! best-effort — a stream that refuses to parse yields empty collections, and
//! the conversion keeps the cell data calamine already read.

use std::fs::File;
use std::io::Read;
use std::path::Path;

use crate::escher::{self, Blip};

/// Per-sheet binary extras calamine does not surface.
pub(crate) struct XlsSheetBinary {
    /// (first column, last column, width in 1/256 character units), 0-based.
    pub(crate) col_infos: Vec<(u16, u16, u16)>,
    /// (row, height in twips, set by hand) for every row the file describes;
    /// the flag separates heights someone chose from the sheet's default
    /// echoed back, and only the former are worth writing to the conversion.
    pub(crate) row_heights: Vec<(u16, u16, bool)>,
    /// The sheet's MSODRAWING record with its continues merged into one
    /// Escher stream, when the sheet carries drawings.
    pub(crate) drawing: Option<Vec<u8>>,
}

pub(crate) struct XlsBinary {
    /// Pictures embedded in the workbook drawing group, in BSE order.
    pub(crate) blips: Vec<Blip>,
    /// One entry per BOUNDSHEET, in sheet order.
    pub(crate) sheets: Vec<XlsSheetBinary>,
}

/// Reads the .xls binary layer. `None` when the file is not an OLE2 workbook
/// (ODF and XML spreadsheets take the same converter) or carries no Workbook
/// stream.
pub(crate) fn read_xls_binary(source: &Path) -> Option<XlsBinary> {
    let mut container = cfb::CompoundFile::open(File::open(source).ok()?).ok()?;
    let mut stream = ["/Workbook", "/Book"]
        .into_iter()
        .find_map(|name| container.open_stream(name).ok())?;
    let mut data = Vec::new();
    stream.read_to_end(&mut data).ok()?;
    if data.is_empty() {
        return None;
    }
    Some(walk(&data))
}

/// One BIFF record: file position of the header, id, payload length.
struct Record {
    at: usize,
    id: u16,
    len: usize,
}

fn scan(data: &[u8]) -> Vec<Record> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= data.len() {
        let id = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let len = u16::from_le_bytes([data[pos + 2], data[pos + 3]]) as usize;
        let at = pos;
        pos += 4 + len;
        // A truncated tail (or a mis-parse) ends the walk: records are only
        // meaningful while the stream stays self-consistent.
        if pos > data.len() {
            break;
        }
        out.push(Record { at, id, len });
    }
    out
}

fn payload<'a>(data: &'a [u8], record: &Record) -> &'a [u8] {
    &data[record.at + 4..record.at + 4 + record.len]
}

/// A record's payload with consecutive Continue records appended. Drawing
/// records fragment their Escher stream across continues; the concatenation
/// is the stream the Escher parser expects.
fn payload_with_continues(data: &[u8], records: &[Record], index: usize) -> Vec<u8> {
    let mut out = payload(data, &records[index]).to_vec();
    for record in records.iter().skip(index + 1) {
        if record.id != 0x003C {
            break;
        }
        out.extend_from_slice(payload(data, record));
    }
    out
}

fn walk(data: &[u8]) -> XlsBinary {
    let records = scan(data);
    // BOUNDSHEET entries carry each sheet's BOF position, in sheet order —
    // that is how records between one BOF and the next are attributed.
    let mut sheet_bofs: Vec<u32> = Vec::new();
    for record in &records {
        if record.id == 0x0085 && record.len >= 8 {
            let bytes = payload(data, record);
            sheet_bofs.push(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]));
        }
    }
    let mut sheets: Vec<XlsSheetBinary> = sheet_bofs
        .iter()
        .map(|_| XlsSheetBinary {
            col_infos: Vec::new(),
            row_heights: Vec::new(),
            drawing: None,
        })
        .collect();

    let mut dgg: Vec<u8> = Vec::new();
    let mut current: Option<usize> = None;
    let mut index = 0usize;
    while index < records.len() {
        let record = &records[index];
        if record.id == 0x0809 {
            // A substream starts. A position that matches a BOUNDSHEET entry
            // opens that sheet; the globals and chart substreams (whose BOFs
            // no BOUNDSHEET points at) detach collection so chart drawings
            // never land on the preceding sheet.
            current = sheet_bofs.iter().position(|&bof| bof as usize == record.at);
            index += 1;
            continue;
        }
        match current {
            None => {
                if record.id == 0x00EB {
                    dgg.extend_from_slice(&payload_with_continues(data, &records, index));
                }
            }
            Some(sheet) => match record.id {
                0x00EC => {
                    let bytes = payload_with_continues(data, &records, index);
                    let drawing = sheets[sheet].drawing.get_or_insert_with(Vec::new);
                    drawing.extend_from_slice(&bytes);
                }
                0x007D if record.len >= 6 => {
                    let bytes = payload(data, record);
                    let first = u16::from_le_bytes([bytes[0], bytes[1]]);
                    let last = u16::from_le_bytes([bytes[2], bytes[3]]);
                    let width = u16::from_le_bytes([bytes[4], bytes[5]]);
                    sheets[sheet].col_infos.push((first, last, width));
                }
                0x0208 if record.len >= 14 => {
                    let bytes = payload(data, record);
                    let row = u16::from_le_bytes([bytes[0], bytes[1]]);
                    let twips = u16::from_le_bytes([bytes[6], bytes[7]]);
                    let grbit = u16::from_le_bytes([bytes[12], bytes[13]]);
                    // bit 0x40 marks a height that is not the sheet default
                    if twips > 0 {
                        sheets[sheet]
                            .row_heights
                            .push((row, twips, grbit & 0x0040 != 0));
                    }
                }
                _ => {}
            },
        }
        index += 1;
    }

    let blips = escher::parse_blips(&dgg);
    XlsBinary { blips, sheets }
}

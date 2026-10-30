//! Streaming cell upserts for worksheet parts above the host's patch cap.
//!
//! The TypeScript gateway cannot hold a part larger than ~500 MiB as one
//! string (V8's maximum string length), so for those parts it ships a
//! per-cell patch program instead: the final serialized `<c>` elements,
//! sorted by row then column. This module splices them into the entry while
//! it streams from the source archive to the target: bytes outside the
//! touched rows are copied verbatim (the byte-faithful save guarantee),
//! memory stays bounded by one row plus one IO chunk, and the OOXML
//! semantics (styles, shared strings, formulas) remain the gateway's
//! business — the sidecar only finds the splice points.
//!
//! Scanning by literal tag bytes is sound for well-formed XML: `<` cannot
//! appear raw in text content or attribute values, so `<row`, `</row>`,
//! `<c` and `</c>` never occur inside another element's data. Parts that
//! bind SpreadsheetML to a prefix (Open XML SDK's `<x:sheetData>`) are
//! handled by reading the prefix off `sheetData` and applying it to every
//! tag we match or emit, including the gateway's unprefixed cell XML.

use std::io::{Read, Write};
use std::path::Path;

use serde::Deserialize;

use crate::SidecarError;

/// One serialized `<c>` element, upserted at a 1-based column.
#[derive(Debug, Deserialize)]
pub struct CellUpsert {
    pub col: u32,
    pub xml: String,
}

/// Upserts for one 1-based row, columns strictly ascending.
#[derive(Debug, Deserialize)]
pub struct RowPatch {
    pub row: u32,
    pub cells: Vec<CellUpsert>,
}

/// A whole part's patch program, rows strictly ascending.
#[derive(Debug, Deserialize)]
pub struct StreamPatchProgram {
    pub rows: Vec<RowPatch>,
}

/// A buffered existing row is capped: real-world rows are < 16384 cells. A
/// "row" needing more than this is a malformed or hostile part.
const MAX_ROW_BYTES: usize = 64 * 1024 * 1024;
/// One serialized cell element from the gateway.
const MAX_CELL_XML_BYTES: usize = 4 * 1024 * 1024;
/// Patch program totals (counts, not bytes): far above any interactive edit
/// burst, far below anything that should have been a full-part rewrite.
const MAX_PATCHED_CELLS: usize = 4_000_000;
const CHUNK: usize = 256 * 1024;

pub fn read_patch_program(path: &Path) -> Result<StreamPatchProgram, SidecarError> {
    let bytes = std::fs::read(path)?;
    let program: StreamPatchProgram = serde_json::from_slice(&bytes)
        .map_err(|error| SidecarError::InvalidRequest(format!("Bad patch program: {error}")))?;
    validate_program(&program)?;
    Ok(program)
}

fn validate_program(program: &StreamPatchProgram) -> Result<(), SidecarError> {
    let mut last_row = 0u32;
    let mut total = 0usize;
    for row in &program.rows {
        if row.row == 0 || row.row <= last_row {
            return Err(SidecarError::InvalidRequest(
                "Patch rows must be 1-based and strictly ascending.".into(),
            ));
        }
        last_row = row.row;
        if row.cells.is_empty() {
            return Err(SidecarError::InvalidRequest(
                "A patch row must carry at least one cell.".into(),
            ));
        }
        let mut last_col = 0u32;
        for cell in &row.cells {
            if cell.col == 0 || cell.col <= last_col {
                return Err(SidecarError::InvalidRequest(
                    "Patch cells must be 1-based and strictly ascending within their row.".into(),
                ));
            }
            last_col = cell.col;
            if cell.xml.is_empty() || cell.xml.len() > MAX_CELL_XML_BYTES {
                return Err(SidecarError::InvalidRequest(
                    "A patch cell's XML must be non-empty and bounded.".into(),
                ));
            }
            total += 1;
        }
    }
    if total > MAX_PATCHED_CELLS {
        return Err(SidecarError::InvalidRequest(
            "Patch program exceeds the cell count bound.".into(),
        ));
    }
    Ok(())
}

/// Chunked scanner over a decompressed entry. Consumed bytes either stream
/// straight to the output (copy mode) or into a bounded capture buffer, so
/// the window never holds more than one row (plus one chunk) at a time.
struct Scanner<R: Read> {
    inner: R,
    buf: Vec<u8>,
    start: usize,
    eof: bool,
}

impl<R: Read> Scanner<R> {
    fn new(inner: R) -> Self {
        Self { inner, buf: Vec::with_capacity(2 * CHUNK), start: 0, eof: false }
    }

    fn pending(&self) -> &[u8] {
        &self.buf[self.start..]
    }

    fn refill(&mut self) -> Result<usize, SidecarError> {
        if self.start > 0 {
            self.buf.drain(..self.start);
            self.start = 0;
        }
        let old = self.buf.len();
        self.buf.resize(old + CHUNK, 0);
        let n = self.inner.read(&mut self.buf[old..])?;
        self.buf.truncate(old + n);
        if n == 0 {
            self.eof = true;
        }
        Ok(n)
    }

    /// Make at least `n` bytes visible without consuming; fewer only at EOF.
    fn peek(&mut self, n: usize) -> Result<&[u8], SidecarError> {
        while self.buf.len() - self.start < n && !self.eof {
            self.refill()?;
        }
        let end = (self.start + n).min(self.buf.len());
        Ok(&self.buf[self.start..end])
    }

    /// Stream to `out` until `needle` has been emitted (inclusive). Returns
    /// false when EOF arrives first (everything was copied).
    fn copy_through<W: Write>(&mut self, needle: &[u8], out: &mut W) -> Result<bool, SidecarError> {
        loop {
            if let Some(at) = find(self.pending(), needle) {
                let end = self.start + at + needle.len();
                out.write_all(&self.buf[self.start..end])?;
                self.start = end;
                return Ok(true);
            }
            // flush all but a needle-sized tail (a match could straddle chunks)
            let keep = needle.len().saturating_sub(1);
            let flush_to = self.buf.len().saturating_sub(keep).max(self.start);
            out.write_all(&self.buf[self.start..flush_to])?;
            self.start = flush_to;
            if self.eof {
                out.write_all(self.pending())?;
                self.start = self.buf.len();
                return Ok(false);
            }
            self.refill()?;
        }
    }

    /// Stream everything that remains to `out`.
    fn copy_rest<W: Write>(&mut self, out: &mut W) -> Result<(), SidecarError> {
        loop {
            out.write_all(self.pending())?;
            self.start = self.buf.len();
            if self.eof {
                return Ok(());
            }
            self.refill()?;
        }
    }

    /// Stream to `out` until the next `<`, which stays unconsumed. Returns
    /// false when EOF arrives first.
    fn copy_until_lt<W: Write>(&mut self, out: &mut W) -> Result<bool, SidecarError> {
        loop {
            if let Some(at) = self.pending().iter().position(|&b| b == b'<') {
                out.write_all(&self.buf[self.start..self.start + at])?;
                self.start += at;
                return Ok(true);
            }
            out.write_all(self.pending())?;
            self.start = self.buf.len();
            if self.eof {
                return Ok(false);
            }
            self.refill()?;
        }
    }

    /// Consume until `needle` inclusive, appending everything to `cap`.
    fn capture_through(
        &mut self,
        needle: &[u8],
        cap: &mut Vec<u8>,
        max: usize,
    ) -> Result<(), SidecarError> {
        loop {
            if let Some(at) = find(self.pending(), needle) {
                let end = self.start + at + needle.len();
                cap.extend_from_slice(&self.buf[self.start..end]);
                self.start = end;
                if cap.len() > max {
                    return Err(too_large());
                }
                return Ok(());
            }
            let keep = needle.len().saturating_sub(1);
            let take_to = self.buf.len().saturating_sub(keep).max(self.start);
            cap.extend_from_slice(&self.buf[self.start..take_to]);
            self.start = take_to;
            if cap.len() > max {
                return Err(too_large());
            }
            if self.eof {
                return Err(SidecarError::InvalidRequest(
                    "Worksheet part ended inside an element.".into(),
                ));
            }
            self.refill()?;
        }
    }
}

fn too_large() -> SidecarError {
    SidecarError::InvalidRequest("A worksheet row exceeds the streaming patch bound.".into())
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|w| w == needle)
}

/// `r="…"` attribute value within a tag's bytes.
fn attr_value<'a>(tag: &'a [u8], name: &[u8]) -> Option<&'a [u8]> {
    let mut i = 0;
    while let Some(at) = find(&tag[i..], name) {
        let p = i + at;
        // must be preceded by whitespace and followed by ="…"
        let before_ok = p > 0 && tag[p - 1].is_ascii_whitespace();
        let rest = &tag[p + name.len()..];
        if before_ok && rest.starts_with(b"=\"") {
            let v = &rest[2..];
            let end = v.iter().position(|&b| b == b'"')?;
            return Some(&v[..end]);
        }
        i = p + name.len();
    }
    None
}

fn parse_u32(bytes: &[u8]) -> Option<u32> {
    std::str::from_utf8(bytes).ok()?.parse().ok()
}

/// 1-based column from an A1 address's letter prefix.
fn col_of_address(addr: &[u8]) -> Option<u32> {
    let mut col = 0u32;
    let mut seen = false;
    for &b in addr {
        if b.is_ascii_uppercase() {
            col = col.checked_mul(26)?.checked_add(u32::from(b - b'A') + 1)?;
            seen = true;
        } else {
            break;
        }
    }
    if seen { Some(col) } else { None }
}

/// Tag spellings for the part's SpreadsheetML prefix (`b""` or `b"x:"`).
struct Tags {
    prefix: Vec<u8>,
    row_open: Vec<u8>,
    row_close: Vec<u8>,
    cell_open: Vec<u8>,
    cell_close: Vec<u8>,
    sheet_data_close: Vec<u8>,
}

impl Tags {
    fn new(prefix: &[u8]) -> Self {
        let tag = |open: &[u8], name: &[u8]| [open, prefix, name].concat();
        Self {
            prefix: prefix.to_vec(),
            row_open: tag(b"<", b"row"),
            row_close: tag(b"</", b"row>"),
            cell_open: tag(b"<", b"c"),
            cell_close: tag(b"</", b"c>"),
            sheet_data_close: tag(b"</", b"sheetData>"),
        }
    }

    /// True when the bytes after `<` open the named element in our prefix.
    fn is_element(&self, after_lt: &[u8], name: &[u8]) -> bool {
        after_lt
            .strip_prefix(self.prefix.as_slice())
            .is_some_and(|rest| is_local_element(rest, name))
    }

    /// The gateway serializes cells unprefixed; a prefixed part needs every
    /// element in the fragment (`c`, `f`, `v`, `is`, `t`, `r`, …) re-bound.
    fn prefixed_cell(&self, xml: &[u8], out: &mut Vec<u8>) {
        let mut i = 0;
        while i < xml.len() {
            let b = xml[i];
            out.push(b);
            if b == b'<' {
                let mut j = i + 1;
                if xml.get(j) == Some(&b'/') {
                    out.push(b'/');
                    j += 1;
                }
                if xml
                    .get(j)
                    .is_some_and(|c| c.is_ascii_alphabetic() || *c == b'_')
                {
                    out.extend_from_slice(&self.prefix);
                }
                i = j;
                continue;
            }
            i += 1;
        }
    }

    fn write_cell<W: Write>(&self, out: &mut W, xml: &str) -> Result<(), SidecarError> {
        if self.prefix.is_empty() {
            out.write_all(xml.as_bytes())?;
        } else {
            let mut bytes = Vec::with_capacity(xml.len() + 32);
            self.prefixed_cell(xml.as_bytes(), &mut bytes);
            out.write_all(&bytes)?;
        }
        Ok(())
    }
}

/// True when `after_lt` opens the named element (exact name, not a prefix).
fn is_local_element(after_lt: &[u8], name: &[u8]) -> bool {
    after_lt.starts_with(name)
        && matches!(
            after_lt.get(name.len()),
            Some(b' ' | b'\t' | b'\r' | b'\n' | b'>' | b'/')
        )
}

/// Split `prefix:` off a qualified tag name (bytes right after `<`).
fn split_qname(after_lt: &[u8]) -> (&[u8], &[u8]) {
    let name_end = after_lt
        .iter()
        .position(|b| !(b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-' | b'.' | b':')))
        .unwrap_or(after_lt.len());
    match after_lt[..name_end].iter().position(|&b| b == b':') {
        Some(colon) => after_lt.split_at(colon + 1),
        None => (&after_lt[..0], after_lt),
    }
}

fn write_new_row<W: Write>(out: &mut W, tags: &Tags, row: &RowPatch) -> Result<(), SidecarError> {
    out.write_all(&tags.row_open)?;
    write!(out, " r=\"{}\">", row.row)?;
    for cell in &row.cells {
        tags.write_cell(out, &cell.xml)?;
    }
    out.write_all(&tags.row_close)?;
    Ok(())
}

/// Splice `patch` cells into one buffered row. `row_tag` is the full open tag
/// (rewritten here only when self-closing); `content` is everything between
/// the open tag and `</row>` (empty for self-closing rows).
fn write_patched_row<W: Write>(
    out: &mut W,
    tags: &Tags,
    row_tag: &[u8],
    content: &[u8],
    patch: &RowPatch,
) -> Result<(), SidecarError> {
    if row_tag.ends_with(b"/>") {
        out.write_all(&row_tag[..row_tag.len() - 2])?;
        out.write_all(b">")?;
    } else {
        out.write_all(row_tag)?;
    }
    let mut cells = patch.cells.iter().peekable();
    let mut pos = 0usize;
    let mut implied_col = 0u32;
    loop {
        let Some(at) = find(&content[pos..], &tags.cell_open) else {
            break;
        };
        let tag_start = pos + at;
        if !tags.is_element(&content[tag_start + 1..], b"c") {
            // not a cell: copy up to it and stop scanning cells (extLst etc.
            // follows every <c> in the schema, so remaining upserts go first)
            break;
        }
        out.write_all(&content[pos..tag_start])?;
        let tag_end = tag_start
            + find(&content[tag_start..], b">")
                .ok_or_else(|| SidecarError::InvalidRequest("Unterminated cell tag.".into()))?
            + 1;
        let tag = &content[tag_start..tag_end];
        let col = attr_value(tag, b"r")
            .and_then(col_of_address)
            .unwrap_or(implied_col + 1);
        implied_col = col;
        let cell_end = if tag.ends_with(b"/>") {
            tag_end
        } else {
            tag_end
                + find(&content[tag_end..], &tags.cell_close)
                    .ok_or_else(|| SidecarError::InvalidRequest("Unterminated cell.".into()))?
                + tags.cell_close.len()
        };
        while let Some(next) = cells.peek() {
            if next.col < col {
                tags.write_cell(out, &cells.next().unwrap().xml)?;
            } else {
                break;
            }
        }
        if cells.peek().is_some_and(|next| next.col == col) {
            tags.write_cell(out, &cells.next().unwrap().xml)?;
        } else {
            out.write_all(&content[tag_start..cell_end])?;
        }
        pos = cell_end;
    }
    for cell in cells {
        tags.write_cell(out, &cell.xml)?;
    }
    out.write_all(&content[pos..])?;
    out.write_all(&tags.row_close)?;
    Ok(())
}

/// Stream up to and including `<[prefix:]sheetData`, returning the prefix.
fn copy_to_sheet_data<R: Read, W: Write>(
    scanner: &mut Scanner<R>,
    out: &mut W,
) -> Result<Option<Vec<u8>>, SidecarError> {
    loop {
        if !scanner.copy_until_lt(out)? {
            return Ok(None);
        }
        let ahead = scanner.peek(64)?.to_vec();
        let (prefix, local) = split_qname(&ahead[1..]);
        if is_local_element(local, b"sheetData") {
            let consumed = 1 + prefix.len() + b"sheetData".len();
            out.write_all(&ahead[..consumed])?;
            scanner.start += consumed;
            return Ok(Some(prefix.to_vec()));
        }
        out.write_all(&ahead[..1])?;
        scanner.start += 1;
    }
}

/// Stream `reader` to `out`, upserting the program's cells. Everything
/// outside the touched rows is copied byte-for-byte.
pub fn stream_patch_entry<R: Read, W: Write>(
    reader: R,
    out: &mut W,
    program: &StreamPatchProgram,
) -> Result<(), SidecarError> {
    let mut scanner = Scanner::new(reader);
    let Some(prefix) = copy_to_sheet_data(&mut scanner, out)? else {
        return Err(SidecarError::InvalidRequest(
            "Worksheet part has no sheetData element.".into(),
        ));
    };
    let tags = Tags::new(&prefix);
    let mut open_rest = Vec::new();
    scanner.capture_through(b">", &mut open_rest, 64 * 1024)?;
    let mut pending = program.rows.iter().peekable();
    if open_rest.ends_with(b"/>") {
        // empty sheet: expand <sheetData/> around the new rows
        out.write_all(&open_rest[..open_rest.len() - 2])?;
        out.write_all(b">")?;
        for row in pending {
            write_new_row(out, &tags, row)?;
        }
        out.write_all(&tags.sheet_data_close)?;
        scanner.copy_rest(out)?;
        return Ok(());
    }
    out.write_all(&open_rest)?;

    let mut implied_row = 0u32;
    loop {
        if !scanner.copy_until_lt(out)? {
            return Err(SidecarError::InvalidRequest(
                "Worksheet part ended inside sheetData.".into(),
            ));
        }
        let ahead = scanner.peek(64)?.to_vec();
        if !tags.is_element(&ahead[1..], b"row") {
            // </sheetData> (or schema-tail element): remaining new rows land here
            for row in pending {
                write_new_row(out, &tags, row)?;
            }
            scanner.copy_rest(out)?;
            return Ok(());
        }
        let mut row_tag = Vec::new();
        scanner.capture_through(b">", &mut row_tag, 1024 * 1024)?;
        let row_num = attr_value(&row_tag, b"r")
            .and_then(parse_u32)
            .unwrap_or(implied_row + 1);
        implied_row = row_num;
        while let Some(next) = pending.peek() {
            if next.row < row_num {
                write_new_row(out, &tags, pending.next().unwrap())?;
            } else {
                break;
            }
        }
        let self_closing = row_tag.ends_with(b"/>");
        if pending.peek().is_some_and(|next| next.row == row_num) {
            let patch = pending.next().unwrap();
            let mut content = Vec::new();
            if !self_closing {
                scanner.capture_through(&tags.row_close, &mut content, MAX_ROW_BYTES)?;
                content.truncate(content.len() - tags.row_close.len());
            }
            write_patched_row(out, &tags, &row_tag, &content, patch)?;
        } else {
            out.write_all(&row_tag)?;
            if !self_closing && !scanner.copy_through(&tags.row_close, out)? {
                return Err(SidecarError::InvalidRequest(
                    "Worksheet part ended inside a row.".into(),
                ));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn patch(rows: Vec<RowPatch>) -> StreamPatchProgram {
        StreamPatchProgram { rows }
    }

    fn cell(col: u32, xml: &str) -> CellUpsert {
        CellUpsert {
            col,
            xml: xml.into(),
        }
    }

    fn run(input: &str, program: &StreamPatchProgram) -> String {
        let mut out = Vec::new();
        stream_patch_entry(input.as_bytes(), &mut out, program).expect("patch");
        String::from_utf8(out).expect("utf8")
    }

    const DOC: &str = r#"<?xml version="1.0"?><worksheet><dimension ref="A1:C3"/><sheetData><row r="1" spans="1:3"><c r="A1"><v>1</v></c><c r="C1" t="s"><v>0</v></c></row><row r="3"><c r="B3"><v>3</v></c></row></sheetData><pageMargins left="0.7"/></worksheet>"#;

    #[test]
    fn replaces_an_existing_cell_and_keeps_every_other_byte() {
        let program = patch(vec![RowPatch {
            row: 1,
            cells: vec![cell(
                3,
                r#"<c r="C1" t="inlineStr"><is><t>new</t></is></c>"#,
            )],
        }]);
        let out = run(DOC, &program);
        assert!(out.contains(r#"<c r="C1" t="inlineStr"><is><t>new</t></is></c>"#));
        assert!(!out.contains(r#"<c r="C1" t="s">"#));
        // untouched regions byte-identical
        assert!(out.contains(r#"<row r="1" spans="1:3"><c r="A1"><v>1</v></c>"#));
        assert!(out.contains(r#"<row r="3"><c r="B3"><v>3</v></c></row>"#));
        assert!(out.starts_with(r#"<?xml version="1.0"?><worksheet><dimension ref="A1:C3"/>"#));
        assert!(out.ends_with(r#"<pageMargins left="0.7"/></worksheet>"#));
    }

    #[test]
    fn inserts_cells_in_column_order() {
        let program = patch(vec![RowPatch {
            row: 1,
            cells: vec![
                cell(2, r#"<c r="B1"><v>9</v></c>"#),
                cell(4, r#"<c r="D1"><v>8</v></c>"#),
            ],
        }]);
        let out = run(DOC, &program);
        assert!(out.contains(
            r#"<c r="A1"><v>1</v></c><c r="B1"><v>9</v></c><c r="C1" t="s"><v>0</v></c><c r="D1"><v>8</v></c></row>"#
        ));
    }

    #[test]
    fn inserts_new_rows_between_at_gap_and_after_last() {
        let program = patch(vec![
            RowPatch {
                row: 2,
                cells: vec![cell(1, r#"<c r="A2"><v>2</v></c>"#)],
            },
            RowPatch {
                row: 9,
                cells: vec![cell(1, r#"<c r="A9"><v>9</v></c>"#)],
            },
        ]);
        let out = run(DOC, &program);
        assert!(out.contains(r#"</row><row r="2"><c r="A2"><v>2</v></c></row><row r="3">"#));
        assert!(out.contains(r#"<row r="9"><c r="A9"><v>9</v></c></row></sheetData>"#));
    }

    #[test]
    fn expands_a_self_closing_sheet_data() {
        let empty = r#"<worksheet><sheetData/><pageMargins left="0.7"/></worksheet>"#;
        let program = patch(vec![RowPatch {
            row: 5,
            cells: vec![cell(2, r#"<c r="B5"><v>5</v></c>"#)],
        }]);
        let out = run(empty, &program);
        assert_eq!(
            out,
            r#"<worksheet><sheetData><row r="5"><c r="B5"><v>5</v></c></row></sheetData><pageMargins left="0.7"/></worksheet>"#
        );
    }

    #[test]
    fn patches_a_self_closing_row() {
        let doc = r#"<worksheet><sheetData><row r="2" ht="30" customHeight="1"/></sheetData></worksheet>"#;
        let program = patch(vec![RowPatch {
            row: 2,
            cells: vec![cell(1, r#"<c r="A2"><v>x</v></c>"#)],
        }]);
        let out = run(doc, &program);
        assert!(
            out.contains(r#"<row r="2" ht="30" customHeight="1"><c r="A2"><v>x</v></c></row>"#)
        );
    }

    #[test]
    fn handles_rows_and_cells_without_r_attributes() {
        let doc = r#"<worksheet><sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row></sheetData></worksheet>"#;
        let program = patch(vec![RowPatch {
            row: 2,
            cells: vec![cell(2, r#"<c r="B2"><v>9</v></c>"#)],
        }]);
        let out = run(doc, &program);
        assert!(out.contains(r#"<row><c><v>3</v></c><c r="B2"><v>9</v></c></row>"#));
    }

    #[test]
    fn upserts_before_a_row_ext_lst() {
        let doc = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><extLst><ext uri="x"/></extLst></row></sheetData></worksheet>"#;
        let program = patch(vec![RowPatch {
            row: 1,
            cells: vec![cell(3, r#"<c r="C1"><v>3</v></c>"#)],
        }]);
        let out = run(doc, &program);
        assert!(out.contains(
            r#"<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c><extLst><ext uri="x"/></extLst></row>"#
        ));
    }

    #[test]
    fn streams_large_untouched_rows_without_buffering() {
        // an unpatched row goes through copy mode, whatever its size
        let big_cells: String = (1..=200_000)
            .map(|i| format!(r#"<c r="A{i}"><v>{i}</v></c>"#))
            .collect();
        let doc = format!(
            r#"<worksheet><sheetData><row r="1">{big_cells}</row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>"#
        );
        let program = patch(vec![RowPatch {
            row: 2,
            cells: vec![cell(1, r#"<c r="A2"><v>9</v></c>"#)],
        }]);
        let out = run(&doc, &program);
        assert!(out.contains(r#"<row r="2"><c r="A2"><v>9</v></c></row>"#));
        assert_eq!(out.len(), doc.len()); // 9 replaced 2 — identical length
    }

    #[test]
    fn follows_the_part_namespace_prefix() {
        let doc = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c><x:c r="C1"><x:v>3</x:v></x:c></x:row><x:row r="4"/></x:sheetData><x:pageMargins left="0.7"/></x:worksheet>"#;
        let program = patch(vec![
            RowPatch {
                row: 1,
                cells: vec![
                    cell(
                        2,
                        r#"<c r="B1" t="inlineStr"><is><t xml:space="preserve">b </t></is></c>"#,
                    ),
                    cell(3, r#"<c r="C1"><f>A1*2</f><v>2</v></c>"#),
                ],
            },
            RowPatch {
                row: 2,
                cells: vec![cell(1, r#"<c r="A2"/>"#)],
            },
            RowPatch {
                row: 4,
                cells: vec![cell(1, r#"<c r="A4"><v>4</v></c>"#)],
            },
            RowPatch {
                row: 7,
                cells: vec![cell(1, r#"<c r="A7"><v>7</v></c>"#)],
            },
        ]);
        let out = run(doc, &program);
        assert_eq!(
            out,
            r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t xml:space="preserve">b </x:t></x:is></x:c><x:c r="C1"><x:f>A1*2</x:f><x:v>2</x:v></x:c></x:row><x:row r="2"><x:c r="A2"/></x:row><x:row r="4"><x:c r="A4"><x:v>4</x:v></x:c></x:row><x:row r="7"><x:c r="A7"><x:v>7</x:v></x:c></x:row></x:sheetData><x:pageMargins left="0.7"/></x:worksheet>"#
        );
    }

    #[test]
    fn expands_a_prefixed_self_closing_sheet_data() {
        let doc = r#"<x:worksheet xmlns:x="u"><x:sheetPr/><x:sheetData/></x:worksheet>"#;
        let program = patch(vec![RowPatch {
            row: 1,
            cells: vec![cell(1, r#"<c r="A1"><v>1</v></c>"#)],
        }]);
        assert_eq!(
            run(doc, &program),
            r#"<x:worksheet xmlns:x="u"><x:sheetPr/><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
    }

    #[test]
    fn rejects_bad_programs() {
        for rows in [
            vec![RowPatch {
                row: 0,
                cells: vec![cell(1, "<c/>")],
            }],
            vec![
                RowPatch {
                    row: 2,
                    cells: vec![cell(1, "<c/>")],
                },
                RowPatch {
                    row: 2,
                    cells: vec![cell(2, "<c/>")],
                },
            ],
            vec![RowPatch {
                row: 1,
                cells: vec![],
            }],
            vec![RowPatch {
                row: 1,
                cells: vec![cell(2, "<c/>"), cell(2, "<c/>")],
            }],
            vec![RowPatch {
                row: 1,
                cells: vec![cell(1, "")],
            }],
        ] {
            assert!(validate_program(&patch(rows)).is_err());
        }
    }
}

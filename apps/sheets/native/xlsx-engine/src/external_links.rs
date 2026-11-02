//! Excel-style external workbook links (genoffice#2109): parse the workbook's
//! `<externalReferences>` — the `[1]Sheet1!A1` indices, each source's stored
//! path and sheet names — resolve the source path the way Excel does, and, on
//! an explicit update, compute live values for formulas that reference a
//! reachable source. The referenced cells' values are inlined into the
//! formula so IronCalc, which cannot parse `[1]…` references, can evaluate
//! it; the stored formula keeps its indexed spelling, and only the computed
//! value travels back to the caller (the file's `<v>`).
//!
//! An unreachable source keeps the file's cached values — the published
//! genoffice#235 behaviour — and reports itself as broken. A formula whose
//! source cell errors, whose range explodes past the inline cap, or that is
//! one instance of a shared formula is left on its cached value too: a stale
//! number beats a wrong one.
//!
//! # Known gaps in the inlined semantics
//!
//! Inlining flattens a live range into an array constant (or a bare scalar),
//! which is a value, not a reference: `OFFSET`/`ROW`/`COLUMN`/`INDIRECT` over
//! an external range answer against the values and lose the source's shape,
//! and a multi-cell span in a scalar context (`=[1]S!A1:A3*2`) implicitly
//! intersects to a single cell rather than spilling per row. External defined
//! names (`[1]!Name`) and 3-D spans over the source are not inlined at all —
//! a formula still carrying such a spelling after the rewrite stays whole on
//! its cache. And a source cell that holds a formula but no cached `<v>`
//! (openpyxl-written files) inlines as `0` — a plausible number rather than a
//! "stay cached"; Excel would have computed the source formula.

use std::collections::HashMap;
use std::path::{Component, Path, PathBuf};

use ironcalc::base::Model;
use ironcalc::base::types::{Cell, CellType, FormulaValue};

use crate::SidecarError;
use crate::recalc::{
    blank_text_literals, pin_unparsable_formulas, references_an_external_workbook,
};
use crate::rels_normalize::load_model;

/// One `[n]` link of the workbook, as the report carries it.
#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalLinkStatus {
    pub index: usize,
    pub raw_target: String,
    /// The resolved local path when the target names one; `None` for schemes
    /// the sidecar cannot read (an `http://` source never becomes reachable).
    pub source: Option<String>,
    pub reachable: bool,
    pub sheet_names: Vec<String>,
}

/// One external-formula cell of the report, exactly one entry per
/// external-formula cell of the workbook: its freshly computed value when
/// every link it references was reachable and inlinable, its cached value
/// otherwise — never a mix, since one span that cannot go live aborts the
/// whole formula (a leftover `[n]…` would evaluate `#ERROR!` through the
/// spans that did inline). `formula` is always the stored `[1]…` spelling.
/// `row` and `column` are 0-based here; the IronCalc model the values were
/// computed in addresses the same cells 1-based.
#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalLinkCell {
    pub sheet: String,
    pub row: u32,
    pub column: u32,
    pub formatted: String,
    pub number: Option<f64>,
    pub formula: String,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExternalLinksReport {
    pub links: Vec<ExternalLinkStatus>,
    pub cells: Vec<ExternalLinkCell>,
}

/// A link's parsed identity, kept private: the report carries the flattened
/// status, the update pass needs nothing beyond the resolved source.
struct LinkEntry {
    index: usize,
    raw_target: String,
    source: Option<PathBuf>,
    sheet_names: Vec<String>,
}

impl LinkEntry {
    fn status(&self) -> ExternalLinkStatus {
        ExternalLinkStatus {
            index: self.index,
            raw_target: self.raw_target.clone(),
            source: self
                .source
                .as_ref()
                .map(|path| path.to_string_lossy().into_owned()),
            reachable: self.reachable(),
            sheet_names: self.sheet_names.clone(),
        }
    }

    fn reachable(&self) -> bool {
        self.source.as_ref().is_some_and(|path| path.is_file())
    }
}

/// Where a `[n]Sheet!REF` span points, normalized to 0-based bounds.
#[derive(Debug, Clone, Copy)]
enum RefSpan {
    Cell(u32, u32),
    Range(u32, u32, u32, u32),
}

impl RefSpan {
    fn bounds(&self) -> (u32, u32, u32, u32) {
        match *self {
            RefSpan::Cell(row, column) => (row, column, row, column),
            RefSpan::Range(r1, c1, r2, c2) => (r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)),
        }
    }
}

/// Inline spans over this many cells at most; a larger range stays on its
/// cached value. A `<v>`-worth of data, not a whole external sheet.
const MAX_INLINE_CELLS: u64 = 65_536;

/// Parses the workbook's external links in `[n]` index order.
fn parse_external_links(workbook: &Path) -> Result<Vec<LinkEntry>, SidecarError> {
    let mut links = Vec::new();
    let file = match std::fs::File::open(workbook) {
        Ok(file) => file,
        Err(_) => return Ok(links),
    };
    let mut archive = zip::ZipArchive::new(file)?;
    let (workbook_xml, rels_xml) = match (
        crate::xml_util::read_zip_string(&mut archive, "xl/workbook.xml"),
        crate::xml_util::read_zip_string(&mut archive, "xl/_rels/workbook.xml.rels"),
    ) {
        (Ok(xml), Ok(rels)) => (xml, rels),
        _ => return Ok(links),
    };
    let rel_targets = relationship_targets(&rels_xml);
    let doc = match roxmltree::Document::parse(&workbook_xml) {
        Ok(doc) => doc,
        Err(_) => return Ok(links),
    };
    let mut rids: Vec<(usize, String)> = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "externalReference" {
            // the element's only attribute is the relationship id (r:id)
            if let Some(r_id) = node.attributes().find(|a| a.name() == "id") {
                rids.push((rids.len() + 1, r_id.value().to_string()));
            }
        }
    }
    for (index, r_id) in rids {
        // The workbook rels resolve against the xl/ directory.
        let part = match rel_targets.get(&r_id) {
            Some(target) => join_package_path("xl", target),
            None => continue,
        };
        let part_xml = match crate::xml_util::read_zip_string(&mut archive, &part) {
            Ok(xml) => xml,
            Err(_) => continue,
        };
        let sheet_names = external_sheet_names(&part_xml);
        let raw_target = external_link_path(&mut archive, &part).unwrap_or_default();
        let source = resolve_external_path(workbook, &raw_target);
        links.push(LinkEntry {
            index,
            raw_target,
            source,
            sheet_names,
        });
    }
    Ok(links)
}

/// rId → target of one .rels part, by relationship Id attribute.
fn relationship_targets(rels_xml: &str) -> HashMap<String, String> {
    let mut out = HashMap::new();
    if let Ok(doc) = roxmltree::Document::parse(rels_xml) {
        for node in doc.descendants() {
            if node.tag_name().name() == "Relationship"
                && let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target"))
            {
                out.insert(id.to_string(), target.to_string());
            }
        }
    }
    out
}

/// The `<sheetName val>` list of an externalLink part.
fn external_sheet_names(part_xml: &str) -> Vec<String> {
    let mut out = Vec::new();
    if let Ok(doc) = roxmltree::Document::parse(part_xml) {
        for node in doc.descendants() {
            if node.tag_name().name() == "sheetName"
                && let Some(val) = node.attribute("val")
            {
                out.push(val.to_string());
            }
        }
    }
    out
}

/// The externalLinkPath target of one externalLink part's own rels — the
/// stored source path (`file:///C:/…`, a relative path, an UNC spelling).
fn external_link_path(archive: &mut zip::ZipArchive<std::fs::File>, part: &str) -> Option<String> {
    let (directory, file_name) = part.rsplit_once('/')?;
    let rels = format!("{directory}/_rels/{file_name}.rels");
    let rels_xml = crate::xml_util::read_zip_string(archive, &rels).ok()?;
    let doc = roxmltree::Document::parse(&rels_xml).ok()?;
    for node in doc.descendants() {
        if node.tag_name().name() != "Relationship" {
            continue;
        }
        let kind = node.attribute("Type").unwrap_or_default();
        if kind.ends_with("/externalLinkPath") {
            return node.attribute("Target").map(str::to_string);
        }
    }
    None
}

/// Resolves a stored source path the way Excel does: a `file:///` URI decodes
/// to a local path (a two-slash `file://host/…` is an UNC spelling), a path
/// that already looks absolute stays as written, and everything else is
/// relative to the dependent workbook's own folder. `None` marks a target the
/// sidecar cannot read at all (an `http://` source).
fn resolve_external_path(workbook: &Path, raw: &str) -> Option<PathBuf> {
    let raw = raw.trim();
    if let Some(rest) = raw.strip_prefix("file://") {
        if let Some(path) = rest.strip_prefix('/') {
            // `file:///C:/…` spells a drive: the third slash separates the
            // authority from the path and the path itself starts at the drive
            // letter. Any other `file:///…` is a POSIX absolute path (Excel
            // for Mac, LibreOffice, openpyxl) whose leading slash IS the path
            // — stripping it would resolve the target against the process's
            // cwd and make a sitting-on-disk source look gone.
            if starts_with_drive(path) {
                return Some(PathBuf::from(percent_decode(path)));
            }
            return Some(PathBuf::from(percent_decode(rest)));
        }
        // file://server/share/… — an UNC spelling
        let (host, path) = rest.split_once('/')?;
        return Some(PathBuf::from(format!(
            "\\\\{host}\\{}",
            percent_decode(path).replace('/', "\\")
        )));
    }
    if let Some(rest) = raw.strip_prefix("//") {
        return Some(PathBuf::from(format!(
            "\\\\{}",
            percent_decode(rest).replace('/', "\\")
        )));
    }
    // any other scheme (https://, http://, ftp://) is not a readable file
    if raw.contains("://") {
        return None;
    }
    if raw.starts_with('\\') || Path::new(raw).is_absolute() || starts_with_drive(raw) {
        return Some(Path::new(raw).to_path_buf());
    }
    let base = workbook.parent()?;
    Some(normalize_path(&base.join(raw)))
}

/// `C:/…` (or `C:\…`): a Windows drive spelling — absolute on Windows, and
/// relative-looking everywhere else.
fn starts_with_drive(path: &str) -> bool {
    let bytes = path.as_bytes();
    bytes.first().is_some_and(u8::is_ascii_alphabetic) && bytes.get(1) == Some(&b':')
}

fn percent_decode(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut at = 0;
    while at < bytes.len() {
        if bytes[at] == b'%'
            && let Some(hex) = bytes.get(at + 1..at + 3)
            && let Ok(value) = u8::from_str_radix(&String::from_utf8_lossy(hex), 16)
        {
            out.push(value);
            at += 3;
            continue;
        }
        out.push(bytes[at]);
        at += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

/// Collapses `.` and `..` components so the report shows the path Excel would.
fn normalize_path(path: &Path) -> PathBuf {
    let mut out = PathBuf::new();
    for component in path.components() {
        match component {
            Component::ParentDir => {
                out.pop();
            }
            Component::CurDir => {}
            other => out.push(other.as_os_str()),
        }
    }
    out
}

/// Joins a relationship target onto the part directory it resolves against,
/// in package-part spelling (`xl` + `externalLinks/externalLink1.xml`).
fn join_package_path(base: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    format!("{base}/{target}")
}

/// Computes fresh values for every external-formula cell against the reachable
/// sources. The workbook is never written here — the caller carries the
/// returned values into the save pipeline as cached `<v>`s, formulas intact.
pub(crate) fn update_external_link_values(
    workbook: &Path,
) -> Result<ExternalLinksReport, SidecarError> {
    let links = parse_external_links(workbook)?;
    let statuses: Vec<ExternalLinkStatus> = links.iter().map(LinkEntry::status).collect();

    // One live source model per reachable link, loaded lazily: an update of
    // three links where only one source exists opens one workbook.
    let mut sources: Vec<Option<Model<'static>>> = Vec::with_capacity(links.len());
    for link in &links {
        sources.push(if link.reachable() {
            load_model(link.source.as_ref().unwrap()).ok()
        } else {
            None
        });
    }

    let mut model = load_model(workbook)?;
    // External formula cells with their file-cached values, collected before
    // anything mutates the model: cells left on their cache report these.
    let mut pending: Vec<(u32, i32, i32, String, CachedValue)> = Vec::new();
    for (sheet, worksheet) in model.workbook.worksheets.iter().enumerate() {
        for (row, column, cell) in worksheet.sheet_data.cells() {
            let Cell::CellFormula { v, .. } = cell.as_ref() else {
                continue;
            };
            // every imported formula lives in shared_formulas (f is its
            // index), so there is no "shared instance" to skip: the cell's
            // own text comes back through get_cell_formula with its `=`
            let text = model
                .get_cell_formula(sheet as u32, row, column)
                .ok()
                .flatten();
            let Some(text) = text else { continue };
            if !references_an_external_workbook(&text) {
                continue;
            }
            pending.push((sheet as u32, row, column, text, CachedValue::of(v)));
        }
    }

    // recalc's own pin, for the same reason recalc applies it: IronCalc
    // cannot evaluate the `[n]…` spellings it was not rewritten away, so
    // every external formula (and every unparsable one) the update does not
    // rewrite would evaluate #ERROR! here and cascade through its dependents.
    // Pinning stands the file's cached value in for those cells first —
    // collected above, since a pin replaces the formula cell — so a rewritten
    // cell's dependents compute against caches, never against an error.
    pin_unparsable_formulas(&mut model);

    let mut rewritten: Vec<(u32, i32, i32)> = Vec::new();
    for (sheet, row, column, text, _) in &pending {
        // the rewrite output carries the formula's own leading `=`, which is
        // exactly the user-input form set_user_input expects
        let applied = rewrite_formula(text, &statuses, &mut sources)
            .and_then(|literal| model.set_user_input(*sheet, *row, *column, literal).ok())
            .is_some();
        if applied {
            rewritten.push((*sheet, *row, *column));
        }
    }
    if !rewritten.is_empty() {
        model.evaluate();
    }
    // One report entry per cell: rewritten cells read the model, the rest
    // report the file's cached value (genoffice#235).
    let mut cells: Vec<ExternalLinkCell> = Vec::new();
    for (sheet, row, column, text, cached) in &pending {
        let rewritten_now = rewritten.contains(&(*sheet, *row, *column));
        let (formatted, number) = if rewritten_now {
            (
                model
                    .get_formatted_cell_value(*sheet, *row, *column)
                    .unwrap_or_default(),
                raw_number(&model, *sheet, *row, *column),
            )
        } else {
            (cached.formatted(), cached.number())
        };
        cells.push(ExternalLinkCell {
            sheet: model.workbook.worksheets[*sheet as usize].name.clone(),
            row: (*row - 1) as u32,
            column: (*column - 1) as u32,
            formatted,
            number,
            formula: text.clone(),
        });
    }
    cells.sort_by_key(|cell| (cell.sheet.clone(), cell.row, cell.column));
    Ok(ExternalLinksReport {
        links: statuses,
        cells,
    })
}

fn raw_number(model: &Model, sheet: u32, row: i32, column: i32) -> Option<f64> {
    use ironcalc::base::cell::CellValue;
    match model.get_cell_value_by_index(sheet, row, column) {
        Ok(CellValue::Number(value)) => Some(value),
        _ => None,
    }
}

/// The file-cached value of a formula cell, for cells the update skips.
/// Carries the same kinds recalc's `PinnedValue` does — a `t="str"` cache is
/// a real string, not a blank.
#[derive(Debug, Clone)]
enum CachedValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Missing,
}

impl CachedValue {
    fn of(value: &FormulaValue) -> CachedValue {
        match value {
            FormulaValue::Number(number) => CachedValue::Number(*number),
            FormulaValue::Text(text) => CachedValue::Text(text.clone()),
            FormulaValue::Boolean(flag) => CachedValue::Bool(*flag),
            _ => CachedValue::Missing,
        }
    }

    fn number(&self) -> Option<f64> {
        match self {
            CachedValue::Number(number) => Some(*number),
            _ => None,
        }
    }

    /// A skipped cell's report value: the file's own cache, spelled the way
    /// the model carries it — numbers bare, text as the string itself.
    fn formatted(&self) -> String {
        match self {
            CachedValue::Number(number) => number.to_string(),
            CachedValue::Text(text) => text.clone(),
            CachedValue::Bool(true) => "TRUE".into(),
            CachedValue::Bool(false) => "FALSE".into(),
            CachedValue::Missing => String::new(),
        }
    }
}

/// Rewrites every `[n]Sheet!REF` span of one formula into an inline literal
/// of the live source's values. `None` leaves the formula on its cached
/// value: a referenced link unreachable, a sheet or range the source cannot
/// answer, an error cell inside the span. Any one span failing aborts the
/// whole formula — a leftover `[n]…` would evaluate `#ERROR!` and poison the
/// value the spans that did inline computed. External defined names
/// (`[1]!Name`) and 3-D spans over the source are not inlinable and abort
/// the same way: the whole cell stays on its cache rather than half-live.
fn rewrite_formula(
    text: &str,
    links: &[ExternalLinkStatus],
    sources: &mut [Option<Model<'static>>],
) -> Option<String> {
    // Span detection runs on the literal-blanked text: a `[1]Sheet1!A2`
    // inside a string literal is text the formula concatenates, not a span
    // to inline. Blanking replaces each character with spaces of the same
    // byte length, so spans found in the blanked copy copy straight out of
    // the original by offset. A blanked byte is never a reference start
    // (starts are `[` and `'`), so one space check gates them all.
    let code = blank_text_literals(text);
    let bytes = code.as_bytes();
    let mut out = String::with_capacity(text.len());
    let mut copied = 0usize;
    let mut at = 0usize;
    let mut changed = false;
    while at < bytes.len() {
        if bytes[at] == b' ' {
            at += 1;
            continue;
        }
        let Some((start, after_bang, link_index, sheet_name)) = external_ref_start(&code, at)
        else {
            at += 1;
            continue;
        };
        let Some((end, span)) = parse_ref_span(&code, after_bang) else {
            // a recognized external reference whose span does not parse
            // (3-D over the source, a deleted `#REF!` range): abort, never
            // rewrite the neighbours
            return None;
        };
        let Some(literal) = inline_literal(links, sources, link_index, &sheet_name, span) else {
            // this span cannot go live: the whole formula stays on its cache
            return None;
        };
        out.push_str(&text[copied..start]);
        out.push_str(&literal);
        copied = end;
        at = end;
        changed = true;
    }
    if !changed {
        return None;
    }
    out.push_str(&text[copied..]);
    // A spelling the span scanner does not recognize (`[1]!Name`, an external
    // defined name) must not survive into set_user_input: IronCalc stores the
    // formula unparsable and the cell evaluates `#ERROR!` as if it were live.
    if references_an_external_workbook(&out) {
        return None;
    }
    Some(out)
}

/// Finds the next external-reference start at or after `from`: either the
/// quoted form `'[1]Sheet Name'!` or the bare `[1]Sheet1!`. Yields the span
/// start (the quote when quoted), the position after the `!`, the link index
/// and the sheet name with its quotes stripped.
fn external_ref_start(text: &str, from: usize) -> Option<(usize, usize, usize, String)> {
    let bytes = text.as_bytes();
    let mut at = from;
    while at < bytes.len() {
        let (quoted, bracket) = match bytes[at] {
            b'\'' if bytes.get(at + 1) == Some(&b'[') => (true, at + 1),
            b'[' => (false, at),
            _ => {
                at += 1;
                continue;
            }
        };
        let mut digits_end = bracket + 1;
        while bytes.get(digits_end).is_some_and(u8::is_ascii_digit) {
            digits_end += 1;
        }
        if digits_end == bracket + 1 || bytes.get(digits_end) != Some(&b']') {
            at += 1;
            continue;
        }
        let index: usize = text[bracket + 1..digits_end].parse().ok()?;
        if index == 0 {
            return None;
        }
        if quoted {
            // the sheet name runs to the closing quote ('' escapes a quote)
            let mut name_at = digits_end + 1;
            let mut name = String::new();
            loop {
                match bytes.get(name_at) {
                    Some(b'\'') if bytes.get(name_at + 1) == Some(&b'\'') => {
                        name.push('\'');
                        name_at += 2;
                    }
                    Some(b'\'') => break,
                    Some(_) => {
                        let mut len = 1;
                        while name_at + len < text.len() && !text.is_char_boundary(name_at + len) {
                            len += 1;
                        }
                        name.push_str(&text[name_at..name_at + len]);
                        name_at += len;
                    }
                    None => return None,
                }
            }
            if bytes.get(name_at + 1) != Some(&b'!') {
                return None;
            }
            return Some((at, name_at + 2, index, name));
        }
        // bare form: sheet-name characters (word, `_`, any non-ASCII) run to `!`
        let mut name_end = digits_end + 1;
        while bytes
            .get(name_end)
            .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_' || *b >= 0x80)
        {
            name_end += 1;
        }
        if name_end == digits_end + 1 || bytes.get(name_end) != Some(&b'!') {
            at += 1;
            continue;
        }
        let name = text[digits_end + 1..name_end].to_string();
        return Some((at, name_end + 1, index, name));
    }
    None
}

/// Parses the `REF` (or `REF:REF`, `A:A`, `1:10`) that follows the `!` of an
/// external reference. Dollar signs are accepted and ignored: the inline
/// literal does not care about absoluteness.
fn parse_ref_span(text: &str, at: usize) -> Option<(usize, RefSpan)> {
    let bytes = text.as_bytes();
    let mut cursor = at;
    let read_letters = |cursor: &mut usize| -> Option<String> {
        let start = *cursor;
        while bytes
            .get(*cursor)
            .is_some_and(|b| b.is_ascii_alphabetic() || *b == b'$')
        {
            *cursor += 1;
        }
        if *cursor == start {
            return None;
        }
        Some(text[start..*cursor].replace('$', "").to_uppercase())
    };
    let read_digits = |cursor: &mut usize| -> Option<u32> {
        let start = *cursor;
        while bytes.get(*cursor).is_some_and(u8::is_ascii_digit) {
            *cursor += 1;
        }
        if *cursor == start {
            return None;
        }
        text[start..*cursor].parse().ok()
    };
    let column_of = |letters: &str| -> Option<u32> {
        if letters.is_empty() || letters.len() > 3 {
            return None;
        }
        let mut value = 0u32;
        for byte in letters.bytes() {
            if !byte.is_ascii_uppercase() {
                return None;
            }
            value = value * 26 + u32::from(byte - b'A') + 1;
        }
        Some(value - 1)
    };
    match read_letters(&mut cursor) {
        Some(letters) => match read_digits(&mut cursor) {
            Some(row) => {
                let column = column_of(&letters)?;
                if bytes.get(cursor) == Some(&b':') {
                    cursor += 1;
                    let second_letters = read_letters(&mut cursor)?;
                    let second_row = read_digits(&mut cursor)?;
                    let second_column = column_of(&second_letters)?;
                    Some((
                        cursor,
                        RefSpan::Range(row - 1, column, second_row - 1, second_column),
                    ))
                } else {
                    Some((cursor, RefSpan::Cell(row - 1, column)))
                }
            }
            None => {
                // A:A — a full-column span
                if bytes.get(cursor) != Some(&b':') {
                    return None;
                }
                cursor += 1;
                let second_letters = read_letters(&mut cursor)?;
                let first = column_of(&letters)?;
                let second = column_of(&second_letters)?;
                Some((
                    cursor,
                    RefSpan::Range(0, first.min(second), u32::MAX, first.max(second)),
                ))
            }
        },
        None => {
            // 1:10 — a full-row span
            let first_row = read_digits(&mut cursor)?;
            if bytes.get(cursor) != Some(&b':') {
                return None;
            }
            cursor += 1;
            let second_row = read_digits(&mut cursor)?;
            Some((
                cursor,
                RefSpan::Range(first_row - 1, 0, second_row - 1, u32::MAX),
            ))
        }
    }
}

/// The inline literal of one reference span: a scalar for a single cell, an
/// array constant `{1,2;3,4}` for a range (cells of a row joined by `,`,
/// rows by `;`). Full-column and full-row spans are bounded to the source
/// sheet's used extent; anything past the cap, a missing sheet, or an error
/// cell inside the span skips the whole formula.
///
/// The array constant is a value, not a reference: `OFFSET`/`ROW`/`COLUMN`/
/// `INDIRECT` over an external range answer against the values and lose the
/// source's shape (the row of a 1×N constant is 1, not the source rows), and
/// a multi-cell span in a scalar context (`=[1]S!A1:A3*2`) implicitly
/// intersects to one cell rather than spilling per row — where Excel's live
/// recalculation would spill or shift with the formula's own row.
fn inline_literal(
    links: &[ExternalLinkStatus],
    sources: &mut [Option<Model<'static>>],
    link_index: usize,
    sheet_name: &str,
    span: RefSpan,
) -> Option<String> {
    let link = links.get(link_index.checked_sub(1)?)?;
    if !link.reachable {
        return None;
    }
    let model = sources.get_mut(link_index - 1)?.as_mut()?;
    let sheet = source_sheet_index(model, sheet_name, link)?;
    let (row1, column1, mut row2, mut column2) = span.bounds();
    if row2 == u32::MAX || column2 == u32::MAX {
        // a full-column / full-row span bounds to the used extent
        let (max_row, max_column) = used_extent(model, sheet);
        if row2 == u32::MAX {
            row2 = max_row;
        }
        if column2 == u32::MAX {
            column2 = max_column;
        }
    }
    let (row1, column1, row2, column2) = (
        row1.min(row2),
        column1.min(column2),
        row1.max(row2),
        column1.max(column2),
    );
    let rows = row2 - row1 + 1;
    let columns = column2 - column1 + 1;
    if u64::from(rows) * u64::from(columns) > MAX_INLINE_CELLS {
        return None;
    }
    let mut literal = String::from("{");
    for row in row1..=row2 {
        if row > row1 {
            literal.push(';');
        }
        for column in column1..=column2 {
            if column > column1 {
                literal.push(',');
            }
            literal.push_str(&cell_literal(model, sheet, row, column)?);
        }
    }
    literal.push('}');
    if rows == 1 && columns == 1 {
        // a single-cell span reads as the bare scalar
        literal = literal[1..literal.len() - 1].to_string();
    }
    Some(literal)
}

/// The source sheet matching a reference's sheet name — by name first (Excel
/// is case-insensitive here), then by the link's own `<sheetName>` order.
fn source_sheet_index(model: &Model, sheet_name: &str, link: &ExternalLinkStatus) -> Option<u32> {
    let wanted = sheet_name.to_lowercase();
    if let Some(index) = model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name.to_lowercase() == wanted)
    {
        return Some(index as u32);
    }
    let position = link
        .sheet_names
        .iter()
        .position(|name| name.to_lowercase() == wanted)?;
    (position < model.workbook.worksheets.len()).then_some(position as u32)
}

/// The used extent (max row, max column, 0-based) of a source sheet, from its
/// stored cells; an empty sheet reports extent 0 (a span over it inlines as
/// a single zero).
fn used_extent(model: &Model, sheet: u32) -> (u32, u32) {
    let worksheet = &model.workbook.worksheets[sheet as usize];
    let mut max_row = 0u32;
    let mut max_column = 0u32;
    for (row, column, _cell) in worksheet.sheet_data.cells() {
        max_row = max_row.max((row - 1) as u32);
        max_column = max_column.max((column - 1) as u32);
    }
    (max_row, max_column)
}

/// One cell of the source as an inline literal: numbers bare, text quoted
/// with doubled quotes, booleans spelled, empty as `0` (Excel's own answer
/// for a blank external cell). An error cell skips the whole formula.
///
/// Known gap: a source cell holding a formula but no cached `<v>` —
/// openpyxl writes those — reads as an empty cell here and inlines as `0`, a
/// plausible wrong number rather than a "stay cached". Excel would compute
/// the source formula; the throwaway source model never evaluates it.
fn cell_literal(model: &Model, sheet: u32, row: u32, column: u32) -> Option<String> {
    use ironcalc::base::cell::CellValue;
    let row_1 = row as i32 + 1;
    let column_1 = column as i32 + 1;
    match model.get_cell_type(sheet, row_1, column_1) {
        Ok(CellType::Number) => model
            .get_cell_value_by_index(sheet, row_1, column_1)
            .ok()
            .map(|value| match value {
                CellValue::Number(number) => number.to_string(),
                _ => "0".into(),
            }),
        Ok(CellType::LogicalValue) => model
            .get_formatted_cell_value(sheet, row_1, column_1)
            .ok()
            .map(|text| text.to_uppercase()),
        Ok(CellType::Text) => model
            .get_formatted_cell_value(sheet, row_1, column_1)
            .ok()
            .map(|text| format!("\"{}\"", text.replace('"', "\"\""))),
        Ok(CellType::ErrorValue) => None,
        _ => Some("0".into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    /// A tiny source workbook: Sheet1 holds A1=10, A2=20, A3=30, B1="Total".
    pub(crate) fn write_source_fixture(path: &Path) {
        let entries: [(&str, &str); 6] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1" t="inlineStr"><is><t>Total</t></is></c></row><row r="2"><c r="A2"><v>20</v></c></row><row r="3"><c r="A3"><v>30</v></c></row></sheetData></worksheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    /// The dependent workbook: one external link whose rels carry `target`,
    /// and five formula cells — two single-cell references (one bare, one
    /// quoted), one range inside MATCH, one concatenation of a text cell,
    /// one reference to a sheet the source does not have, and one SUM over a
    /// rectangular span mixing numbers and text.
    pub(crate) fn write_dependent_fixture(path: &Path, target: &str) {
        let entries: [(&str, &str); 8] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId3"/></externalReferences></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
            (
                "xl/externalLinks/externalLink1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/><sheetName val="Missing Sheet"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>84</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#,
            ),
            (
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                &format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="{target}" TargetMode="External"/></Relationships>"#
                ),
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"><f>[1]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>MATCH(20,[1]Sheet1!A1:A3,0)</f><v>2</v></c><c r="C1" t="str"><f>[1]Sheet1!B1&amp;"!"</f><v>Total!</v></c></row><row r="2"><c r="A2"><f>SUM([1]Sheet1!A1:B3)</f><v>60</v></c></row><row r="3"><c r="A3"><f>'[1]Missing Sheet'!A1+1</f><v>9</v></c></row></sheetData></worksheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    /// The dependent workbook with TWO external links: link 1's target is
    /// given verbatim (the tests pass a `file:///…` URI or a relative path),
    /// link 2's target never exists. Five formula cells probe the partial
    /// rewrite: a formula mixing both links, link 2's own cell, a dependent
    /// of that cell, a text cache on the gone link, and a `[1]…` spelling
    /// inside a string literal.
    pub(crate) fn write_two_link_dependent_fixture(path: &Path, target1: &str, target2: &str) {
        let entries: [(&str, &str); 10] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId3"/><externalReference r:id="rId4"/></externalReferences></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
            (
                "xl/externalLinks/externalLink1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>10</v></cell></row></sheetDataSet></externalBook></externalLink>"#,
            ),
            (
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                &format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="{target1}" TargetMode="External"/></Relationships>"#
                ),
            ),
            (
                "xl/externalLinks/externalLink2.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>5</v></cell></row></sheetDataSet></externalBook></externalLink>"#,
            ),
            (
                "xl/externalLinks/_rels/externalLink2.xml.rels",
                &format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="{target2}" TargetMode="External"/></Relationships>"#
                ),
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"><f>[1]Sheet1!A1+[2]Sheet1!A1</f><v>99</v></c><c r="B1" t="str"><f>[2]Sheet1!B1&amp;"!"</f><v>Total!</v></c><c r="C1" t="str"><f>[1]Sheet1!A1&amp;"[1]Sheet1!A2"</f><v>old</v></c></row><row r="2"><c r="A2"><f>[2]Sheet1!A1</f><v>5</v></c></row><row r="3"><c r="A3"><f>[1]Sheet1!A1+A2</f><v>15</v></c></row></sheetData></worksheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    fn cell_of<'a>(report: &'a ExternalLinksReport, formula: &str) -> &'a ExternalLinkCell {
        report
            .cells
            .iter()
            .find(|cell| cell.formula.trim_start_matches('=') == formula)
            .unwrap()
    }

    /// A reachable source: every reference to it evaluates live — singles,
    /// a MATCH over a range (the issue's own shape), a text concatenation,
    /// a SUM over a rectangular span — while each stored formula keeps its
    /// `[1]…` spelling. A reference to a sheet the source does not have
    /// stays on the file's cached value.
    #[test]
    fn update_computes_live_values_from_a_reachable_source() {
        let dir = tempfile::tempdir().unwrap();
        write_source_fixture(&dir.path().join("source.xlsx"));
        let path = dir.path().join("book.xlsx");
        // a relative target resolves against the dependent workbook's folder
        write_dependent_fixture(&path, "source.xlsx");
        let report = update_external_link_values(&path).unwrap();
        assert_eq!(report.links.len(), 1);
        assert!(report.links[0].reachable);
        assert!(
            report.links[0]
                .source
                .as_deref()
                .unwrap()
                .ends_with("source.xlsx")
        );
        assert_eq!(report.links[0].sheet_names, vec!["Sheet1", "Missing Sheet"]);

        let live = cell_of(&report, "[1]Sheet1!A1*2");
        assert_eq!((live.formatted.as_str(), live.number), ("20", Some(20.0)));
        let matched = cell_of(&report, "MATCH(20,[1]Sheet1!A1:A3,0)");
        assert_eq!(matched.formatted, "2");
        let text = cell_of(&report, "[1]Sheet1!B1&\"!\"");
        assert_eq!(text.formatted, "Total!");
        let summed = cell_of(&report, "SUM([1]Sheet1!A1:B3)");
        assert_eq!(summed.formatted, "60");
        // the source has no such sheet: the file's cached value stands
        let missing = cell_of(&report, "'[1]Missing Sheet'!A1+1");
        assert_eq!(missing.formatted, "9");
    }

    /// An unreachable source: every cell reports its cached value and the
    /// link reports itself broken — the published genoffice#235 behaviour.
    #[test]
    fn update_keeps_cached_values_when_the_source_is_gone() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("book.xlsx");
        write_dependent_fixture(&path, "gone.xlsx");
        let report = update_external_link_values(&path).unwrap();
        assert!(!report.links[0].reachable);
        assert_eq!(cell_of(&report, "[1]Sheet1!A1*2").formatted, "84");
        assert_eq!(
            cell_of(&report, "MATCH(20,[1]Sheet1!A1:A3,0)").formatted,
            "2"
        );
        assert_eq!(cell_of(&report, "SUM([1]Sheet1!A1:B3)").formatted, "60");
    }

    /// A partial rewrite pins, it does not poison: two links, the first
    /// reachable through a POSIX-absolute `file:///Users/…`-shaped URI (the
    /// leading slash must survive resolution), the second gone. A formula
    /// mixing both stays whole on its cache and reports once; the gone
    /// link's own cell and its dependents keep their caches instead of
    /// cascading `#ERROR!`; a text cache reports its text; and a `[1]…`
    /// spelling inside a string literal stays text.
    #[test]
    fn update_pins_the_cache_through_a_partial_rewrite() {
        let dir = tempfile::tempdir().unwrap();
        write_source_fixture(&dir.path().join("source.xlsx"));
        let path = dir.path().join("book.xlsx");
        let target = format!("file://{}", dir.path().join("source.xlsx").display());
        write_two_link_dependent_fixture(&path, &target, "gone.xlsx");
        let report = update_external_link_values(&path).unwrap();

        assert_eq!(report.links.len(), 2);
        assert!(
            report.links[0].reachable,
            "a file:///Users/… target keeps its leading slash: {:?}",
            report.links[0].source
        );
        assert!(!report.links[1].reachable);

        // exactly one report entry per external-formula cell
        assert_eq!(report.cells.len(), 5);
        assert_eq!(
            report
                .cells
                .iter()
                .filter(|cell| cell.formula == "=[1]Sheet1!A1+[2]Sheet1!A1")
                .count(),
            1
        );

        // one span of the mixed formula inlines, the other cannot: the whole
        // formula aborts and the file's cached value stands
        let mixed = cell_of(&report, "[1]Sheet1!A1+[2]Sheet1!A1");
        assert_eq!((mixed.formatted.as_str(), mixed.number), ("99", Some(99.0)));
        // the gone link's own cell: cached, never an error
        let gone = cell_of(&report, "[2]Sheet1!A1");
        assert_eq!((gone.formatted.as_str(), gone.number), ("5", Some(5.0)));
        // its dependent computes against the pinned cache, not against the
        // #ERROR! an unpinned formula would evaluate to
        let dependent = cell_of(&report, "[1]Sheet1!A1+A2");
        assert_eq!(
            (dependent.formatted.as_str(), dependent.number),
            ("15", Some(15.0))
        );
        // a text cache survives the skip
        let text = cell_of(&report, "[2]Sheet1!B1&\"!\"");
        assert_eq!((text.formatted.as_str(), text.number), ("Total!", None));
        // the span inside the string literal is text, not a reference: only
        // the live `[1]Sheet1!A1` inlines
        let literal = cell_of(&report, "[1]Sheet1!A1&\"[1]Sheet1!A2\"");
        assert_eq!(literal.formatted, "10[1]Sheet1!A2");
    }

    /// Path resolution, Excel-style: file URIs, UNC spellings, drive paths
    /// and folder-relative targets.
    #[test]
    fn resolves_stored_source_paths() {
        let workbook = Path::new("/tmp/dep/book.xlsx");
        assert_eq!(
            resolve_external_path(workbook, "file:///C:/data/source.xlsx").unwrap(),
            PathBuf::from("C:/data/source.xlsx")
        );
        // Excel for Mac, LibreOffice and openpyxl spell a POSIX source with
        // all three slashes; the leading slash is part of the path
        assert_eq!(
            resolve_external_path(workbook, "file:///Users/me/source.xlsx").unwrap(),
            PathBuf::from("/Users/me/source.xlsx")
        );
        assert_eq!(
            resolve_external_path(workbook, "source.xlsx").unwrap(),
            PathBuf::from("/tmp/dep/source.xlsx")
        );
        assert_eq!(
            resolve_external_path(workbook, "../data/source.xlsx").unwrap(),
            PathBuf::from("/tmp/data/source.xlsx")
        );
        assert_eq!(
            resolve_external_path(workbook, "\\\\server\\share\\s.xlsx").unwrap(),
            PathBuf::from("\\\\server\\share\\s.xlsx")
        );
        assert_eq!(
            resolve_external_path(workbook, "https://example.com/s.xlsx"),
            None
        );
    }
}

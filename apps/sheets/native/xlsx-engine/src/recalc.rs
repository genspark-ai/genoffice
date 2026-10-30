//! IronCalc-backed formula recalculation with a session-resident model cache.
//!
//! The first request for a workbook loads it into IronCalc's model; later
//! requests reuse that model and only apply the edits that changed, so a
//! debounced keystroke pays evaluate() instead of a full file import. The
//! cache invalidates itself when the file on disk changes (mtime+size) and
//! rebuilds when the edit set shrinks (an undo must restore file content the
//! model no longer has). The caller stays fail-soft: any load or evaluation
//! problem (including panics from IronCalc's strict importer) surfaces as a
//! normal sidecar error and the renderer keeps showing cached values.

use std::collections::HashMap;
use std::io::Read;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

use ironcalc::base::Model;
use ironcalc::base::types::{Cell, CellType, FormulaValue, SpillValue};
use serde::{Deserialize, Serialize};

use crate::rels_normalize::load_model;
use crate::{CellRange, SidecarError};

pub const MAX_RECALC_EDITS: usize = 10_000;
pub const MAX_RECALC_READ_CELLS: usize = 20_000;
/// Resident models are large (the whole workbook's cell graph); the sidecar
/// serves one document, so two covers the active file plus one recently
/// closed-and-reopened neighbour.
const MAX_RESIDENT_MODELS: usize = 2;
/// Source files above this size (compressed bytes) count as heavy for the
/// residency rule in `evict_beyond_cap`.
const HEAVY_SOURCE_BYTES: u64 = 8_000_000;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcEdit {
    pub sheet: String,
    /// 0-based coordinates on the wire (IronCalc itself is 1-based).
    pub row: u32,
    pub column: u32,
    /// User input: `=SUM(A1:A3)`, `42`, `text`; empty clears.
    pub input: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcRead {
    pub sheet: String,
    pub range: CellRange,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcCell {
    pub sheet: String,
    pub row: u32,
    pub column: u32,
    /// Display string after evaluation (number formats applied).
    pub formatted: String,
    /// Raw numeric value when the cell evaluates to a number.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number: Option<f64>,
    /// Raw boolean when the cell evaluates to a boolean (`=A1>3`): the cache
    /// for one is `t="b"` with the numeric 0/1, not the text `TRUE`/`FALSE`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub boolean: Option<bool>,
    /// The engine typed the result as an error (`#DIV/0!`), as opposed to a
    /// formula whose text result merely spells one (`="#N/A"`).
    pub is_error: bool,
    pub is_formula: bool,
}

/// Why a requested stale-cache audit came back without an answer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
pub enum AuditSkipReason {
    /// The workbook uses the 1904 date system (`<workbookPr date1904="1"/>`).
    /// IronCalc 0.9.1 has no date1904 handling — nothing in the engine reads
    /// the flag — so every date result it computes is 1462 days away from the
    /// value Excel cached. Comparing the two would overwrite correct caches.
    #[serde(rename = "date1904")]
    Date1904,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcResult {
    pub cells: Vec<RecalcCell>,
    /// True when a resident model served this request (no file re-import).
    pub cached: bool,
    /// Formula cells whose computed value differs from the cached `<v>` the
    /// file carried on import. Only present when the request asked for it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stale: Option<Vec<RecalcCell>>,
    /// Spill cells of a dynamic or CSE array formula whose cached `<v>` no
    /// longer matches what the anchor computes. They carry no formula of
    /// their own, so a caller cannot refresh them in place: this list is for
    /// reporting, not for patching.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stale_spill_cells: Option<Vec<RecalcCell>>,
    /// Formula and spill cells whose cache was kept although the engine now
    /// computes an error: the audit never downgrades a usable cache to an
    /// engine error (RATE/TEXT-style engine gaps), but a changed input can
    /// legitimately turn a dependent into `#DIV/0!`/`#N/A`, so these caches
    /// may be wrong. They were not refreshed — a caller may not claim the
    /// cached values are current while this list is non-empty.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stale_unrefreshed: Option<Vec<RecalcCell>>,
    /// Set when the audit was requested but deliberately skipped; `stale` and
    /// `stale_spill_cells` are then absent and the caches were not checked.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub audit_skipped: Option<AuditSkipReason>,
}

type EditKey = (String, u32, u32);

struct ResidentModel {
    model: Model<'static>,
    mtime: SystemTime,
    size: u64,
    /// Edits already in the model, keyed by cell; a request whose edit set no
    /// longer covers these keys forces a rebuild (only the file knows the
    /// original content of a reverted cell).
    applied: HashMap<EditKey, String>,
    last_used: u64,
}

#[derive(Default)]
pub struct RecalcCache {
    entries: HashMap<PathBuf, ResidentModel>,
    tick: u64,
}

impl RecalcCache {
    pub fn new() -> Self {
        Self::default()
    }

    /// Drop the model for a path whose bytes are about to change (save) or
    /// whose session closed.
    pub fn purge(&mut self, path: &Path) {
        self.entries.remove(&cache_key(path));
    }

    fn evict_beyond_cap(&mut self) {
        while self.entries.len() > MAX_RESIDENT_MODELS {
            let Some(oldest) = self
                .entries
                .iter()
                .min_by_key(|(_, entry)| entry.last_used)
                .map(|(path, _)| path.clone())
            else {
                return;
            };
            self.entries.remove(&oldest);
        }
        // Heavy sources import into models that dwarf everything else in the
        // process (a 31MB workbook's model holds ~1.2GB resident); keep at
        // most one of those — the most recently used.
        loop {
            let mut heavy: Vec<(PathBuf, u64)> = self
                .entries
                .iter()
                .filter(|(_, entry)| entry.size > HEAVY_SOURCE_BYTES)
                .map(|(path, entry)| (path.clone(), entry.last_used))
                .collect();
            if heavy.len() <= 1 {
                return;
            }
            heavy.sort_by_key(|(_, last_used)| *last_used);
            let Some((oldest, _)) = heavy.first() else {
                return;
            };
            self.entries.remove(&oldest.clone());
        }
    }
}

/// Cache keys are canonicalized: sessions store the canonical workbook path
/// (open() resolves it) while recalc requests carry the renderer's raw path,
/// and on macOS temp files those differ (/var vs /private/var) — a raw key
/// would make the close/save purge miss the resident model.
fn cache_key(path: &Path) -> PathBuf {
    crate::canonicalize(path).unwrap_or_else(|_| path.to_path_buf())
}

pub fn recalc_cells(
    cache: &mut RecalcCache,
    path: &Path,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
) -> Result<RecalcResult, SidecarError> {
    recalc_cells_reporting(cache, path, edits, reads, false)
}

/// [`recalc_cells`] with an optional stale-cache audit: with `report_stale`
/// the result lists every formula cell whose computed value differs from the
/// cached `<v>` the file carried on import, plus the array spill cells it
/// cannot refresh. The audit compares against values only an import stages, so
/// it always works on a fresh model (`cached: false`) — unless it is skipped,
/// which for now means a `date1904` workbook (`audit_skipped`), where the
/// engine's date results cannot be trusted against the file's caches.
pub fn recalc_cells_reporting(
    cache: &mut RecalcCache,
    path: &Path,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
    report_stale: bool,
) -> Result<RecalcResult, SidecarError> {
    if edits.len() > MAX_RECALC_EDITS {
        return Err(SidecarError::InvalidRequest(format!(
            "Too many recalc edits (limit {MAX_RECALC_EDITS})."
        )));
    }
    let read_cells: usize = reads
        .iter()
        .map(|read| {
            let rows = read.range.end_row.saturating_sub(read.range.start_row) + 1;
            let columns = read
                .range
                .end_column
                .saturating_sub(read.range.start_column)
                + 1;
            rows.saturating_mul(columns)
        })
        .sum();
    if read_cells > MAX_RECALC_READ_CELLS {
        return Err(SidecarError::InvalidRequest(format!(
            "Recalc read exceeds {MAX_RECALC_READ_CELLS} cells."
        )));
    }
    // An edit can only leave a cached formula value stale when the workbook
    // holds formulas at all. Finding that out by importing costs seconds on a
    // formula-free data workbook, so scan the worksheet parts first.
    if report_stale && reads.is_empty() && !workbook_has_formulas(path)? {
        return Ok(RecalcResult {
            cells: Vec::new(),
            cached: false,
            stale: Some(Vec::new()),
            stale_spill_cells: Some(Vec::new()),
            stale_unrefreshed: Some(Vec::new()),
            audit_skipped: None,
        });
    }
    // IronCalc 0.9.1 ignores workbookPr@date1904, so on a 1904-system workbook
    // every date it computes is 1462 days from the file's cache. The audit
    // would compare those caches against a wrong answer and overwrite values
    // Excel wrote correctly; skip it and tell the caller the caches were not
    // checked. Reads still run: the caller asked for values, not for the audit.
    let audit_skipped = if report_stale && workbook_uses_1904_dates(path)? {
        Some(AuditSkipReason::Date1904)
    } else {
        None
    };
    let report_stale = report_stale && audit_skipped.is_none();

    let metadata = std::fs::metadata(path).map_err(|error| SidecarError::Io(error.to_string()))?;
    let mtime = metadata
        .modified()
        .map_err(|error| SidecarError::Io(error.to_string()))?;
    let size = metadata.len();

    let key = cache_key(path);
    // Take the entry out while working on it: a panic or error mid-apply
    // leaves the model in an unknown state, and a removed entry can never be
    // reused by the next request.
    let resident = cache.entries.remove(&key).filter(|entry| {
        entry.mtime == mtime
            && entry.size == size
            && entry.applied.keys().all(|key| {
                edits
                    .iter()
                    .any(|edit| edit.sheet == key.0 && edit.row == key.1 && edit.column == key.2)
            })
    });
    // The audit compares against the caches the import staged; a resident
    // model has already evaluated them away.
    let resident = if report_stale { None } else { resident };
    let cached = resident.is_some();

    // IronCalc's importer is strict and can panic on non-Excel producers
    // (missing cellStyles, whitespace nodes); contain that to this request.
    let (entry, cells, stale_audit) = catch_unwind(AssertUnwindSafe(|| {
        run(resident, path, mtime, size, edits, reads, report_stale)
    }))
    .map_err(|_| {
        SidecarError::Workbook("The formula engine could not process this workbook.".into())
    })??;

    cache.tick += 1;
    let mut entry = entry;
    entry.last_used = cache.tick;
    cache.entries.insert(key, entry);
    cache.evict_beyond_cap();

    let (stale, stale_spill_cells, stale_unrefreshed) = match stale_audit {
        Some(audit) => (
            Some(audit.cells),
            Some(audit.spill_cells),
            Some(audit.unrefreshed),
        ),
        None => (None, None, None),
    };
    Ok(RecalcResult {
        cells,
        cached,
        stale,
        stale_spill_cells,
        stale_unrefreshed,
        audit_skipped,
    })
}

/// What the fresh import's staged caches no longer agree with: formula cells
/// (patchable on save), array spill cells (report only, see `RecalcResult`)
/// and the cells whose cache was kept against an engine error
/// (`unrefreshed`, also report only).
struct StaleAudit {
    cells: Vec<RecalcCell>,
    spill_cells: Vec<RecalcCell>,
    unrefreshed: Vec<RecalcCell>,
}

fn run(
    resident: Option<ResidentModel>,
    path: &Path,
    mtime: SystemTime,
    size: u64,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
    report_stale: bool,
) -> Result<(ResidentModel, Vec<RecalcCell>, Option<StaleAudit>), SidecarError> {
    let mut entry = match resident {
        Some(entry) => entry,
        None => {
            let mut model = load_model(path)?;
            requote_bare_sheet_names(&mut model);
            pin_unparsable_formulas(&mut model);
            ResidentModel {
                model,
                mtime,
                size,
                applied: HashMap::new(),
                last_used: 0,
            }
        }
    };

    // The staged `<v>` values are the file's own caches: they only exist
    // before the first evaluation replaces them with the engine's results.
    let cached_values = report_stale.then(|| snapshot_cached_values(&entry.model));

    let mut dirty = false;
    for edit in edits {
        let key = (edit.sheet.clone(), edit.row, edit.column);
        if entry.applied.get(&key) == Some(&edit.input) {
            continue;
        }
        let sheet = sheet_index(&entry.model, &edit.sheet)?;
        entry
            .model
            .set_user_input(
                sheet,
                edit.row as i32 + 1,
                edit.column as i32 + 1,
                edit.input.clone(),
            )
            .map_err(|error| {
                SidecarError::Workbook(format!("Formula engine rejected an edit: {error}"))
            })?;
        entry.applied.insert(key, edit.input.clone());
        dirty = true;
    }
    if dirty || entry.last_used == 0 {
        entry.model.evaluate();
    }

    let mut cells = Vec::new();
    for read in reads {
        let sheet = sheet_index(&entry.model, &read.sheet)?;
        for row in read.range.start_row..=read.range.end_row {
            for column in read.range.start_column..=read.range.end_column {
                let row_1 = row as i32 + 1;
                let column_1 = column as i32 + 1;
                let formatted = entry
                    .model
                    .get_formatted_cell_value(sheet, row_1, column_1)
                    .unwrap_or_default();
                let formula = entry
                    .model
                    .get_cell_formula(sheet, row_1, column_1)
                    .ok()
                    .flatten();
                let is_formula = formula.is_some();
                if formatted.is_empty() && !is_formula {
                    continue;
                }
                // Known engine gap where the file's cached value beats the
                // error: RATE's Newton solver dies near -100% where Excel
                // converges (#185).
                if formatted == "#NUM!" {
                    if let Some(text) = &formula {
                        if text.to_uppercase().contains("RATE(") {
                            continue;
                        }
                    }
                }
                let number = raw_number(&entry.model, sheet, row_1, column_1);
                let boolean = raw_boolean(&entry.model, sheet, row_1, column_1);
                let is_error = matches!(
                    entry.model.get_cell_type(sheet, row_1, column_1),
                    Ok(CellType::ErrorValue)
                );
                cells.push(RecalcCell {
                    sheet: read.sheet.clone(),
                    row: row as u32,
                    column: column as u32,
                    formatted,
                    number,
                    boolean,
                    is_error,
                    is_formula,
                });
            }
        }
    }
    let stale = cached_values.map(|cached| stale_cached_values(&entry.model, cached));
    Ok((entry, cells, stale))
}

/// IronCalc writes a sheet name that needs quoting without its quotes when it
/// converts an imported formula to R1C1: `name_needs_quoting`
/// (`expressions/utils`) lists only `()'$,;-+{}` and space, so a name holding
/// an `&` — `R&D`, `P&L` — slips through. The stored `R&D!R[-17]C[0]` never
/// parses again, and every reference to that sheet is lost along with
/// everything that depends on it.
///
/// Put the quotes back and re-parse. The text is the engine's own R1C1 and
/// the sheet is one of this workbook's, so this restores what the file said
/// rather than guessing at it: a formula is rewritten only when it already
/// failed to parse and only when the quoted form parses cleanly.
fn requote_bare_sheet_names(model: &mut Model) {
    use std::sync::Arc;

    use ironcalc::base::expressions::lexer::LexerMode;
    use ironcalc::base::expressions::parser::{
        Node, new_parser_english, static_analysis::StaticResult,
    };
    use ironcalc::base::expressions::types::CellReferenceRC;

    let names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|worksheet| worksheet.name.clone())
        .collect();
    let mut parser = new_parser_english(
        names.clone(),
        model.workbook.get_defined_names_with_scope(),
        model.workbook.tables.clone(),
    );
    parser.set_lexer_mode(LexerMode::R1C1);

    for (sheet, worksheet) in model.workbook.worksheets.iter_mut().enumerate() {
        let Some(parsed) = model.parsed_formulas.get_mut(sheet) else {
            continue;
        };
        // Stored R1C1 is not tied to a cell, so the sheet is the whole
        // context — the same one `Model::parse_formulas` uses.
        let context = CellReferenceRC {
            sheet: worksheet.name.clone(),
            row: 1,
            column: 1,
        };
        for (index, formula) in worksheet.shared_formulas.iter_mut().enumerate() {
            if !matches!(
                parsed.get(index),
                Some((node, _)) if matches!(node.as_ref(), Node::ParseErrorKind { .. })
            ) {
                continue;
            }
            let Some(quoted) = quote_bare_sheet_names(formula, &names) else {
                continue;
            };
            let node = parser.parse(&quoted, &context);
            if matches!(node, Node::ParseErrorKind { .. }) {
                continue;
            }
            // `run_static_analysis_on_node` is crate-private, so the repaired
            // node carries `Unknown`: evaluation then assumes nothing about
            // arrays or implicit intersection, which is the right reading for
            // a formula that failed to parse a moment ago.
            parsed[index] = (Arc::new(node), StaticResult::Unknown);
            *formula = quoted;
        }
    }
}

/// Wraps every bare `Name!` reference to one of `names` in quotes, doubling
/// any `'` inside the name as Excel does. Returns None when there was nothing
/// to quote. Text literals and already-quoted names are copied through
/// untouched, so a formula that merely spells a sheet name inside a string is
/// left alone.
fn quote_bare_sheet_names(formula: &str, names: &[String]) -> Option<String> {
    let mut out = String::with_capacity(formula.len());
    let mut index = 0;
    let mut in_text = false;
    let mut in_quoted_name = false;
    let mut changed = false;
    while index < formula.len() {
        let Some(char) = formula[index..].chars().next() else {
            break;
        };
        if in_text || in_quoted_name {
            // A doubled quote toggles off and straight back on, which lands
            // in the same state an escape should leave us in.
            if char == '"' && in_text {
                in_text = false;
            } else if char == '\'' && in_quoted_name {
                in_quoted_name = false;
            }
            out.push(char);
            index += char.len_utf8();
            continue;
        }
        if char == '"' || char == '\'' {
            if char == '"' {
                in_text = true;
            } else {
                in_quoted_name = true;
            }
            out.push(char);
            index += char.len_utf8();
            continue;
        }
        // Longest wins: "R&D" must not shadow "R&D Notes".
        let matched = names
            .iter()
            .filter(|name| !name.is_empty() && formula[index..].starts_with(&format!("{name}!")))
            .max_by_key(|name| name.len());
        if let Some(name) = matched
            && !continues_a_sheet_name(&formula[..index])
        {
            out.push('\'');
            out.push_str(&name.replace('\'', "''"));
            out.push_str("'!");
            index += name.len() + 1;
            changed = true;
            continue;
        }
        out.push(char);
        index += char.len_utf8();
    }
    changed.then_some(out)
}

enum PinnedValue {
    Number(f64),
    Text(String),
    Bool(bool),
}

/// IronCalc cannot parse external-workbook references (`[1]Sheet1!A1`):
/// such a formula evaluates to #ERROR! and the error cascades through every
/// dependent. Excel keeps the cached values when the linked workbook is
/// unreachable, so pin those cells to the value the file carries and let the
/// dependents compute against it. The renderer never sees the cell as a
/// formula afterwards, so its own cached copy stays on screen.
///
/// A formula still naming a sheet of *this* workbook after
/// `requote_bare_sheet_names` has had its turn is excluded: the data is in
/// the file, so a parse failure there is an engine gap, not an unreachable
/// source, and the cached value is not an answer — a file written without
/// cached values (openpyxl writes an empty `<v/>`) imports as zero, which
/// would pin a number nobody computed. Those cells keep their #ERROR!,
/// which the renderer already declines to display.
fn pin_unparsable_formulas(model: &mut Model) {
    use ironcalc::base::expressions::parser::Node;
    let names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|worksheet| worksheet.name.clone())
        .collect();
    let mut pins = Vec::new();
    for (sheet, worksheet) in model.workbook.worksheets.iter().enumerate() {
        let Some(parsed) = model.parsed_formulas.get(sheet) else {
            continue;
        };
        for (row, column, cell) in worksheet.sheet_data.cells() {
            let (formula, value) = match cell.as_ref() {
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Number(v),
                    ..
                } => (*f, PinnedValue::Number(*v)),
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Text(v),
                    ..
                } => (*f, PinnedValue::Text(v.clone())),
                Cell::CellFormula {
                    f,
                    v: FormulaValue::Boolean(v),
                    ..
                } => (*f, PinnedValue::Bool(*v)),
                _ => continue,
            };
            let unparsable = matches!(
                parsed.get(formula as usize),
                Some((node, _)) if matches!(node.as_ref(), Node::ParseErrorKind { .. })
            );
            let text = worksheet
                .shared_formulas
                .get(formula as usize)
                .map(String::as_str)
                .unwrap_or_default();
            if references_a_sheet_of_this_workbook(text, &names) {
                continue;
            }
            // 0.9.x parses an external-workbook reference instead of leaving a
            // parse error behind, and evaluation of it still lands on #REF!:
            // the cached value is the only truth for those cells either way.
            if !(unparsable || references_an_external_workbook(text)) {
                continue;
            }
            pins.push((sheet as u32, row, column, value));
        }
    }
    for (sheet, row, column, value) in pins {
        // A pin that fails leaves the cell as it was: #ERROR! on that cell,
        // which the renderer already declines to display.
        let _ = match value {
            PinnedValue::Number(number) => {
                model.update_cell_with_number(sheet, row, column, number)
            }
            PinnedValue::Text(text) => model.update_cell_with_text(sheet, row, column, &text),
            PinnedValue::Bool(flag) => model.update_cell_with_bool(sheet, row, column, flag),
        };
    }
}

/// True when the formula text carries a bare `Name!` reference to one of the
/// workbook's own sheets. The prefix test rejects the external-workbook forms
/// (`[1]Sheet1!A1`, `'[1]Sheet1'!A1`) and any longer name that merely ends in
/// this one.
fn references_a_sheet_of_this_workbook(formula: &str, names: &[String]) -> bool {
    let code = blank_text_literals(formula);
    names.iter().any(|name| {
        !name.is_empty()
            && code
                .match_indices(&format!("{name}!"))
                .any(|(at, _)| at == 0 || !continues_a_sheet_name(&code[..at]))
    })
}

/// True when the formula reads from another workbook — `[1]Sheet1!A1` or
/// `'[1]Sheet1'!A1`. IronCalc resolves no external data whether the
/// reference parses or not, so the file's cached value is the only truth.
/// A digits-only `[N]` is not enough on its own — an R1C1 offset brackets a
/// number too (`C[0]`) — so the bracket must be followed by a sheet name
/// and `!` to count as a workbook index. A name with spaces or CJK
/// characters only appears quoted, and Excel forbids `[` inside a sheet
/// name, so a `'` right before the bracket is always the external form.
fn references_an_external_workbook(formula: &str) -> bool {
    let code = blank_text_literals(formula);
    let bytes = code.as_bytes();
    for (at, byte) in bytes.iter().enumerate() {
        if *byte != b'[' {
            continue;
        }
        let mut end = at + 1;
        while end < bytes.len() && bytes[end].is_ascii_digit() {
            end += 1;
        }
        if end == at + 1 || end >= bytes.len() || bytes[end] != b']' {
            continue;
        }
        // `'[1]Sheet Name'!A1` — the opening quote precedes the bracket, so
        // the sheet name runs to the closing quote: spaces, CJK, anything
        if at > 0 && bytes[at - 1] == b'\'' {
            let mut close = end + 1;
            while close < bytes.len() && bytes[close] != b'\'' {
                close += 1;
            }
            if bytes.get(close + 1) == Some(&b'!') {
                return true;
            }
            continue;
        }
        let mut after = end + 1;
        if bytes.get(after) == Some(&b'\'') {
            after += 1;
        }
        let name_start = after;
        // word characters, or any non-ASCII byte: a CJK sheet name needs no
        // quotes (`[1]<cjk-name>!A1`), while a space does and lands in the branch above
        while after < bytes.len()
            && (bytes[after].is_ascii_alphanumeric()
                || bytes[after] == b'_'
                || bytes[after] >= 0x80)
        {
            after += 1;
        }
        if after > name_start && bytes.get(after) == Some(&b'\'') {
            after += 1;
        }
        if bytes.get(after) == Some(&b'!') {
            return true;
        }
    }
    false
}

/// The formula with every `"..."` literal replaced by spaces of the same byte
/// length, so a sheet name mentioned inside text does not count as a reference
/// and byte offsets still line up with the original.
fn blank_text_literals(formula: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut in_text = false;
    for char in formula.chars() {
        if char == '"' {
            in_text = !in_text;
            out.push(char);
        } else if in_text {
            out.extend(std::iter::repeat_n(' ', char.len_utf8()));
        } else {
            out.push(char);
        }
    }
    out
}

fn continues_a_sheet_name(before: &str) -> bool {
    match before.chars().next_back() {
        // ']' closes an external-workbook index, '\'' closes a quoted name
        // (which parses fine and never reaches here), '!' a 3-D range.
        Some(']') | Some('\'') | Some('!') => true,
        Some(char) => char.is_alphanumeric() || char == '_' || char == '.',
        None => false,
    }
}

fn sheet_index(model: &Model, name: &str) -> Result<u32, SidecarError> {
    model
        .workbook
        .worksheets
        .iter()
        .position(|worksheet| worksheet.name == name)
        .map(|index| index as u32)
        .ok_or_else(|| SidecarError::InvalidRequest(format!("Unknown sheet: {name}")))
}

fn raw_number(model: &Model, sheet: u32, row: i32, column: i32) -> Option<f64> {
    use ironcalc::base::cell::CellValue;
    match model.get_cell_value_by_index(sheet, row, column) {
        Ok(CellValue::Number(value)) => Some(value),
        _ => None,
    }
}

fn raw_boolean(model: &Model, sheet: u32, row: i32, column: i32) -> Option<bool> {
    use ironcalc::base::cell::CellValue;
    match model.get_cell_value_by_index(sheet, row, column) {
        Ok(CellValue::Boolean(value)) => Some(value),
        _ => None,
    }
}

/// A formula cell's identity and the cached `<v>` the file carried,
/// captured before the model evaluates it away. Array-formula anchors carry
/// their cached value the same way regular formulas do.
struct CachedFormula {
    sheet: usize,
    row: i32,
    column: i32,
    value: FormulaValue,
}

/// A spill cell — a member of a dynamic or CSE array formula that has no
/// formula of its own — and the value the file cached for it.
struct CachedSpill {
    sheet: usize,
    row: i32,
    column: i32,
    value: SpillValue,
}

struct CachedValues {
    formulas: Vec<CachedFormula>,
    spills: Vec<CachedSpill>,
}

fn snapshot_cached_values(model: &Model) -> CachedValues {
    let mut formulas = Vec::new();
    let mut spills = Vec::new();
    for (sheet, worksheet) in model.workbook.worksheets.iter().enumerate() {
        for (row, column, cell) in worksheet.sheet_data.cells() {
            match cell.as_ref() {
                Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => {
                    formulas.push(CachedFormula {
                        sheet,
                        row,
                        column,
                        value: v.clone(),
                    });
                }
                Cell::SpillCell { v, .. } => spills.push(CachedSpill {
                    sheet,
                    row,
                    column,
                    value: v.clone(),
                }),
                _ => {}
            }
        }
    }
    CachedValues { formulas, spills }
}

fn worksheet_cell<'a>(
    model: &'a Model<'a>,
    sheet: usize,
    row: i32,
    column: i32,
) -> Option<std::borrow::Cow<'a, Cell>> {
    model
        .workbook
        .worksheets
        .get(sheet)
        .and_then(|worksheet| worksheet.cell(row, column))
}

/// One reported stale cell in wire coordinates (0-based row/column, like
/// `reads`), with the engine's current display value.
fn stale_cell(model: &Model, sheet: usize, row: i32, column: i32, is_formula: bool) -> RecalcCell {
    let sheet_index = sheet as u32;
    RecalcCell {
        sheet: model.workbook.worksheets[sheet].name.clone(),
        // IronCalc coordinates are 1-based; the wire is 0-based like `reads`.
        row: (row - 1) as u32,
        column: (column - 1) as u32,
        formatted: model
            .get_formatted_cell_value(sheet_index, row, column)
            .unwrap_or_default(),
        number: raw_number(model, sheet_index, row, column),
        boolean: raw_boolean(model, sheet_index, row, column),
        is_error: matches!(
            model.get_cell_type(sheet_index, row, column),
            Ok(CellType::ErrorValue)
        ),
        is_formula,
    }
}

/// The cells whose evaluated value disagrees with the cached `<v>` the file
/// carried. Cells pinned to their cache (external-workbook references) are
/// value cells by the time this runs, so they are left alone.
fn stale_cached_values(model: &Model, cached: CachedValues) -> StaleAudit {
    let mut cells = Vec::new();
    let mut unrefreshed = Vec::new();
    for entry in cached.formulas {
        let Some(cell) = worksheet_cell(model, entry.sheet, entry.row, entry.column) else {
            continue;
        };
        let computed = match cell.as_ref() {
            Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => v,
            // The cell changed kind while evaluating (pinned to a constant,
            // dropped from an array): there is no computed value to compare.
            _ => continue,
        };
        if formula_values_agree(&entry.value, computed) {
            continue;
        }
        // Never downgrade a usable cache to an engine error. IronCalc errors
        // where Excel does not — RATE()'s Newton solver near -100% (#185) is
        // one case, TEXT() with a format code the engine cannot parse
        // another — and replacing the file's value with #NUM!/#VALUE! would
        // destroy a cache Excel computed correctly. The reverse (a cached
        // error that now computes) stays reportable. The cache is kept, but
        // the cell is listed as unrefreshed: a changed input legitimately
        // turns dependents into #DIV/0!/#N/A, so the kept cache may be wrong
        // and the caller may not claim the cached values are current.
        if is_error_value(computed) && !is_error_value(&entry.value) {
            unrefreshed.push(stale_cell(
                model,
                entry.sheet,
                entry.row,
                entry.column,
                true,
            ));
            continue;
        }
        // An unevaluated cell has no answer to write; the file's cache stays.
        if matches!(computed, FormulaValue::Unevaluated) {
            continue;
        }
        cells.push(stale_cell(
            model,
            entry.sheet,
            entry.row,
            entry.column,
            true,
        ));
    }
    let mut spill_cells = Vec::new();
    for entry in cached.spills {
        let Some(cell) = worksheet_cell(model, entry.sheet, entry.row, entry.column) else {
            continue;
        };
        let Cell::SpillCell { v, .. } = cell.as_ref() else {
            continue;
        };
        if spill_values_agree(&entry.value, v) {
            continue;
        }
        // Same rule as formula cells: a spill the engine turns into an error
        // keeps the value Excel spilled there, and is listed as unrefreshed
        // for the same reason.
        if matches!(v, SpillValue::Error(_)) && !matches!(entry.value, SpillValue::Error(_)) {
            unrefreshed.push(stale_cell(
                model,
                entry.sheet,
                entry.row,
                entry.column,
                false,
            ));
            continue;
        }
        // A spill cell carries no `<f>` of its own, so the caller's
        // cached-value patch cannot rewrite it; the caller reports it instead.
        spill_cells.push(stale_cell(
            model,
            entry.sheet,
            entry.row,
            entry.column,
            false,
        ));
    }
    StaleAudit {
        cells,
        spill_cells,
        unrefreshed,
    }
}

fn is_error_value(value: &FormulaValue) -> bool {
    matches!(value, FormulaValue::Error { .. })
}

fn spill_values_agree(cached: &SpillValue, computed: &SpillValue) -> bool {
    match (cached, computed) {
        (SpillValue::Number(a), SpillValue::Number(b)) => numbers_agree(*a, *b),
        (SpillValue::Text(a), SpillValue::Text(b)) => a == b,
        (SpillValue::Boolean(a), SpillValue::Boolean(b)) => a == b,
        (SpillValue::Error(a), SpillValue::Error(b)) => a == b,
        _ => false,
    }
}

fn formula_values_agree(cached: &FormulaValue, computed: &FormulaValue) -> bool {
    match (cached, computed) {
        (FormulaValue::Number(a), FormulaValue::Number(b)) => numbers_agree(*a, *b),
        (FormulaValue::Text(a), FormulaValue::Text(b)) => a == b,
        (FormulaValue::Boolean(a), FormulaValue::Boolean(b)) => a == b,
        (FormulaValue::Error { ei: a, .. }, FormulaValue::Error { ei: b, .. }) => a == b,
        _ => false,
    }
}

/// Excel keeps at most 15 significant digits in a cached `<v>`; two doubles
/// that agree there are the same cached number. Without this tolerance an
/// unrelated formula whose result Excel rounded would be rewritten on every
/// edit over the last bits of an f64.
fn numbers_agree(a: f64, b: f64) -> bool {
    if a == b {
        return true;
    }
    if a.is_nan() && b.is_nan() {
        return true;
    }
    if !a.is_finite() || !b.is_finite() {
        return false;
    }
    (a - b).abs() <= a.abs().max(b.abs()) * 1e-13
}

/// True when any worksheet part carries a `<f>` element. A stale-cache audit
/// needs the workbook imported; scanning the deflated worksheet parts first
/// lets a formula-free workbook skip that import entirely.
fn workbook_has_formulas(path: &Path) -> Result<bool, SidecarError> {
    let file = std::fs::File::open(path).map_err(|error| SidecarError::Io(error.to_string()))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|error| {
        SidecarError::Workbook(format!(
            "The formula engine could not read the workbook archive: {error}"
        ))
    })?;
    let mut names = Vec::new();
    for index in 0..archive.len() {
        let entry = archive.by_index(index).map_err(|error| {
            SidecarError::Workbook(format!(
                "The formula engine could not read the workbook archive: {error}"
            ))
        })?;
        let name = entry.name().to_owned();
        if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
            names.push(name);
        }
    }
    for name in names {
        let mut entry = archive.by_name(&name).map_err(|error| {
            SidecarError::Workbook(format!("The formula engine could not read {name}: {error}"))
        })?;
        if part_has_formula(&mut entry)? {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Streaming `<f` scan: a worksheet XML can be hundreds of MB, and only the
/// presence of a formula matters. The two-byte match can straddle a chunk.
fn part_has_formula(part: &mut impl Read) -> Result<bool, SidecarError> {
    let mut buffer = [0u8; 64 * 1024];
    let mut previous_was_lt = false;
    loop {
        let read = part
            .read(&mut buffer)
            .map_err(|error| SidecarError::Io(error.to_string()))?;
        if read == 0 {
            return Ok(false);
        }
        let chunk = &buffer[..read];
        if previous_was_lt && chunk[0] == b'f' {
            return Ok(true);
        }
        if chunk.windows(2).any(|pair| pair == b"<f") {
            return Ok(true);
        }
        previous_was_lt = chunk.last() == Some(&b'<');
    }
}

/// True when `xl/workbook.xml` declares the 1904 date system
/// (`<workbookPr date1904="1"/>`). IronCalc 0.9.1 has no date1904 support —
/// nothing in the engine reads the flag — so every date result it computes is
/// 1462 days away from the value Excel cached.
///
/// The workbook part is a small metadata part, and parsing it with the read
/// path's own `read_sheet_declarations` keeps this flag's interpretation
/// (`1`/`true`, namespaced documents) byte-for-byte consistent with the
/// renderer's.
fn workbook_uses_1904_dates(path: &Path) -> Result<bool, SidecarError> {
    let file = std::fs::File::open(path).map_err(|error| SidecarError::Io(error.to_string()))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|error| {
        SidecarError::Workbook(format!(
            "The formula engine could not read the workbook archive: {error}"
        ))
    })?;
    let (_, _, _, date_1904) = crate::workbook::read_sheet_declarations(&mut archive)?;
    Ok(date_1904)
}

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc::export::save_to_xlsx;

    fn fixture(path: &Path) {
        write_fixture(path, &[("A1", "10"), ("A2", "20"), ("A3", "=SUM(A1:A2)")]);
    }

    fn write_fixture(path: &Path, cells: &[(&str, &str)]) {
        let mut model = Model::new_empty("fixture", "en", "UTC", "en").unwrap();
        for (address, input) in cells {
            let column = (address.as_bytes()[0] - b'A' + 1) as i32;
            let row: i32 = address[1..].parse().unwrap();
            model
                .set_user_input(0, row, column, (*input).to_string())
                .unwrap();
        }
        model.evaluate();
        if path.exists() {
            std::fs::remove_file(path).unwrap();
        }
        save_to_xlsx(&model, path.to_str().unwrap()).unwrap();
    }

    fn edit(address: &str, input: &str) -> RecalcEdit {
        RecalcEdit {
            sheet: "Sheet1".into(),
            row: address[1..].parse::<u32>().unwrap() - 1,
            column: (address.as_bytes()[0] - b'A') as u32,
            input: input.into(),
        }
    }

    fn read_a1_a3() -> RecalcRead {
        RecalcRead {
            sheet: "Sheet1".into(),
            range: CellRange {
                start_row: 0,
                end_row: 2,
                start_column: 0,
                end_column: 0,
            },
        }
    }

    fn sum_value(result: &RecalcResult) -> String {
        result
            .cells
            .iter()
            .find(|cell| cell.row == 2)
            .unwrap()
            .formatted
            .clone()
    }

    #[test]
    fn keeps_at_most_one_heavy_model_resident() {
        let mut cache = RecalcCache::new();
        for (name, size) in [
            ("heavy-old", HEAVY_SOURCE_BYTES + 1),
            ("heavy-new", HEAVY_SOURCE_BYTES + 2),
            ("light", 1),
        ] {
            cache.tick += 1;
            cache.entries.insert(
                PathBuf::from(name),
                ResidentModel {
                    model: Model::new_empty("fixture", "en", "UTC", "en").unwrap(),
                    mtime: SystemTime::now(),
                    size,
                    applied: HashMap::new(),
                    last_used: cache.tick,
                },
            );
            cache.evict_beyond_cap();
        }
        assert!(!cache.entries.contains_key(Path::new("heavy-old")));
        assert!(cache.entries.contains_key(Path::new("heavy-new")));
        assert!(cache.entries.contains_key(Path::new("light")));
    }

    #[test]
    fn recalculates_after_edits() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let result =
            recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        let sum = result.cells.iter().find(|cell| cell.row == 2).unwrap();
        assert_eq!(sum.formatted, "120");
        assert_eq!(sum.number, Some(120.0));
        assert!(sum.is_formula);
        assert!(!result.cached);
    }

    #[test]
    fn types_error_results_but_not_error_looking_text() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[edit("A1", "=1/0"), edit("A2", "=\"#N/A\"")],
            &[read_a1_a3()],
        )
        .unwrap();
        let at = |row: u32| result.cells.iter().find(|cell| cell.row == row).unwrap();
        assert_eq!(at(0).formatted, "#DIV/0!");
        assert!(at(0).is_error);
        assert_eq!(at(1).formatted, "#N/A");
        assert!(!at(1).is_error);
    }

    #[test]
    fn reuses_the_resident_model_and_applies_edits_incrementally() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert!(!first.cached);
        assert_eq!(sum_value(&first), "120");
        // same edit again: reused model, nothing re-applied
        let second =
            recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert!(second.cached);
        assert_eq!(sum_value(&second), "120");
        // changed edit: reused model, incremental set_user_input
        let third = recalc_cells(&mut cache, &path, &[edit("A1", "200")], &[read_a1_a3()]).unwrap();
        assert!(third.cached);
        assert_eq!(sum_value(&third), "220");
    }

    #[test]
    fn rebuilds_when_an_edit_is_reverted() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[edit("A1", "100")], &[read_a1_a3()]).unwrap();
        assert_eq!(sum_value(&first), "120");
        // undo removed the A1 edit: only the file knows A1's original value
        let second = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!second.cached);
        assert_eq!(sum_value(&second), "30");
    }

    #[test]
    fn rebuilds_when_the_file_changes_on_disk() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let first = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!first.cached);
        assert_eq!(sum_value(&first), "30");
        write_fixture(
            &path,
            &[("A1", "11"), ("A2", "20"), ("A3", "=SUM(A1:A2)+1000")],
        );
        let second = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!second.cached);
        assert_eq!(sum_value(&second), "1031");
    }

    #[test]
    fn purge_drops_the_resident_model() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        cache.purge(&path);
        let result = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!result.cached);
    }

    #[test]
    fn an_errored_request_does_not_poison_the_cache() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        let error = recalc_cells(
            &mut cache,
            &path,
            &[],
            &[RecalcRead {
                sheet: "Nope".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
        // the entry was taken out and not re-inserted; next call reloads cleanly
        let after = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(!after.cached);
        assert_eq!(sum_value(&after), "30");
    }

    #[test]
    fn caps_resident_models() {
        let dir = tempfile::tempdir().unwrap();
        let mut cache = RecalcCache::new();
        let mut paths = Vec::new();
        for index in 0..3 {
            let path = dir.path().join(format!("recalc-{index}.xlsx"));
            fixture(&path);
            recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
            paths.push(path);
        }
        assert_eq!(cache.entries.len(), MAX_RESIDENT_MODELS);
        // the oldest was evicted, the newest survives (entries key by
        // canonical path)
        assert!(!cache.entries.contains_key(&cache_key(&paths[0])));
        assert!(cache.entries.contains_key(&cache_key(&paths[2])));
    }

    #[test]
    fn purge_hits_regardless_of_path_spelling() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert_eq!(cache.entries.len(), 1);
        // The raw and canonical spellings differ on macOS temp dirs
        // (/var vs /private/var); the close/save purge may hold either.
        cache.purge(&path);
        assert!(cache.entries.is_empty());
        recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        cache.purge(&cache_key(&path));
        assert!(cache.entries.is_empty());
    }

    #[test]
    fn rejects_unknown_sheets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("recalc.xlsx");
        fixture(&path);
        let error = recalc_cells(
            &mut RecalcCache::new(),
            &path,
            &[],
            &[RecalcRead {
                sheet: "Nope".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
    }

    #[test]
    fn caps_read_volume() {
        let error = recalc_cells(
            &mut RecalcCache::new(),
            Path::new("/nonexistent.xlsx"),
            &[],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 999,
                    start_column: 0,
                    end_column: 999,
                },
            }],
        )
        .unwrap_err();
        assert!(matches!(error, SidecarError::InvalidRequest(_)));
    }

    /// Two sheets, one named `R&D`, and a Summary formula written the way
    /// every other producer writes it: quoted name, no cached value.
    fn write_ampersand_sheet_fixture(path: &Path, formula: &str) {
        use std::io::Write;
        let sheet1 = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><f>{formula}</f><v /></c></row></sheetData></worksheet>"#
        );
        let entries: [(&str, &str); 7] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="R&amp;D" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
            ("xl/worksheets/sheet1.xml", &sheet1),
            (
                "xl/worksheets/sheet2.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetData><row r="1"><c r="A1"><v>59.5</v></c></row><row r="2"><c r="A2"><v>10</v></c></row></sheetData></worksheet>"#,
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

    fn read_summary_a1(path: &Path) -> RecalcCell {
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            path,
            &[],
            &[RecalcRead {
                sheet: "Summary".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 0,
                },
            }],
        )
        .unwrap();
        result.cells.into_iter().next().unwrap()
    }

    /// IronCalc 0.7.1 drops the quotes around a sheet name containing `&`
    /// when it converts an imported formula to R1C1 (`name_needs_quoting` in
    /// `expressions/utils` omits `&`), so the file's `'R&D'!A1` is stored as
    /// the unparsable `R&D!A1`. The sidecar puts the quotes back, and the
    /// value matches what Excel and the app's grid show.
    #[test]
    fn ampersand_sheet_references_resolve_across_sheets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("ampersand.xlsx");
        write_ampersand_sheet_fixture(&path, "'R&amp;D'!A1");
        let cell = read_summary_a1(&path);
        assert_eq!(cell.formatted, "59.5");
        assert_eq!(cell.number, Some(59.5));
        assert!(!cell.is_error);
    }

    /// One formula, several references to the same broken name: the repair
    /// walks the whole text, not just the first match.
    #[test]
    fn every_reference_in_one_formula_is_requoted() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("ampersand-twice.xlsx");
        write_ampersand_sheet_fixture(&path, "'R&amp;D'!A1+'R&amp;D'!A2");
        let cell = read_summary_a1(&path);
        assert_eq!(cell.number, Some(69.5));
        assert!(!cell.is_error);
    }

    /// A formula naming a sheet of this workbook that still will not parse
    /// after the repair keeps its #ERROR!. The pin exists for
    /// external-workbook references whose source is unreachable; here the
    /// data is in the file, and a producer that writes no cached value
    /// (openpyxl writes a bare `<v />`) imports as zero — pinning that would
    /// report a number nobody computed.
    #[test]
    fn an_irreparable_local_reference_is_not_pinned_to_an_invented_zero() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("malformed.xlsx");
        write_ampersand_sheet_fixture(&path, "R&amp;D!A1++");
        let cell = read_summary_a1(&path);
        assert_eq!(cell.formatted, "#ERROR!");
        assert!(cell.is_error);
        assert_eq!(cell.number, None);
    }

    #[test]
    fn quoting_covers_every_bare_reference_and_leaves_the_rest_alone() {
        let names = vec![
            "Sheet1".to_string(),
            "R&D".to_string(),
            "R&D Notes".to_string(),
        ];
        assert_eq!(
            quote_bare_sheet_names("R&D!R[-17]C[0]", &names).unwrap(),
            "'R&D'!R[-17]C[0]"
        );
        assert_eq!(
            quote_bare_sheet_names("R&D!R[1]C[1]+R&D!R[2]C[1]", &names).unwrap(),
            "'R&D'!R[1]C[1]+'R&D'!R[2]C[1]"
        );
        // longest name wins, so the shorter one does not truncate it
        assert_eq!(
            quote_bare_sheet_names("R&D Notes!R[1]C[1]", &names).unwrap(),
            "'R&D Notes'!R[1]C[1]"
        );
        // a sheet name inside a text literal is not a reference
        assert_eq!(quote_bare_sheet_names(r#""R&D!""#, &names), None);
        // external workbooks and already-quoted names are left as they are
        assert_eq!(quote_bare_sheet_names("[1]Sheet1!R[0]C[0]*2", &names), None);
        assert_eq!(quote_bare_sheet_names("'R&D'!R[1]C[1]", &names), None);
        assert_eq!(
            quote_bare_sheet_names("SUM(R[1]C[1]:R[2]C[1])", &names),
            None
        );
    }

    #[test]
    fn a_sheet_name_is_recognised_only_as_a_bare_reference() {
        let names = vec!["Sheet1".to_string(), "R&D".to_string()];
        assert!(references_a_sheet_of_this_workbook(
            "R&D!R[-17]C[0]",
            &names
        ));
        assert!(references_a_sheet_of_this_workbook(
            "SUM(R&D!R[1]C[1])",
            &names
        ));
        // external workbooks keep their pin
        assert!(!references_a_sheet_of_this_workbook(
            "[1]Sheet1!R[0]C[0]*2",
            &names
        ));
        assert!(!references_a_sheet_of_this_workbook(
            r#"'[1]Sheet1'!R[0]C[0]&" units""#,
            &names
        ));
        // a longer name that merely ends in one of ours
        assert!(!references_a_sheet_of_this_workbook(
            "OldSheet1!R[0]C[0]",
            &names
        ));
        // a sheet name inside a text literal is not a reference
        assert!(!references_a_sheet_of_this_workbook(
            r#"[1]Sheet1!R[0]C[0]&" from Sheet1! ""#,
            &names
        ));
        assert!(references_a_sheet_of_this_workbook(
            r#""see Sheet1!"&Sheet1!R[0]C[0]"#,
            &names
        ));
    }

    #[test]
    fn an_external_reference_with_a_spaced_or_cjk_sheet_name_is_recognised() {
        assert!(references_an_external_workbook("'[1]Sheet Name'!A1"));
        assert!(references_an_external_workbook("SUM('[1]Sheet Name'!A1:B2)"));
        // a CJK name needs no quotes
        assert!(references_an_external_workbook("[1]数据!A1"));
        // an R1C1 offset is still not an external workbook, quoted or not
        assert!(!references_an_external_workbook("R[1]C[0]"));
        assert!(!references_an_external_workbook("'My Sheet'!R[1]C[0]"));
        // a workbook index in a text literal is not a reference
        assert!(!references_an_external_workbook(r#""[1]Sheet Name""#));
    }

    fn write_external_link_fixture(path: &Path) {
        use std::io::Write;
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
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>"#,
            ),
            (
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/data/source.xlsx" TargetMode="External"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><f>[1]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>A1+1</f><v>85</v></c><c r="C1" t="str"><f>'[1]Sheet1'!A1&amp;" units"</f><v>42 units</v></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:A2)</f><v>89</v></c></row><row r="3"><c r="A3"><f>'[1]Sheet Name'!A1+1</f><v>43</v></c><c r="B3"><f>[1]数据!A1+1</f><v>8</v></c></row></sheetData></worksheet>"#,
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

    /// External-workbook references never parse in IronCalc; the file's
    /// cached values must stand in for them so dependents keep computing
    /// instead of cascading #ERROR! (public issue 235).
    #[test]
    fn external_link_formulas_keep_their_cached_values() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("external.xlsx");
        write_external_link_fixture(&path);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[edit("A2", "10")],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 2,
                    start_column: 0,
                    end_column: 2,
                },
            }],
        )
        .unwrap();
        let at = |row: u32, column: u32| {
            result
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
                .unwrap()
        };
        assert!(result.cells.iter().all(|cell| cell.formatted != "#ERROR!"));
        // pinned to the cache and no longer reported as formulas
        assert_eq!(
            (at(0, 0).formatted.as_str(), at(0, 0).is_formula),
            ("84", false)
        );
        assert_eq!(
            (at(0, 2).formatted.as_str(), at(0, 2).is_formula),
            ("42 units", false)
        );
        // dependents compute against the pinned values
        assert_eq!(
            (at(0, 1).formatted.as_str(), at(0, 1).is_formula),
            ("85", true)
        );
        assert_eq!(
            (at(1, 1).formatted.as_str(), at(1, 1).is_formula),
            ("94", true)
        );
        // a spaced sheet name (quoted) and a CJK one (bare) pin just the same
        assert_eq!(
            (at(2, 0).formatted.as_str(), at(2, 0).is_formula),
            ("43", false)
        );
        assert_eq!(
            (at(2, 1).formatted.as_str(), at(2, 1).is_formula),
            ("8", false)
        );
    }

    /// `xfId` is optional on a cellXfs `<xf>` (ECMA-376 CT_Xf; Excel defaults it
    /// to 0), and producers built on the Open XML SDK leave it out. IronCalc
    /// 0.7.1 read it with a required-attribute lookup, so the first formula
    /// edit failed to import and every later read of the saved file died with
    /// `Missing "xfId" XML attribute` (genoffice#2280).
    #[test]
    fn imports_a_workbook_whose_cell_xfs_omit_xfid() {
        use std::io::Write;

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("no-xfid.xlsx");
        let entries = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData></worksheet>"#,
            ),
            // The `<xf>` in `cellXfs` carries no `xfId` — the whole point.
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
        ];
        let mut writer = zip::ZipWriter::new(std::fs::File::create(&path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();

        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[RecalcEdit {
                sheet: "Sheet1".into(),
                row: 0,
                column: 1,
                input: "=A1*2".into(),
            }],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 0,
                    end_column: 1,
                },
            }],
        )
        .unwrap();
        let b1 = result
            .cells
            .iter()
            .find(|cell| cell.row == 0 && cell.column == 1)
            .unwrap();
        assert_eq!(b1.formatted, "6");
        assert!(b1.is_formula);
    }

    /// Rewrites one package part in place, the way an outside writer leaves a
    /// file whose inputs and cached formula values disagree.
    fn patch_entry(path: &Path, name: &str, from: &str, to: &str) {
        use std::io::{Read, Write};

        let mut archive = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut entries = Vec::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).unwrap();
            let entry_name = entry.name().to_owned();
            let mut content = Vec::new();
            entry.read_to_end(&mut content).unwrap();
            entries.push((entry_name, content));
        }
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (entry_name, content) in entries {
            let content = if entry_name == name {
                let text = String::from_utf8(content).unwrap();
                let patched = text.replace(from, to);
                assert_ne!(patched, text, "fixture part never carried {from:?}");
                patched.into_bytes()
            } else {
                content
            };
            writer.start_file(entry_name, options).unwrap();
            writer.write_all(&content).unwrap();
        }
        writer.finish().unwrap();
    }

    /// A value edit through an outside writer: A1 changed, B1's cached <v> not.
    #[test]
    fn reports_the_dependents_a_value_edit_left_stale() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("stale.xlsx");
        write_fixture(&path, &[("A1", "3"), ("B1", "=A1*2"), ("C1", "=B1+1")]);
        // A1 = 5, its cache rewritten, B1/C1 still carry the old results
        patch_entry(
            &path,
            "xl/worksheets/sheet1.xml",
            "<c r=\"A1\"><v>3</v></c>",
            "<c r=\"A1\"><v>5</v></c>",
        );

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        assert!(!result.cached);
        assert!(result.cells.is_empty());
        let stale = result.stale.unwrap();
        let at = |column: u32| {
            stale
                .iter()
                .find(|cell| cell.row == 0 && cell.column == column)
                .unwrap_or_else(|| panic!("column {column} was not reported stale"))
        };
        // the direct dependent and the one behind it, across the chain
        assert_eq!((at(1).formatted.as_str(), at(1).number), ("10", Some(10.0)));
        assert_eq!((at(2).formatted.as_str(), at(2).number), ("11", Some(11.0)));
        assert!(at(1).is_formula);
        assert!(at(2).is_formula);
    }

    #[test]
    fn reports_no_stale_cells_when_the_caches_agree() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("fresh.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        assert_eq!(result.stale.unwrap().len(), 0);
    }

    #[test]
    fn only_reports_stale_cells_when_the_request_asks_for_them() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("quiet.xlsx");
        fixture(&path);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(result.stale.is_none());
        assert!(result.stale_spill_cells.is_none());
        assert!(result.stale_unrefreshed.is_none());
        assert!(result.audit_skipped.is_none());
    }

    #[test]
    fn finds_formula_cells_without_importing_the_workbook() {
        use std::io::Write;

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("formulas.xlsx");
        let mut writer = zip::ZipWriter::new(std::fs::File::create(&path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        // Padding pushes the final "<f>" across the scan's 64 KiB chunk edge.
        let padding = " ".repeat(64 * 1024 - 1);
        let sheet = format!(
            "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>{padding}<f>1+1</f></worksheet>"
        );
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer.write_all(sheet.as_bytes()).unwrap();
        // A workbook IronCalc could not import (no workbook part) still answers
        // the scan, proving the audit never reaches the importer.
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer.write_all(b"<Types/>").unwrap();
        writer.finish().unwrap();
        assert!(workbook_has_formulas(&path).unwrap());

        let plain = dir.path().join("plain.xlsx");
        let mut writer = zip::ZipWriter::new(std::fs::File::create(&plain).unwrap());
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer.write_all(
            b"<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>",
        )
        .unwrap();
        writer.finish().unwrap();
        assert!(!workbook_has_formulas(&plain).unwrap());

        // no formulas: the audit answers empty without importing anything
        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &plain, &[], &[], true).unwrap();
        assert_eq!(result.stale.unwrap().len(), 0);
        assert!(result.cells.is_empty());
        assert_eq!(cache.entries.len(), 0);
    }
    /// An openpyxl-shaped workbook: the worksheet relationships use absolute
    /// OPC targets (`/xl/comments/comment1.xml`), and the comment part exists
    /// unless `include_comments` is false.
    fn write_absolute_relationship_fixture(path: &Path, include_comments: bool) {
        use std::io::Write;

        let mut entries: Vec<(&str, &str)> = vec![
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>"#,
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
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>10</v></c></row><row r="2"><c r="A2"><v>20</v></c></row><row r="3"><c r="A3"><f>SUM(A1:A2)</f></c></row></sheetData></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="/xl/comments/comment1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="/xl/drawings/commentsDrawing1.vml"/></Relationships>"#,
            ),
            (
                "xl/drawings/commentsDrawing1.vml",
                r#"<xml xmlns:v="urn:schemas-microsoft-com:vml"></xml>"#,
            ),
        ];
        if include_comments {
            entries.push((
                "xl/comments/comment1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>openpyxl</author></authors><commentList><comment ref="A1" authorId="0"><text><t>note</t></text></comment></commentList></comments>"#,
            ));
        }
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    /// openpyxl writes the comments and VML drawing relationships of a
    /// worksheet as absolute package targets. IronCalc's importer assumes a
    /// `../` relative target and failed the whole load with an opaque zip
    /// error (genoffice#2306); the sidecar now rewrites them before the
    /// import.
    #[test]
    fn imports_workbooks_with_absolute_relationship_targets() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("absolute-rels.xlsx");
        write_absolute_relationship_fixture(&path, true);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert_eq!(sum_value(&result), "30");
        assert!(!result.cached);
        // The resident model belongs to the source path; the rewrite is
        // per request and leaves no temp file behind to go stale.
        let again = recalc_cells(&mut cache, &path, &[], &[read_a1_a3()]).unwrap();
        assert!(again.cached);
        assert_eq!(sum_value(&again), "30");
    }

    #[test]
    fn names_a_missing_part_of_an_absolute_relationship_target() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("absolute-rels-missing.xlsx");
        write_absolute_relationship_fixture(&path, false);
        let error = recalc_cells(&mut RecalcCache::new(), &path, &[], &[read_a1_a3()]).unwrap_err();
        let SidecarError::Workbook(message) = error else {
            panic!("unexpected error: {error:?}");
        };
        assert!(
            message.contains("workbook relationship target not found:"),
            "{message}"
        );
        assert!(
            message.contains("xl/worksheets/sheet1.xml -> xl/comments/comment1.xml"),
            "{message}"
        );
        assert!(
            message.contains("(resolved from /xl/comments/comment1.xml)"),
            "{message}"
        );
    }

    // ── stale-cache audit: date1904, engine errors, array formulas ──────

    /// Boilerplate parts of a one-sheet package around `sheet_xml`.
    fn write_sheet_package(path: &Path, date_1904: bool, sheet_xml: &str) {
        use std::io::Write;

        let workbook_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            if date_1904 {
                r#"<workbookPr date1904="1"/>"#
            } else {
                ""
            }
        );
        let entries: [(&str, &str); 6] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            ("xl/workbook.xml", &workbook_xml),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#,
            ),
            ("xl/worksheets/sheet1.xml", sheet_xml),
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

    fn worksheet(dimension: &str, rows: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/><sheetData>{rows}</sheetData></worksheet>"#
        )
    }

    fn find_stale(result: &RecalcResult, row: u32, column: u32) -> Option<&RecalcCell> {
        result
            .stale
            .as_ref()?
            .iter()
            .find(|cell| cell.row == row && cell.column == column)
    }

    fn find_stale_spill(result: &RecalcResult, row: u32, column: u32) -> Option<&RecalcCell> {
        result
            .stale_spill_cells
            .as_ref()?
            .iter()
            .find(|cell| cell.row == row && cell.column == column)
    }

    #[test]
    fn detects_the_1904_date_system() {
        let dir = tempfile::tempdir().unwrap();
        let rows = r#"<row r="1"><c r="A1"><v>1</v></c></row>"#;
        let plain = dir.path().join("plain.xlsx");
        write_sheet_package(&plain, false, &worksheet("A1", rows));
        assert!(!workbook_uses_1904_dates(&plain).unwrap());
        let dated = dir.path().join("dated.xlsx");
        write_sheet_package(&dated, true, &worksheet("A1", rows));
        assert!(workbook_uses_1904_dates(&dated).unwrap());
    }

    /// IronCalc 0.9.1 evaluates dates on the 1900 system whatever the file
    /// says, so on a `date1904="1"` workbook its answer for a date formula is
    /// 1462 days away from Excel's cached value. The audit must not treat
    /// that difference as staleness and overwrite the file's caches; it skips
    /// and says why. The same sheet bytes without the flag are audited
    /// normally, which proves the gate is what changed.
    #[test]
    fn skips_the_audit_on_a_date_1904_workbook() {
        let dir = tempfile::tempdir().unwrap();
        let rows = r#"<row r="1"><c r="A1"><f>DATE(2024,1,1)</f><v>43830</v></c><c r="B1"><f>A1+1</f><v>43831</v></c></row>"#;
        let sheet = worksheet("A1:B1", rows);

        let dated = dir.path().join("dated.xlsx");
        write_sheet_package(&dated, true, &sheet);
        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &dated, &[], &[], true).unwrap();
        assert_eq!(result.audit_skipped, Some(AuditSkipReason::Date1904));
        assert!(result.stale.is_none());
        assert!(result.stale_spill_cells.is_none());
        assert!(result.stale_unrefreshed.is_none());
        assert!(result.cells.is_empty());

        let plain = dir.path().join("plain.xlsx");
        write_sheet_package(&plain, false, &sheet);
        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &plain, &[], &[], true).unwrap();
        assert!(result.audit_skipped.is_none());
        // both caches were written for the 1900 system, so both look stale
        let stale = result.stale.unwrap();
        assert_eq!(stale.len(), 2);
    }

    /// A formula-free 1904 workbook cannot carry a stale cache: the audit
    /// answers empty without importing, so it need not be skipped.
    #[test]
    fn audits_a_formula_free_date_1904_workbook_normally() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("data-1904.xlsx");
        write_sheet_package(
            &path,
            true,
            &worksheet("A1", r#"<row r="1"><c r="A1"><v>1</v></c></row>"#),
        );
        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        assert!(result.audit_skipped.is_none());
        assert_eq!(result.stale.unwrap().len(), 0);
        assert_eq!(result.stale_spill_cells.unwrap().len(), 0);
        assert_eq!(result.stale_unrefreshed.unwrap().len(), 0);
    }

    /// The audit never downgrades a usable cache to an engine error. Here
    /// `TEXT` uses a format code IronCalc's formatter rejects (`mmss.0`,
    /// ECMA-376 built-in 47): Excel's cached text is right, the engine's
    /// #VALUE! is not, so the cache is left in place — but the cell is
    /// reported as unrefreshed, because the caller may not claim the cached
    /// values are current while a cache went unverified. A cached error that
    /// now computes is still reported (and patchable).
    #[test]
    fn never_downgrades_a_usable_cache_to_an_engine_error() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("text-error.xlsx");
        let rows = concat!(
            r#"<row r="1"><c r="A1"><v>5</v></c>"#,
            r#"<c r="B1" t="str"><f>TEXT(A1,"mmss.0")</f><v>20:34.5</v></c>"#,
            r#"<c r="C1"><f>A1*2</f><v>8</v></c>"#,
            r#"<c r="D1"><f>A1*2</f><v>10</v></c>"#,
            r#"<c r="E1" t="e"><f>A1*2</f><v>#VALUE!</v></c></row>"#
        );
        write_sheet_package(&path, false, &worksheet("A1:E1", rows));

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        // the engine's #VALUE! must not replace the valid text cache …
        assert!(find_stale(&result, 0, 1).is_none());
        // … but the cell is listed: its cache was not refreshed
        let unrefreshed = result.stale_unrefreshed.as_ref().unwrap();
        let text = unrefreshed
            .iter()
            .find(|cell| cell.row == 0 && cell.column == 1)
            .expect("the TEXT cell must be reported unrefreshed");
        assert!(text.is_error);
        // the reverse direction stays reportable: a cached error that computes
        assert_eq!(
            find_stale(&result, 0, 4).map(|cell| (cell.number, cell.is_error)),
            Some((Some(10.0), false))
        );
        // and a genuinely stale number is still refreshed
        assert_eq!(
            find_stale(&result, 0, 2).map(|cell| cell.number),
            Some(Some(10.0))
        );
        assert!(find_stale(&result, 0, 3).is_none());
        assert!(unrefreshed.iter().all(|cell| cell.column == 1));
    }

    /// The maintainer's probe: A1 edited to 0 turns `D1 =10/A1` (cached 10)
    /// into #DIV/0!. The stale number is never downgraded to the engine's
    /// error, but the cell must be reported — silently keeping the old `<v>`
    /// while claiming `cached_values` would leave a wrong answer in the file.
    #[test]
    fn reports_a_dependent_that_turned_into_an_engine_error() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("div0.xlsx");
        let rows = r#"<row r="1"><c r="A1"><v>1</v></c><c r="D1"><f>10/A1</f><v>10</v></c></row>"#;
        write_sheet_package(&path, false, &worksheet("A1:D1", rows));
        patch_entry(
            &path,
            "xl/worksheets/sheet1.xml",
            "<c r=\"A1\"><v>1</v></c>",
            "<c r=\"A1\"><v>0</v></c>",
        );

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        // not patchable: the engine error must not replace the cached number
        assert!(find_stale(&result, 0, 3).is_none());
        let unrefreshed = result.stale_unrefreshed.expect("audit reported");
        let cell = unrefreshed
            .iter()
            .find(|cell| cell.row == 0 && cell.column == 3)
            .expect("D1 must be reported unrefreshed");
        assert_eq!(cell.formatted, "#DIV/0!");
        assert!(cell.is_error);
    }

    /// Excel caches a boolean formula result as `t="b"` with the numeric 0/1;
    /// the reported cell carries the raw boolean so the gateway can write the
    /// same shape instead of a `t="str"` `TRUE` text cache.
    #[test]
    fn reports_a_stale_boolean_dependent_as_a_boolean() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bool-stale.xlsx");
        let rows = concat!(
            r#"<row r="1"><c r="A1"><v>3</v></c>"#,
            r#"<c r="C1" t="b"><f>A1&gt;3</f><v>0</v></c>"#,
            r#"</row>"#
        );
        write_sheet_package(&path, false, &worksheet("A1:C1", rows));
        patch_entry(
            &path,
            "xl/worksheets/sheet1.xml",
            "<c r=\"A1\"><v>3</v></c>",
            "<c r=\"A1\"><v>10</v></c>",
        );

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        let cell = find_stale(&result, 0, 2).expect("C1 should be stale");
        assert_eq!(cell.boolean, Some(true));
        assert_eq!(cell.number, None);
        assert!(!cell.is_error);
    }

    /// The read path types boolean results too, so a freshly written `=A1>3`
    /// gets a `t="b"` cache on the refresh pass as well.
    #[test]
    fn reads_a_boolean_result_as_a_boolean() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bool-read.xlsx");
        write_fixture(&path, &[("A1", "3"), ("C1", "=A1>3")]);
        let mut cache = RecalcCache::new();
        let result = recalc_cells(
            &mut cache,
            &path,
            &[edit("A1", "10")],
            &[RecalcRead {
                sheet: "Sheet1".into(),
                range: CellRange {
                    start_row: 0,
                    end_row: 0,
                    start_column: 2,
                    end_column: 2,
                },
            }],
        )
        .unwrap();
        let cell = result.cells.into_iter().next().unwrap();
        assert_eq!(cell.boolean, Some(true));
        assert_eq!(cell.number, None);
        assert!(!cell.is_error);
        assert!(cell.is_formula);
    }

    /// A CSE array formula anchor is a formula cell whose cached first value
    /// can be compared (and patched) like any other; its spill cells cannot be
    /// patched, so they are reported separately.
    #[test]
    fn reports_a_dependent_cse_array_formula() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("cse.xlsx");
        let rows = concat!(
            r#"<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>A1*2</f><v>8</v></c></row>"#,
            r#"<row r="2"><c r="A2"><f t="array" ref="A2:A3">A1*2</f><v>6</v></c></row>"#,
            r#"<row r="3"><c r="A3"><v>6</v></c></row>"#
        );
        write_sheet_package(&path, false, &worksheet("A1:B3", rows));

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        // the regular dependent and the array anchor are both patchable
        assert_eq!(
            find_stale(&result, 0, 1).map(|cell| (cell.number, cell.is_formula)),
            Some((Some(10.0), true))
        );
        assert_eq!(
            find_stale(&result, 1, 0).map(|cell| (cell.number, cell.is_formula)),
            Some((Some(10.0), true))
        );
        // the spill cell has no formula of its own: reported, never patched
        let spill = find_stale_spill(&result, 2, 0).expect("A3 should be a stale spill cell");
        assert_eq!((spill.number, spill.is_formula), (Some(10.0), false));
    }

    /// A dynamic (spilling) array formula is imported the same way: its anchor
    /// carries a cached value the audit can compare, its spilled cells do not.
    #[test]
    fn reports_a_stale_dynamic_array_spill_cell() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("dynamic.xlsx");
        let rows = concat!(
            r#"<row r="2"><c r="A2" cm="1"><f t="array" ref="A2:A3">_xlfn.SEQUENCE(2)</f><v>1</v></c></row>"#,
            r#"<row r="3"><c r="A3" cm="1"><v>99</v></c></row>"#
        );
        write_sheet_package(&path, false, &worksheet("A2:A3", rows));

        let mut cache = RecalcCache::new();
        let result = recalc_cells_reporting(&mut cache, &path, &[], &[], true).unwrap();
        // the anchor's first value agrees with the file
        assert!(find_stale(&result, 1, 0).is_none());
        let spill = find_stale_spill(&result, 2, 0).expect("A3 should be a stale spill cell");
        assert_eq!((spill.number, spill.is_formula), (Some(2.0), false));
    }
}

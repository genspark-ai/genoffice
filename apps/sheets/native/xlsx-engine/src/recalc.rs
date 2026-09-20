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
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

use ironcalc::base::Model;
use ironcalc::base::types::CellType;
use ironcalc::import::load_from_xlsx;
use serde::{Deserialize, Serialize};

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
    /// The engine typed the result as an error (`#DIV/0!`), as opposed to a
    /// formula whose text result merely spells one (`="#N/A"`).
    pub is_error: bool,
    pub is_formula: bool,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecalcResult {
    pub cells: Vec<RecalcCell>,
    /// True when a resident model served this request (no file re-import).
    pub cached: bool,
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
    path.canonicalize().unwrap_or_else(|_| path.to_path_buf())
}

pub fn recalc_cells(
    cache: &mut RecalcCache,
    path: &Path,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
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
    let cached = resident.is_some();

    // IronCalc's importer is strict and can panic on non-Excel producers
    // (missing cellStyles, whitespace nodes); contain that to this request.
    let (entry, cells) = catch_unwind(AssertUnwindSafe(|| {
        run(resident, path, mtime, size, edits, reads)
    }))
    .map_err(|_| {
        SidecarError::Workbook("The formula engine could not process this workbook.".into())
    })??;

    cache.tick += 1;
    let mut entry = entry;
    entry.last_used = cache.tick;
    cache.entries.insert(key, entry);
    cache.evict_beyond_cap();

    Ok(RecalcResult { cells, cached })
}

fn run(
    resident: Option<ResidentModel>,
    path: &Path,
    mtime: SystemTime,
    size: u64,
    edits: &[RecalcEdit],
    reads: &[RecalcRead],
) -> Result<(ResidentModel, Vec<RecalcCell>), SidecarError> {
    let mut entry = match resident {
        Some(entry) => entry,
        None => {
            let path_text = path.to_str().ok_or_else(|| {
                SidecarError::InvalidRequest("Workbook path is not valid UTF-8.".into())
            })?;
            let mut model = load_from_xlsx(path_text, "en", "UTC", "en").map_err(|error| {
                SidecarError::Workbook(format!("Formula engine import failed: {error}"))
            })?;
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
                // Known engine gaps where the file's cached value beats the
                // error: CELL("filename") is unimplemented (#175) and RATE's
                // Newton solver dies near -100% where Excel converges (#185).
                if formatted.starts_with('#') {
                    if let Some(text) = &formula {
                        let upper = text.to_uppercase();
                        if upper.contains("CELL(") && upper.contains("\"FILENAME\"") {
                            continue;
                        }
                        if formatted == "#NUM!" && upper.contains("RATE(") {
                            continue;
                        }
                    }
                }
                let number = raw_number(&entry.model, sheet, row_1, column_1);
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
                    is_error,
                    is_formula,
                });
            }
        }
    }
    Ok((entry, cells))
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
    use ironcalc::base::expressions::lexer::LexerMode;
    use ironcalc::base::expressions::parser::{Node, new_parser_english};
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
            if !matches!(parsed.get(index), Some(Node::ParseErrorKind { .. })) {
                continue;
            }
            let Some(quoted) = quote_bare_sheet_names(formula, &names) else {
                continue;
            };
            let node = parser.parse(&quoted, &context);
            if matches!(node, Node::ParseErrorKind { .. }) {
                continue;
            }
            parsed[index] = node;
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
    use ironcalc::base::types::Cell;
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
        for (row, columns) in &worksheet.sheet_data {
            for (column, cell) in columns {
                let (formula, value) = match cell {
                    Cell::CellFormulaNumber { f, v, .. } => (*f, PinnedValue::Number(*v)),
                    Cell::CellFormulaString { f, v, .. } => (*f, PinnedValue::Text(v.clone())),
                    Cell::CellFormulaBoolean { f, v, .. } => (*f, PinnedValue::Bool(*v)),
                    _ => continue,
                };
                if !matches!(
                    parsed.get(formula as usize),
                    Some(Node::ParseErrorKind { .. })
                ) {
                    continue;
                }
                let text = worksheet
                    .shared_formulas
                    .get(formula as usize)
                    .map(String::as_str)
                    .unwrap_or_default();
                if references_a_sheet_of_this_workbook(text, &names) {
                    continue;
                }
                pins.push((sheet as u32, *row, *column, value));
            }
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
    names.iter().any(|name| {
        !name.is_empty()
            && formula
                .match_indices(&format!("{name}!"))
                .any(|(at, _)| at == 0 || !continues_a_sheet_name(&formula[..at]))
    })
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
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><f>[1]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>A1+1</f><v>85</v></c><c r="C1" t="str"><f>'[1]Sheet1'!A1&amp;" units"</f><v>42 units</v></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:A2)</f><v>89</v></c></row></sheetData></worksheet>"#,
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
                    end_row: 1,
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
    }
}

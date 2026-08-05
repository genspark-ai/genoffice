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
use ironcalc::import::load_from_xlsx;
use serde::{Deserialize, Serialize};

use crate::{CellRange, SidecarError};

pub const MAX_RECALC_EDITS: usize = 10_000;
pub const MAX_RECALC_READ_CELLS: usize = 20_000;
/// Resident models are large (the whole workbook's cell graph); the sidecar
/// serves one document, so two covers the active file plus one recently
/// closed-and-reopened neighbour.
const MAX_RESIDENT_MODELS: usize = 2;

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

    /// Drop the model for a path whose bytes are about to change (save).
    pub fn purge(&mut self, path: &Path) {
        self.entries.remove(path);
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
    }
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

    // Take the entry out while working on it: a panic or error mid-apply
    // leaves the model in an unknown state, and a removed entry can never be
    // reused by the next request.
    let resident = cache.entries.remove(path).filter(|entry| {
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
    cache.entries.insert(path.to_path_buf(), entry);
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
            let model = load_from_xlsx(path_text, "en", "UTC", "en").map_err(|error| {
                SidecarError::Workbook(format!("Formula engine import failed: {error}"))
            })?;
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
                cells.push(RecalcCell {
                    sheet: read.sheet.clone(),
                    row: row as u32,
                    column: column as u32,
                    formatted,
                    number,
                    is_formula,
                });
            }
        }
    }
    Ok((entry, cells))
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
        // the oldest was evicted, the newest survives
        assert!(!cache.entries.contains_key(&paths[0]));
        assert!(cache.entries.contains_key(&paths[2]));
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
}

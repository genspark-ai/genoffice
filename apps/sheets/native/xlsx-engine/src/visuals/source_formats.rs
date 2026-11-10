//! Number formats of the cells a chart's `c:f` ranges point at. Excel's
//! `c:numFmt sourceLinked="1"` means "use the source cells' format"; the
//! attribute's own formatCode is a stale copy some writers (Numbers) never
//! refresh, and their numCache carries no formatCode either.

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::BufReader;

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::xml_util::{attribute_value, zip_entry};

use super::CellStyle;

// Counts how many times a worksheet part is opened to resolve source-format
// lookups, so tests can assert the one-pass-per-worksheet contract. Thread
// local because the test harness runs tests in parallel and other tests open
// worksheets of their own.
#[cfg(test)]
thread_local! {
    static SHEET_PASSES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
pub(crate) fn count_sheet_pass() {
    SHEET_PASSES.with(|passes| passes.set(passes.get() + 1));
}

#[cfg(not(test))]
pub(crate) fn count_sheet_pass() {}

#[cfg(test)]
pub(crate) fn sheet_passes() -> usize {
    SHEET_PASSES.with(std::cell::Cell::get)
}

#[cfg(test)]
pub(crate) fn reset_sheet_passes() {
    SHEET_PASSES.with(|passes| passes.set(0));
}

#[derive(Default)]
pub struct SourceFormats {
    /// (lower-cased sheet name, worksheet part path).
    sheets: Vec<(String, String)>,
    xf_formats: Vec<Option<String>>,
    cache: HashMap<(String, String), Option<String>>,
}

impl SourceFormats {
    pub fn new(sheets: Vec<(String, String)>, styles: &[CellStyle]) -> Self {
        Self {
            sheets: sheets
                .into_iter()
                .map(|(name, path)| (name.to_lowercase(), path))
                .collect(),
            xf_formats: styles
                .iter()
                .map(|style| style.number_format.clone())
                .collect(),
            cache: HashMap::new(),
        }
    }

    /// Number format of the first cell of `'Sheet'!$B$4:$C$8` — the cell
    /// Excel takes a linked format from. General and text (`@`) formats
    /// resolve to None so callers fall through to the next candidate.
    pub fn first_cell_format(
        &mut self,
        archive: &mut ZipArchive<File>,
        reference: &str,
    ) -> Option<String> {
        let (sheet_name, cell) = split_first_cell(reference)?;
        let path = self.path_for_sheet(&sheet_name)?;
        let key = (path, cell);
        if let Some(cached) = self.cache.get(&key) {
            return cached.clone();
        }
        let format = self.format_of(
            cell_style_index(archive, &key.0, &key.1)
                .and_then(|index| self.xf_formats.get(index).cloned().flatten()),
        );
        self.cache.insert(key, format.clone());
        format
    }

    /// Resolve every referenced first cell up front, one streaming pass per
    /// worksheet. `first_cell_format` reopens the sheet part and scans from
    /// byte zero on a cache miss, so a workbook whose charts carry many
    /// source-linked series used to rescan the same large worksheet once per
    /// referenced cell.
    pub fn prefetch(&mut self, archive: &mut ZipArchive<File>, references: &[String]) {
        // (worksheet path, cells still wanted) grouped per sheet, so each part
        // is opened at most once no matter how many charts point into it.
        let mut wanted: HashMap<String, HashSet<String>> = HashMap::new();
        for reference in references {
            let Some((sheet_name, cell)) = split_first_cell(reference) else {
                continue;
            };
            let Some(path) = self.path_for_sheet(&sheet_name) else {
                continue;
            };
            if self.cache.contains_key(&(path.clone(), cell.clone())) {
                continue;
            }
            if row_of(&cell).is_none() {
                continue;
            }
            wanted.entry(path).or_default().insert(cell);
        }
        for (path, cells) in wanted {
            self.resolve_cells(archive, &path, &cells);
        }
    }

    fn path_for_sheet(&self, sheet_name: &str) -> Option<String> {
        let wanted = sheet_name.to_lowercase();
        self.sheets
            .iter()
            .find(|(name, _)| *name == wanted)
            .map(|(_, path)| path.clone())
    }

    /// General and text (`@`) resolve to None so callers fall through to the
    /// next candidate.
    fn format_of(&self, format: Option<String>) -> Option<String> {
        format.filter(|format| format != "General" && format != "@")
    }

    /// One pass over the sheet part, collecting the style index of every
    /// wanted cell. Rows are emitted in ascending order, so the pass stops
    /// after the last wanted row; cells the sheet does not define resolve to
    /// None just as a single-cell lookup would.
    fn resolve_cells(
        &mut self,
        archive: &mut ZipArchive<File>,
        path: &str,
        wanted: &HashSet<String>,
    ) {
        if wanted.is_empty() {
            return;
        }
        let last_row = wanted.iter().filter_map(|cell| row_of(cell)).max();
        let Some(last_row) = last_row else {
            return;
        };
        let mut found: HashMap<String, Option<usize>> = HashMap::new();
        if let Ok(entry) = zip_entry(archive, path) {
            count_sheet_pass();
            let mut reader = Reader::from_reader(BufReader::new(entry));
            let mut buffer = Vec::new();
            loop {
                let event = match reader.read_event_into(&mut buffer) {
                    Ok(event) => event,
                    Err(_) => break,
                };
                match event {
                    Event::Start(element) | Event::Empty(element)
                        if element.local_name().as_ref() == b"row" =>
                    {
                        let row = attribute_value(&reader, &element, b"r")
                            .ok()
                            .flatten()
                            .and_then(|value| value.parse::<u64>().ok());
                        if row.is_some_and(|row| row > last_row) {
                            break;
                        }
                    }
                    Event::Start(element) | Event::Empty(element)
                        if element.local_name().as_ref() == b"c" =>
                    {
                        let Ok(Some(address)) = attribute_value(&reader, &element, b"r") else {
                            buffer.clear();
                            continue;
                        };
                        if !wanted.contains(&address) {
                            buffer.clear();
                            continue;
                        }
                        let index = match attribute_value(&reader, &element, b"s").ok().flatten() {
                            Some(index) => index.parse().ok(),
                            None => Some(0),
                        };
                        found.insert(address, index);
                        if found.len() == wanted.len() {
                            break;
                        }
                    }
                    Event::End(element) if element.local_name().as_ref() == b"sheetData" => break,
                    Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }
        }
        for cell in wanted {
            let format = found
                .get(cell)
                .copied()
                .flatten()
                .and_then(|index| self.xf_formats.get(index).cloned().flatten());
            let format = self.format_of(format);
            self.cache.insert((path.to_owned(), cell.clone()), format);
        }
    }
}

/// `'My Sheet'!$B$4:$C$8` → ("My Sheet", "B4"). Unqualified references have
/// no sheet to look in and yield None.
fn split_first_cell(reference: &str) -> Option<(String, String)> {
    let (sheet, range) = reference.trim().rsplit_once('!')?;
    let sheet = sheet
        .strip_prefix('\'')
        .and_then(|inner| inner.strip_suffix('\''))
        .map(|inner| inner.replace("''", "'"))
        .unwrap_or_else(|| sheet.to_owned());
    let first = range
        .split(':')
        .next()?
        .replace('$', "")
        .to_ascii_uppercase();
    let letters = first.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letters == 0
        || letters == first.len()
        || !first[letters..].bytes().all(|b| b.is_ascii_digit())
    {
        return None;
    }
    Some((sheet, first))
}

fn row_of(cell: &str) -> Option<u64> {
    cell.trim_start_matches(|c: char| c.is_ascii_alphabetic())
        .parse()
        .ok()
}

/// Streams the sheet part until the wanted cell (or past its row).
fn cell_style_index(archive: &mut ZipArchive<File>, path: &str, cell: &str) -> Option<usize> {
    let target_row = row_of(cell)?;
    let entry = zip_entry(archive, path).ok()?;
    count_sheet_pass();
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).ok()? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"row" =>
            {
                let row = attribute_value(&reader, &element, b"r")
                    .ok()
                    .flatten()
                    .and_then(|value| value.parse::<u64>().ok());
                if row.is_some_and(|row| row > target_row) {
                    return None;
                }
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"c" =>
            {
                if attribute_value(&reader, &element, b"r")
                    .ok()
                    .flatten()
                    .as_deref()
                    == Some(cell)
                {
                    return match attribute_value(&reader, &element, b"s").ok().flatten() {
                        Some(index) => index.parse().ok(),
                        None => Some(0),
                    };
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"sheetData" => return None,
            Event::Eof => return None,
            _ => {}
        }
        buffer.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn splits_sheet_qualified_first_cells() {
        assert_eq!(
            split_first_cell("'Table and Chart'!$B$4:$C$8"),
            Some(("Table and Chart".into(), "B4".into()))
        );
        assert_eq!(
            split_first_cell("'It''s'!$a$1"),
            Some(("It's".into(), "A1".into()))
        );
        assert_eq!(
            split_first_cell("Data!C3:C9"),
            Some(("Data".into(), "C3".into()))
        );
        assert_eq!(split_first_cell("$B$4:$C$8"), None);
        assert_eq!(split_first_cell("Data!Table1[Col]"), None);
    }

    fn style(format: Option<&str>) -> CellStyle {
        CellStyle {
            number_format: format.map(ToOwned::to_owned),
            ..CellStyle::default()
        }
    }

    /// Four referenced cells, one pass; the exact-cell lookups that follow must
    /// not reopen the sheet, and each cell keeps its own format.
    #[test]
    fn prefetch_resolves_every_referenced_cell_in_one_pass() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row><row r="5"><c r="B5" s="2"><v>2</v></c></row><row r="9"><c r="C9" s="3"><v>3</v></c></row><row r="40"><c r="D40" s="1"><v>4</v></c></row><row r="41"><c r="E41"><v>5</v></c></row></sheetData></worksheet>"#;
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("fixture.xlsx");
        {
            let mut writer = zip::ZipWriter::new(std::fs::File::create(&path).unwrap());
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            writer
                .start_file("xl/worksheets/sheet1.xml", options)
                .unwrap();
            std::io::Write::write_all(&mut writer, sheet.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        let mut archive = ZipArchive::new(std::fs::File::open(&path).unwrap()).unwrap();
        let mut formats = SourceFormats::new(
            vec![("Data".into(), "xl/worksheets/sheet1.xml".into())],
            &[
                style(None),
                style(Some("#,##0")),
                style(Some("0.00%")),
                style(Some("@")),
            ],
        );
        let references: Vec<String> = [
            // Duplicates and an unresolvable reference must not cost extra passes.
            "Data!$B$5:$B$8",
            "Data!$B$5:$B$8",
            "Data!$C$9",
            "Data!$A$1",
            "Data!$D$40",
            "Missing!$A$1",
        ]
        .iter()
        .map(|value| (*value).to_owned())
        .collect();
        reset_sheet_passes();
        formats.prefetch(&mut archive, &references);
        assert_eq!(sheet_passes(), 1);
        // Exact-cell cache hits: no further worksheet pass.
        assert_eq!(
            formats.first_cell_format(&mut archive, "Data!$B$5:$B$8"),
            Some("0.00%".into())
        );
        assert_eq!(sheet_passes(), 1);
        assert_eq!(
            formats.first_cell_format(&mut archive, "Data!$C$9"),
            None,
            "a text (@) source format falls through to the next candidate"
        );
        assert_eq!(
            formats.first_cell_format(&mut archive, "Data!$A$1"),
            Some("#,##0".into())
        );
        assert_eq!(sheet_passes(), 1);
        // A reference the pre-pass never saw still resolves, on its own pass.
        assert_eq!(formats.first_cell_format(&mut archive, "Data!$E$41"), None);
        assert_eq!(sheet_passes(), 2);
    }
}

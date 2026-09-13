//! Number formats of the cells a chart's `c:f` ranges point at. Excel's
//! `c:numFmt sourceLinked="1"` means "use the source cells' format"; the
//! attribute's own formatCode is a stale copy some writers (Numbers) never
//! refresh, and their numCache carries no formatCode either.

use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::xml_util::{attribute_value, zip_entry};

use super::CellStyle;

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
        let path = self
            .sheets
            .iter()
            .find(|(name, _)| *name == sheet_name.to_lowercase())
            .map(|(_, path)| path.clone())?;
        let key = (path, cell);
        if let Some(cached) = self.cache.get(&key) {
            return cached.clone();
        }
        let format = cell_style_index(archive, &key.0, &key.1)
            .and_then(|index| self.xf_formats.get(index).cloned().flatten())
            .filter(|format| format != "General" && format != "@");
        self.cache.insert(key, format.clone());
        format
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
    use super::split_first_cell;

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
}

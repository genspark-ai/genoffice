//! Legacy-format import: reads an .xls or .ods (anything calamine understands)
//! and writes a minimal .xlsx with values, formulas, and date number formats.
//! Styles beyond date formats are not carried over — the converted file is a
//! fresh workbook, not a byte-preserving edit. ODF arrives in its own reference
//! dialect, so formulas go through `odf::translate_formula` first and any that
//! cannot be proven correct keep only their cached result.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use calamine::{Data, Reader, open_workbook_auto};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::SidecarError;

pub struct ConvertResult {
    pub sheets: usize,
    pub cells: usize,
}

/// Why the reader refused a file, when the refusal is about the container and
/// a friendlier sentence helps. `None` keeps the reader's own error.
///
/// Legacy workbooks are OLE2/CFB containers; a file that starts with anything
/// else was almost certainly renamed rather than exported. The common real
/// case is an XML spreadsheet (SpreadsheetML) saved as `.xls` — Excel sniffs
/// the content and opens it fine, so the file "works in Excel" and then
/// surfaces here as a raw `Invalid OLE (signature ...)` error.
fn unreadable_container_hint(source: &Path) -> Option<String> {
    let mut head = [0u8; 1024];
    let read = File::open(source).ok()?.read(&mut head).ok()?;
    let head = &head[..read];
    if head.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
        // a real CFB container: whatever is wrong inside is the reader's to explain
        return None;
    }
    let text = String::from_utf8_lossy(head);
    let trimmed = text.trim_start_matches('\u{FEFF}').trim_start();
    if !trimmed.starts_with("<?xml") && !trimmed.starts_with('<') {
        return None;
    }
    if text.contains("urn:schemas-microsoft-com:office:spreadsheet")
        || text.contains("<Workbook")
        || text.contains("<ss:Worksheet")
    {
        return Some(
            "This file holds an XML spreadsheet (SpreadsheetML 2003) saved with a legacy \
             workbook name. GenOffice cannot convert that format: open the file in Excel \
             or LibreOffice and re-save it as .xlsx."
                .into(),
        );
    }
    let first_line = trimmed.lines().next().unwrap_or("").trim();
    let first_line = if first_line.chars().count() > 60 {
        let cut: String = first_line.chars().take(60).collect();
        format!("{cut}…")
    } else {
        first_line.to_string()
    };
    Some(format!(
        "This file is not a workbook container: its content is text or XML rather than \
         an Excel/ODF package, and it begins with \"{first_line}\"."
    ))
}

pub fn convert_to_xlsx(source: &Path, target: &Path) -> Result<ConvertResult, SidecarError> {
    let mut workbook =
        open_workbook_auto(source).map_err(|error| match unreadable_container_hint(source) {
            Some(hint) => SidecarError::Workbook(hint),
            None => SidecarError::Workbook(format!("Unable to read the workbook: {error}")),
        })?;
    let names: Vec<String> = workbook.sheet_names().to_vec();
    if names.is_empty() {
        return Err(SidecarError::Workbook("The workbook has no sheets.".into()));
    }

    let mut sheet_xmls: Vec<String> = Vec::new();
    let mut cells = 0usize;
    for name in &names {
        let range = workbook
            .worksheet_range(name)
            .map_err(|error| SidecarError::Workbook(format!("Sheet {name}: {error}")))?;
        let formulas = workbook.worksheet_formula(name).ok();
        let mut formula_map: HashMap<(u32, u32), String> = HashMap::new();
        if let Some(formula_range) = formulas {
            let (start_row, start_col) = formula_range.start().unwrap_or((0, 0));
            for (row, column, formula) in formula_range.used_cells() {
                if formula.is_empty() {
                    continue;
                }
                // An ODF formula with no provably-correct A1 spelling keeps only
                // its cached result: `of:=[.A1]` is not an OOXML token sequence,
                // and Excel answers one with a repair prompt.
                let Some(translated) = crate::odf::translate_formula(formula) else {
                    continue;
                };
                formula_map.insert(
                    (start_row + row as u32, start_col + column as u32),
                    translated,
                );
            }
        }
        let (xml, sheet_cells) = worksheet_xml(&range, &formula_map);
        cells += sheet_cells;
        sheet_xmls.push(xml);
    }

    let file = File::create(target)?;
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut add = |name: &str, content: &str| -> Result<(), SidecarError> {
        writer.start_file(name, options)?;
        writer.write_all(content.as_bytes())?;
        Ok(())
    };

    add("[Content_Types].xml", &content_types_xml(names.len()))?;
    add("_rels/.rels", ROOT_RELS)?;
    add("xl/workbook.xml", &workbook_xml(&names))?;
    add(
        "xl/_rels/workbook.xml.rels",
        &workbook_rels_xml(names.len()),
    )?;
    add("xl/styles.xml", STYLES_XML)?;
    for (index, xml) in sheet_xmls.iter().enumerate() {
        add(&format!("xl/worksheets/sheet{}.xml", index + 1), xml)?;
    }
    writer.finish()?.sync_all()?;
    Ok(ConvertResult {
        sheets: names.len(),
        cells,
    })
}

fn worksheet_xml(
    range: &calamine::Range<Data>,
    formulas: &HashMap<(u32, u32), String>,
) -> (String, usize) {
    // Rows carrying values plus rows carrying only formulas.
    let mut rows: HashMap<u32, Vec<(u32, String)>> = HashMap::new();
    let mut cells = 0usize;
    let (start_row, start_col) = range.start().unwrap_or((0, 0));
    for (row, column, value) in range.used_cells() {
        let absolute = (start_row + row as u32, start_col + column as u32);
        let formula = formulas.get(&absolute).map(String::as_str);
        if let Some(cell) = cell_xml(absolute, value, formula) {
            cells += 1;
            rows.entry(absolute.0).or_default().push((absolute.1, cell));
        }
    }
    for (position, formula) in formulas {
        let covered = range
            .get_value((position.0, position.1))
            .is_some_and(|v| *v != Data::Empty);
        if !covered {
            cells += 1;
            rows.entry(position.0).or_default().push((
                position.1,
                format!(
                    r#"<c r="{}"><f>{}</f></c>"#,
                    cell_reference(position.0, position.1),
                    escape_xml(formula),
                ),
            ));
        }
    }

    let mut row_numbers: Vec<u32> = rows.keys().copied().collect();
    row_numbers.sort_unstable();
    let mut body = String::new();
    for row in row_numbers {
        let mut line = rows.remove(&row).unwrap_or_default();
        line.sort_unstable_by_key(|(column, _)| *column);
        body.push_str(&format!(r#"<row r="{}">"#, row + 1));
        for (_, cell) in line {
            body.push_str(&cell);
        }
        body.push_str("</row>");
    }

    let dimension = match (range.start(), range.end()) {
        (Some(start), Some(end)) => format!(
            "{}:{}",
            cell_reference(start.0, start.1),
            cell_reference(end.0, end.1),
        ),
        _ => "A1:A1".into(),
    };
    (
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/><sheetData>{body}</sheetData></worksheet>"#,
        ),
        cells,
    )
}

fn cell_xml(position: (u32, u32), value: &Data, formula: Option<&str>) -> Option<String> {
    let reference = cell_reference(position.0, position.1);
    let formula_xml = formula
        .map(|text| format!("<f>{}</f>", escape_xml(text)))
        .unwrap_or_default();
    let cell = match value {
        Data::Empty => {
            if formula.is_none() {
                return None;
            }
            format!(r#"<c r="{reference}">{formula_xml}</c>"#)
        }
        Data::String(text) => format!(
            r#"<c r="{reference}" t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
            escape_xml(text),
        ),
        Data::Float(number) => format!(r#"<c r="{reference}">{formula_xml}<v>{number}</v></c>"#),
        Data::Int(number) => format!(r#"<c r="{reference}">{formula_xml}<v>{number}</v></c>"#),
        Data::Bool(flag) => format!(
            r#"<c r="{reference}" t="b">{formula_xml}<v>{}</v></c>"#,
            if *flag { 1 } else { 0 },
        ),
        Data::DateTime(datetime) => date_cell(&reference, &formula_xml, datetime.as_f64()),
        Data::Error(error) => format!(
            r#"<c r="{reference}" t="e">{formula_xml}<v>{}</v></c>"#,
            escape_xml(&error.to_string()),
        ),
        // ODF keeps a date as ISO text (`office:date-value="2026-03-05"`) rather
        // than as a serial, so the text form is turned back into one here: left
        // as text, every imported date would sort and calculate as a string.
        Data::DateTimeIso(text) => match iso_serial(text) {
            Some(serial) => date_cell(&reference, &formula_xml, serial),
            None => text_cell(&reference, &formula_xml, text),
        },
        Data::DurationIso(text) => text_cell(&reference, &formula_xml, text),
    };
    Some(cell)
}

/// A date cell: whole serials take the short-date format, timed ones the
/// date+time one (xf 1 / xf 2 of STYLES_XML).
fn date_cell(reference: &str, formula_xml: &str, serial: f64) -> String {
    let style = if serial.fract() == 0.0 { 1 } else { 2 };
    format!(r#"<c r="{reference}" s="{style}">{formula_xml}<v>{serial}</v></c>"#)
}

fn text_cell(reference: &str, formula_xml: &str, text: &str) -> String {
    format!(
        r#"<c r="{reference}" t="inlineStr">{formula_xml}<is><t xml:space="preserve">{}</t></is></c>"#,
        escape_xml(text),
    )
}

/// Excel serial for the ISO forms ODF writes (`2026-03-05`, optionally followed
/// by `T10:20:30` and a fraction). `None` keeps the cell a string instead of
/// guessing — including below 1900-03-01, where Excel's phantom 1900 leap day
/// makes a computed serial one too low.
fn iso_serial(text: &str) -> Option<f64> {
    let (date, time) = match text.split_once('T') {
        Some((date, time)) => (date, Some(time)),
        None => (text, None),
    };
    let mut fields = date.split('-');
    let year: i64 = fields.next()?.parse().ok()?;
    let month: i64 = fields.next()?.parse().ok()?;
    let day: i64 = fields.next()?.parse().ok()?;
    if fields.next().is_some()
        || !(1900..=9999).contains(&year)
        || !(1..=12).contains(&month)
        || day < 1
        || day > i64::from(days_in_month(year, month))
    {
        return None;
    }
    // Serial 1 is 1900-01-01 for Excel's own arithmetic, but the 1900-02-29 it
    // never had puts every later date one above the plain day count, which the
    // 1899-12-30 epoch absorbs.
    let mut serial = (days_from_civil(year, month, day) - days_from_civil(1899, 12, 30)) as f64;
    if let Some(time) = time {
        let mut parts = time.split(':');
        let hours: f64 = unsigned_number(parts.next()?)?;
        let minutes: f64 = unsigned_number(parts.next()?)?;
        let seconds_text = parts.next()?;
        if parts.next().is_some() {
            return None;
        }
        let (whole, fraction) = match seconds_text.split_once('.') {
            Some((whole, fraction)) => (
                unsigned_number(whole)?,
                unsigned_number(fraction)? / 10f64.powi(fraction.chars().count() as i32),
            ),
            None => (unsigned_number(seconds_text)?, 0.0),
        };
        if hours > 23.0 || minutes > 59.0 || whole > 60.0 {
            return None;
        }
        serial += (hours * 3600.0 + minutes * 60.0 + whole + fraction) / 86_400.0;
    }
    if serial < 61.0 || serial > 2_958_465.0 {
        return None;
    }
    Some(serial)
}

/// Digits only, no sign and no exponent: `"05"` parses to 5 while `"+5"`,
/// `"1e3"` and `"30Z"` are refused.
fn unsigned_number(text: &str) -> Option<f64> {
    if text.is_empty() || !text.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    text.parse().ok()
}

fn days_in_month(year: i64, month: i64) -> u32 {
    const LENGTHS: [u32; 12] = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let leap = month == 2 && (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
    if leap {
        29
    } else {
        LENGTHS[(month - 1) as usize]
    }
}

/// Days from the civil calendar epoch 1970-01-01 (Howard Hinnant's algorithm).
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let shifted = if month <= 2 { year - 1 } else { year };
    let era = if shifted >= 0 { shifted } else { shifted - 399 } / 400;
    let year_of_era = shifted - era * 400;
    let day_of_year = (153 * (if month > 2 { month - 3 } else { month + 9 }) + 2) / 5 + day - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

fn cell_reference(row: u32, column: u32) -> String {
    let mut letters = String::new();
    let mut remaining = column + 1;
    while remaining > 0 {
        remaining -= 1;
        letters.insert(0, char::from(b'A' + (remaining % 26) as u8));
        remaining /= 26;
    }
    format!("{letters}{}", row + 1)
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn content_types_xml(sheet_count: usize) -> String {
    let overrides: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        ))
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{overrides}</Types>"#,
    )
}

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

fn workbook_xml(names: &[String]) -> String {
    let sheets: String = names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                escape_xml(name),
                index + 1,
                index + 1,
            )
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#,
    )
}

fn workbook_rels_xml(sheet_count: usize) -> String {
    let mut relationships: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>"#,
        ))
        .collect();
    relationships.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        sheet_count + 1,
    ));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#,
    )
}

/// xf 1 = short date (numFmt 14), xf 2 = date+time (numFmt 22).
const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    fn read_entry(path: &Path, name: &str) -> String {
        let mut archive = zip::ZipArchive::new(File::open(path).unwrap()).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut content = String::new();
        entry.read_to_string(&mut content).unwrap();
        content
    }

    #[test]
    fn converts_a_calamine_readable_workbook_into_minimal_xlsx() {
        // calamine reads xlsx through the same auto-detect path as xls, so a
        // generated xlsx fixture exercises the converter end to end.
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        convert_fixture(&source);
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert!(result.cells >= 4);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(
            r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">Name &amp; Co</t></is></c>"#
        ));
        assert!(sheet.contains(r#"<c r="B1"><v>42</v></c>"#));
        assert!(sheet.contains(r#"<c r="B2"><f>B1*2</f>"#));
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Data" sheetId="1" r:id="rId1"/>"#));
    }

    #[test]
    fn iso_dates_become_serials_and_suspicious_ones_stay_text() {
        assert_eq!(iso_serial("2026-03-05"), Some(46_086.0));
        assert_eq!(iso_serial("1900-03-01"), Some(61.0));
        assert_eq!(iso_serial("2026-03-05T12:00:00"), Some(46_086.5));
        let timed = iso_serial("2026-03-05T06:30:15.5").unwrap();
        assert!((timed - 46_086.271_012_731_48).abs() < 1e-9);
        // Excel's phantom 1900 leap day sits under the 61 line.
        assert_eq!(iso_serial("1900-02-28"), None);
        assert_eq!(iso_serial("2026-02-30"), None);
        assert_eq!(iso_serial("2026-13-01"), None);
        assert_eq!(iso_serial("2026-03-05T1:2"), None);
        assert_eq!(iso_serial("2026-03-05T10:20:30Z"), None);
        assert_eq!(iso_serial("PT1H30M"), None);
    }

    #[test]
    fn converts_an_odf_workbook_translating_formulas_and_dates() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("probe.ods");
        ods_fixture(&source);
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 2);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        // the `of:` reference dialect comes back as A1 semantics
        assert!(sheet.contains(r#"<c r="C1"><f>A1+B1*2</f><v>20</v></c>"#));
        assert!(sheet.contains(r#"<c r="A2"><f>SUM(A1:B1)</f><v>15</v></c>"#));
        assert!(sheet.contains(r#"<c r="B2"><f>Other!B1*2</f><v>100</v></c>"#));
        // a construct with no OOXML spelling keeps its cached result only
        assert!(sheet.contains(r#"<c r="D1"><v>7</v></c>"#));
        assert!(!sheet.contains("$Sales"));
        // office:date-value lands as a dated number, not the text it arrived as
        assert!(sheet.contains(r#"<c r="A3" s="1"><v>46086</v></c>"#));
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Other" sheetId="2" r:id="rId2"/>"#));
    }

    // A workbook saved by LibreOffice itself, committed rather than generated
    // because CI has no LibreOffice. Regenerate with
    /// `soffice --headless --convert-to ods --outdir <dir> <repo>/apps/sheets/native/xlsx-engine/tests/fixtures/cross-sheet-references.fods`
    /// from the committed `.fods` source next to it.
    ///
    /// It pins the dialect a real save produces: every other-sheet reference is
    /// absolute-sheet and apostrophe-quoted (`[$Other.A1]`, `[$'My Sheet'.A1]`,
    /// `SUM([$Data.A1:$Data.B2])`), which a hand-written fixture never produces.
    #[test]
    fn converts_a_libreoffice_saved_workbook_with_absolute_sheet_references() {
        let dir = tempfile::tempdir().unwrap();
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/cross-sheet-references.ods");
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 4);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<c r="C1"><f>Other!A1*2</f><v>200</v></c>"#));
        assert!(sheet.contains(r#"<c r="D1"><f>'My Sheet'!A1+1</f><v>11</v></c>"#));
        assert!(sheet.contains(r#"<c r="E1"><f>SUM(Data!A1:B2)</f><v>30</v></c>"#));
        assert!(sheet.contains(r#"<c r="F1"><f>SUM(Data!$A$1:$B$2)</f><v>30</v></c>"#));
        // the self-sheet form LibreOffice shares with every other ODF writer
        assert!(sheet.contains(r#"<c r="G1"><f>SUM(A1:B1)</f><v>15</v></c>"#));
        // no reference keeps ODF's `$` sheet marker or its doubled quoting
        assert!(!sheet.contains("'$"), "{sheet}");
        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="My Sheet""#));
    }

    /// A minimal ODF spreadsheet: the `mimetype` entry first, then a
    /// `content.xml` kept free of inter-element whitespace, which the ODF
    /// reader treats as cell text.
    fn ods_fixture(path: &Path) {
        ods_fixture_with_mimetype(path, "application/vnd.oasis.opendocument.spreadsheet");
    }

    /// `.ots` templates carry the base mimetype with a `-template` suffix; the
    /// ODF reader compares only the first 46 bytes, which keeps accepting them —
    /// this pins that behavior, the .ots import route depends on it.
    #[test]
    fn converts_an_odf_spreadsheet_template() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("probe.ots");
        ods_fixture_with_mimetype(
            &source,
            "application/vnd.oasis.opendocument.spreadsheet-template",
        );
        let target = dir.path().join("converted.xlsx");

        let result = convert_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 2);
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<c r="C1"><f>A1+B1*2</f><v>20</v></c>"#));
    }

    fn ods_fixture_with_mimetype(path: &Path, mimetype: &str) {
        const NAMESPACES: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2""#;
        let cell = |attributes: &str, text: &str| {
            format!(r#"<table:table-cell {attributes}><text:p>{text}</text:p></table:table-cell>"#)
        };
        let probe = [
            "<table:table table:name=\"Probe\"><table:table-row>",
            &cell(r#"office:value-type="float" office:value="10""#, "10"),
            &cell(r#"office:value-type="float" office:value="5""#, "5"),
            &cell(
                r#"table:formula="of:=[.A1]+[.B1]*2" office:value-type="float" office:value="20""#,
                "20",
            ),
            &cell(
                r#"table:formula="of:=$Sales" office:value-type="float" office:value="7""#,
                "7",
            ),
            "</table:table-row><table:table-row>",
            &cell(
                r#"table:formula="of:=SUM([.A1:.B1])" office:value-type="float" office:value="15""#,
                "15",
            ),
            &cell(
                r#"table:formula="of:=[Other.B1]*2" office:value-type="float" office:value="100""#,
                "100",
            ),
            "</table:table-row><table:table-row>",
            &cell(
                r#"office:value-type="date" office:date-value="2026-03-05""#,
                "2026-03-05",
            ),
            "</table:table-row></table:table>",
        ]
        .join("");
        let other = format!(
            r#"<table:table table:name="Other"><table:table-row>{}</table:table-row></table:table>"#,
            cell(r#"office:value-type="float" office:value="50""#, "50"),
        );
        let content = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content {NAMESPACES}><office:body><office:spreadsheet>{probe}{other}</office:spreadsheet></office:body></office:document-content>"#
        );

        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        let mut add = |name: &str, body: &str| {
            writer.start_file(name, options).unwrap();
            writer.write_all(body.as_bytes()).unwrap();
        };
        add("mimetype", mimetype);
        add(
            "META-INF/manifest.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        );
        add("content.xml", &content);
        writer.finish().unwrap();
    }

    fn convert_fixture(path: &Path) {
        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        let mut add = |name: &str, content: &str| {
            writer.start_file(name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        };
        add(
            "[Content_Types].xml",
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        );
        add(
            "_rels/.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        );
        add(
            "xl/workbook.xml",
            r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        );
        add(
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        );
        add(
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name &amp; Co</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2"><f>B1*2</f><v>84</v></c></row></sheetData></worksheet>"#,
        );
        writer.finish().unwrap();
    }

    #[test]
    fn a_renamed_xml_spreadsheet_gets_the_spreadsheetml_hint() {
        // The shape genoffice#2038 reports: an XML spreadsheet (SpreadsheetML 2003)
        // saved under a legacy name. Excel sniffs content over extension and
        // opens it, so it reaches the converter and used to die as a raw
        // "Invalid OLE (signature ...)".
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("report.xls");
        std::fs::write(
            &source,
            r#"<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Data">
  <Table><Row><Cell><Data ss:Type="Number">42</Data></Cell></Row></Table>
 </Worksheet>
</Workbook>
"#,
        )
        .unwrap();
        let Err(error) = convert_to_xlsx(&source, &dir.path().join("out.xlsx")) else {
            panic!("a SpreadsheetML file must not convert");
        };
        let message = error.to_string();
        assert!(message.contains("SpreadsheetML"), "{message}");
        assert!(message.contains("re-save it as .xlsx"), "{message}");
    }

    #[test]
    fn other_text_content_says_what_the_file_is_not() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("page.xls");
        std::fs::write(
            &source,
            "<html><body><table><tr><td>1</td></tr></table></body></html>\n",
        )
        .unwrap();
        let Err(error) = convert_to_xlsx(&source, &dir.path().join("out.xlsx")) else {
            panic!("html is not a workbook");
        };
        let message = error.to_string();
        assert!(message.contains("not a workbook container"), "{message}");
        assert!(message.contains("<html>"), "{message}");
    }

    #[test]
    fn a_real_container_keeps_the_reader_error() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("broken.xls");
        let mut bytes = vec![0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        bytes.extend(vec![0u8; 600]);
        std::fs::write(&source, &bytes).unwrap();
        let Err(error) = convert_to_xlsx(&source, &dir.path().join("out.xlsx")) else {
            panic!("a truncated container must not convert");
        };
        let message = error.to_string();
        assert!(message.contains("Unable to read the workbook"), "{message}");
    }
}

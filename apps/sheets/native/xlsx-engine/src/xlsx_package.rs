//! The .xlsx package every legacy-format import writes into. Each dialect reads
//! its own container — an OLE2 workbook, an ODF zip, an XML spreadsheet — but
//! they all emit the same OOXML parts, so the writer, the relationship set, and
//! the reference and date helpers live here instead of being restated per
//! dialect. What a dialect owns is only its own cell content.

use std::fs::File;
use std::io::Write;
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::SidecarError;

/// One entry of the workbook's sheet list, in document order.
pub(crate) struct SheetSpec {
    pub(crate) name: String,
    /// Hidden in the source dialect; OOXML spells the same thing `state="hidden"`.
    pub(crate) hidden: bool,
}

/// Writes the whole package. `sheet_xmls` runs in the same order as `sheets`:
/// entry *n* becomes `xl/worksheets/sheet{n}.xml`.
pub(crate) fn write_package(
    target: &Path,
    sheets: &[SheetSpec],
    defined_names: &str,
    sheet_xmls: &[String],
) -> Result<(), SidecarError> {
    let file = File::create(target)?;
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut add = |name: &str, content: &str| -> Result<(), SidecarError> {
        writer.start_file(name, options)?;
        writer.write_all(content.as_bytes())?;
        Ok(())
    };

    add("[Content_Types].xml", &content_types_xml(sheets.len()))?;
    add("_rels/.rels", ROOT_RELS)?;
    add("xl/workbook.xml", &workbook_xml(sheets, defined_names))?;
    add(
        "xl/_rels/workbook.xml.rels",
        &workbook_rels_xml(sheets.len()),
    )?;
    add("xl/styles.xml", STYLES_XML)?;
    for (index, xml) in sheet_xmls.iter().enumerate() {
        add(&format!("xl/worksheets/sheet{}.xml", index + 1), xml)?;
    }
    writer.finish()?.sync_all()?;
    Ok(())
}

/// The `<worksheet>` wrapper: a dimension, then the parts the schema fixes in
/// this order — `cols` before the data, and the dialect's trailing parts
/// (merges, filters) after it.
pub(crate) fn worksheet_document(
    dimension: &str,
    columns: &str,
    sheet_data: &str,
    tail: &str,
) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/>{columns}<sheetData>{sheet_data}</sheetData>{tail}</worksheet>"#,
    )
}

pub(crate) fn cell_reference(row: u32, column: u32) -> String {
    let mut letters = String::new();
    let mut remaining = column + 1;
    while remaining > 0 {
        remaining -= 1;
        letters.insert(0, char::from(b'A' + (remaining % 26) as u8));
        remaining /= 26;
    }
    format!("{letters}{}", row + 1)
}

/// The locked spelling, for a reference the source marked absolute. R1C1's
/// `R1C1` has to arrive as `$A$1`: the plain `A1` would drift the moment the
/// formula was copied somewhere else.
pub(crate) fn absolute_cell_reference(row: u32, column: u32) -> String {
    let mut letters = String::new();
    let mut remaining = column + 1;
    while remaining > 0 {
        remaining -= 1;
        letters.insert(0, char::from(b'A' + (remaining % 26) as u8));
        remaining /= 26;
    }
    format!("${letters}${}", row + 1)
}

pub(crate) fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Excel serial for an ISO date (`2026-03-05`, optionally followed by `T10:20:30`
/// and a fraction). `None` keeps the cell a string instead of guessing —
/// including below 1900-03-01, where Excel's phantom 1900 leap day makes a
/// computed serial one too low.
pub(crate) fn iso_serial(text: &str) -> Option<f64> {
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

/// `defined_names` arrives already spelled as `<definedName>` elements, and
/// goes in after `<sheets>` where the schema wants it.
fn workbook_xml(sheets: &[SheetSpec], defined_names: &str) -> String {
    let entries: String = sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| {
            let state = if sheet.hidden {
                r#" state="hidden""#
            } else {
                ""
            };
            format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"{state}/>"#,
                escape_xml(&sheet.name),
                index + 1,
                index + 1,
            )
        })
        .collect();
    let names = if defined_names.is_empty() {
        String::new()
    } else {
        format!("<definedNames>{defined_names}</definedNames>")
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{entries}</sheets>{names}</workbook>"#,
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

//! Shared-formula expansion (OOXML 18.3.1.40).
//!
//! In a shared-formula group only the master cell carries the formula text
//! (`<f t="shared" ref="E3:H3" si="0">D3+1</f>`); the followers are bare
//! `<f t="shared" si="0"/>`. Excel materializes each follower by shifting the
//! master's relative references by the (row, column) offset between the two
//! cells. Absolute parts (`$`) stay put; string literals and quoted sheet
//! names are never touched.

/// Registry of shared-formula masters within one worksheet part, keyed by
/// `si`. The master always precedes its followers in the part, so a single
/// streaming pass can expand every follower.
#[derive(Default)]
pub struct SharedFormulas {
    masters: std::collections::HashMap<u32, (usize, usize, String)>,
}

impl SharedFormulas {
    /// Record the master cell of group `si` (0-based row/column).
    pub fn register(&mut self, si: u32, row: usize, column: usize, formula: &str) {
        self.masters
            .entry(si)
            .or_insert_with(|| (row, column, formula.to_owned()));
    }

    /// Expand a follower at (row, column): the master's formula with relative
    /// references shifted by the offset. None for unknown groups or when a
    /// shifted reference would leave the sheet.
    pub fn expand(&self, si: u32, row: usize, column: usize) -> Option<String> {
        let (master_row, master_column, formula) = self.masters.get(&si)?;
        translate_shared_formula(
            formula,
            row as i64 - *master_row as i64,
            column as i64 - *master_column as i64,
        )
    }
}

const MAX_ROW: i64 = 1_048_576;
const MAX_COLUMN: i64 = 16_384; // XFD

/// Shift the relative A1-style references in `formula` by the given deltas.
/// Handles plain refs (`B2`, `$A$1`, `Sheet2!C3`), whole-column ranges
/// (`B:B`) and whole-row ranges (`3:3`); skips double-quoted string literals,
/// single-quoted sheet names, and function names (token followed by `(`).
/// Returns None when any shifted reference falls off the sheet.
pub fn translate_shared_formula(
    formula: &str,
    row_delta: i64,
    column_delta: i64,
) -> Option<String> {
    if row_delta == 0 && column_delta == 0 {
        return Some(formula.to_owned());
    }
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'"' || bytes[i] == b'\'' {
            let quote = bytes[i];
            let start = i;
            i += 1;
            while i < bytes.len() {
                if bytes[i] == quote {
                    if bytes.get(i + 1) == Some(&quote) {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            out.push_str(&formula[start..i]);
            continue;
        }
        let start = i;
        let end = formula_name_end(formula, start);
        if end == start {
            let character = formula[i..].chars().next()?;
            out.push(character);
            i += character.len_utf8();
            continue;
        }
        let token = &formula[start..end];
        if bytes.get(end) == Some(&b'(') {
            out.push_str(token);
            i = end;
            continue;
        }
        if bytes.get(end) == Some(&b':') {
            let second_start = end + 1;
            let second_end = formula_name_end(formula, second_start);
            let second = &formula[second_start..second_end];
            let cell_after = bytes.get(second_end) != Some(&b'(');
            if cell_after {
                if let (Some(a), Some(b)) = (
                    shift_axis_token(
                        token,
                        column_delta,
                        MAX_COLUMN,
                        parse_column,
                        column_to_letters,
                    ),
                    shift_axis_token(
                        second,
                        column_delta,
                        MAX_COLUMN,
                        parse_column,
                        column_to_letters,
                    ),
                ) {
                    out.push_str(&a);
                    out.push(':');
                    out.push_str(&b);
                    i = second_end;
                    continue;
                }
                if let (Some(a), Some(b)) = (
                    shift_axis_token(token, row_delta, MAX_ROW, parse_row, |row| row.to_string()),
                    shift_axis_token(second, row_delta, MAX_ROW, parse_row, |row| row.to_string()),
                ) {
                    out.push_str(&a);
                    out.push(':');
                    out.push_str(&b);
                    i = second_end;
                    continue;
                }
            }
        }
        out.push_str(&shift_cell_token(token, row_delta, column_delta)?);
        i = end;
    }
    Some(out)
}

fn formula_name_end(formula: &str, start: usize) -> usize {
    let mut end = start;
    for (offset, character) in formula[start..].char_indices() {
        if !character.is_alphanumeric() && !matches!(character, '_' | '.' | '\\' | '$') {
            break;
        }
        end = start + offset + character.len_utf8();
    }
    end
}

/// `$?letters$?digits` → shifted reference; anything else passes through.
/// None only when the token is a reference and the shift leaves the sheet.
fn shift_cell_token(token: &str, row_delta: i64, column_delta: i64) -> Option<String> {
    let bytes = token.as_bytes();
    let mut i = 0;
    let column_absolute = bytes.first() == Some(&b'$');
    if column_absolute {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let letters = &token[letters_start..i];
    let row_absolute = bytes.get(i) == Some(&b'$');
    if row_absolute {
        i += 1;
    }
    let digits_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    let digits = &token[digits_start..i];
    // Not a cell reference (function/defined name, number literal, leftover):
    // pass through untouched.
    if i != token.len() || letters.is_empty() || letters.len() > 3 || digits.is_empty() {
        return Some(token.to_owned());
    }
    let column = parse_column(letters)?;
    let row = digits.parse::<i64>().ok().filter(|row| *row >= 1)?;
    let column = if column_absolute {
        column
    } else {
        column + column_delta
    };
    let row = if row_absolute { row } else { row + row_delta };
    if !(1..=MAX_COLUMN).contains(&column) || !(1..=MAX_ROW).contains(&row) {
        return None;
    }
    let mut out = String::new();
    if column_absolute {
        out.push('$');
    }
    out.push_str(&column_to_letters(column));
    if row_absolute {
        out.push('$');
    }
    out.push_str(&row.to_string());
    Some(out)
}

/// One side of a whole-column/whole-row range: `$?letters` or `$?digits`.
/// None when the token is not such an axis reference or the shift leaves
/// the sheet (the caller then falls back to plain token handling).
fn shift_axis_token(
    token: &str,
    delta: i64,
    max: i64,
    parse: fn(&str) -> Option<i64>,
    render: fn(i64) -> String,
) -> Option<String> {
    let (absolute, body) = match token.strip_prefix('$') {
        Some(rest) => (true, rest),
        None => (false, token),
    };
    let value = parse(body)?;
    let shifted = if absolute { value } else { value + delta };
    if !(1..=max).contains(&shifted) {
        return None;
    }
    Some(if absolute {
        format!("${}", render(shifted))
    } else {
        render(shifted)
    })
}

fn parse_column(letters: &str) -> Option<i64> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut column = 0i64;
    for byte in letters.bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column * 26 + i64::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    Some(column)
}

fn parse_row(digits: &str) -> Option<i64> {
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    digits.parse::<i64>().ok().filter(|row| *row >= 1)
}

fn column_to_letters(mut column: i64) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        letters.push(b'A' + remainder);
        column = (column - 1) / 26;
    }
    letters.reverse();
    String::from_utf8(letters).unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifts_relative_refs_and_keeps_absolute_parts() {
        assert_eq!(
            translate_shared_formula("D3+$A$1+B$2+$C4", 2, 1).as_deref(),
            Some("E5+$A$1+C$2+$C6")
        );
    }

    #[test]
    fn spec_example_followers() {
        // <f t="shared" ref="E3:H3" si="0">D3+1</f> at E3: F3 → E3+1, H3 → G3+1
        assert_eq!(
            translate_shared_formula("D3+1", 0, 1).as_deref(),
            Some("E3+1")
        );
        assert_eq!(
            translate_shared_formula("D3+1", 0, 3).as_deref(),
            Some("G3+1")
        );
    }

    #[test]
    fn leaves_strings_quoted_sheets_and_functions_alone() {
        assert_eq!(
            translate_shared_formula("\"A1\"&LOG10(A1)&'My A1 Sheet'!B2", 1, 0).as_deref(),
            Some("\"A1\"&LOG10(A2)&'My A1 Sheet'!B3")
        );
        // Unquoted sheet names shift only their cell part.
        assert_eq!(
            translate_shared_formula("Sheet2!A1+Sheet2!$A$1", 1, 1).as_deref(),
            Some("Sheet2!B2+Sheet2!$A$1")
        );
    }

    #[test]
    fn shifts_whole_column_and_row_ranges() {
        assert_eq!(
            translate_shared_formula("SUM(B:B)+SUM($C:$C)", 5, 1).as_deref(),
            Some("SUM(C:C)+SUM($C:$C)")
        );
        assert_eq!(
            translate_shared_formula("SUM(3:4)", 2, 0).as_deref(),
            Some("SUM(5:6)")
        );
    }

    #[test]
    fn escaped_quotes_stay_verbatim() {
        assert_eq!(
            translate_shared_formula("\"say \"\"B1\"\"\"&B1", 1, 0).as_deref(),
            Some("\"say \"\"B1\"\"\"&B2")
        );
    }

    #[test]
    fn numbers_and_names_pass_through() {
        assert_eq!(
            translate_shared_formula("ROUND(1.5,0)+TaxRate+3", 1, 1).as_deref(),
            Some("ROUND(1.5,0)+TaxRate+3")
        );
    }

    #[test]
    fn preserves_unicode_defined_names() {
        let mut shared = SharedFormulas::default();
        shared.register(0, 0, 0, "é”€å”®é¡^M+A1");

        assert_eq!(shared.expand(0, 0, 1).as_deref(), Some("é”€å”®é¡^M+B1"));
    }

    #[test]
    fn does_not_shift_ascii_suffixes_inside_unicode_names() {
        assert_eq!(
            translate_shared_formula("\u{540d}A1+A1", 0, 1).as_deref(),
            Some("\u{540d}A1+B1")
        );
    }

    #[test]
    fn out_of_bounds_shift_fails() {
        assert_eq!(translate_shared_formula("A1", -1, 0), None);
        assert_eq!(translate_shared_formula("A1", 0, -1), None);
        assert_eq!(translate_shared_formula("XFD1", 0, 1), None);
    }

    #[test]
    fn registry_expands_followers_from_the_master() {
        let mut shared = SharedFormulas::default();
        shared.register(0, 2, 4, "D3+1"); // master E3
        assert_eq!(shared.expand(0, 2, 5).as_deref(), Some("E3+1")); // F3
        assert_eq!(shared.expand(0, 2, 7).as_deref(), Some("G3+1")); // H3
        assert_eq!(shared.expand(9, 0, 0), None); // unknown group
    }
}

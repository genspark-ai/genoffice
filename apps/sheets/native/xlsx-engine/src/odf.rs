//! ODF (`of:`) formula translation for the legacy import path.
//!
//! calamine hands back an `.ods` formula in the ODF prefix dialect, whose
//! references (`[.A1]`, `[Sheet.B2]`, `[$'My Sheet'.B2]`) and argument
//! separator (`;`) are not OOXML tokens. Writing them into `<f>` verbatim yields
//! a workbook Excel offers to repair and no engine can recalc, so references
//! are rewritten into A1 semantics here. Any construct this translator cannot
//! prove correct gives up rather than guess: the caller then keeps the cell's
//! cached result, which is what ODF stores beside the formula anyway.

/// Translate one formula. `None` means "no safe OOXML counterpart" and the
/// caller must drop the formula and keep the value. A formula that is not in
/// the ODF dialect (`.xls`/`.xlsx`, whose references calamine already returns
/// in A1 form) is returned untouched.
pub fn translate_formula(text: &str) -> Option<String> {
    let Some(body) = odf_body(text) else {
        if text.contains('[') {
            // A bracket outside the ODF dialect is an external-workbook
            // reference this converter has no way to renumber.
            return None;
        }
        return Some(text.trim().to_string());
    };
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::with_capacity(body.len());
    let mut index = 0usize;
    let mut depth = 0usize;
    while index < chars.len() {
        let copied = match chars[index] {
            '[' => {
                let (reference, next) = read_reference(&chars, index)?;
                out.push_str(&reference);
                index = next;
                continue;
            }
            '"' => {
                let (literal, next) = read_string(&chars, index)?;
                out.push_str(&literal);
                index = next;
                continue;
            }
            '(' => {
                depth += 1;
                "("
            }
            ')' => {
                depth = depth.checked_sub(1)?;
                ")"
            }
            ';' => ",",
            // OOXML has no token for any of these: array constants and named
            // ranges ($ / #), the intersection and union operators, the
            // alternative string quoting, the indirect-reference operator.
            '{' | '}' | '#' | '!' | '$' | '\'' | '@' => return None,
            '.' => {
                // Only a decimal point survives the move. A '.' outside a
                // reference is either a sheet separator that escaped the
                // bracket reader or part of a vendor function name
                // (COM.MICROSOFT.*), and neither has an A1 spelling.
                let between_digits = index
                    .checked_sub(1)
                    .is_some_and(|previous| chars[previous].is_ascii_digit())
                    && chars
                        .get(index + 1)
                        .is_some_and(|next| next.is_ascii_digit());
                if !between_digits {
                    return None;
                }
                "."
            }
            other => {
                out.push(other);
                index += 1;
                continue;
            }
        };
        out.push_str(copied);
        index += 1;
    }
    if depth != 0 || out.trim().is_empty() {
        return None;
    }
    Some(out)
}

/// The expression after the `of:` namespace marker, or `None` when the formula
/// is not in the ODF dialect.
fn odf_body(text: &str) -> Option<&str> {
    let after = text.trim().strip_prefix("of:")?;
    Some(after.strip_prefix('=').unwrap_or(after).trim())
}

/// A `[...]` reference group becomes `A1`, `Sheet!A1` or `A1:B2`.
fn read_reference(chars: &[char], open: usize) -> Option<(String, usize)> {
    let rest = &chars[open + 1..];
    let offset = rest.iter().position(|c| *c == ']')?;
    let close = open + 1 + offset;
    // A nested '[' is not something the ODF grammar puts inside a reference.
    if chars[open + 1..close].contains(&'[') {
        return None;
    }
    let inner: String = chars[open + 1..close].iter().collect();
    let parts: Vec<&str> = inner.split(':').collect();
    if parts.len() > 2 {
        return None;
    }
    let mut sheets: Vec<Option<String>> = Vec::with_capacity(parts.len());
    let mut references: Vec<String> = Vec::with_capacity(parts.len());
    for part in &parts {
        let (sheet, reference) = split_reference(part)?;
        sheets.push(sheet);
        references.push(reference);
    }
    // ODF qualifies every endpoint of a range, so all endpoints must agree: a
    // mixed pair would silently move a reference to another sheet, and a
    // two-sheet range has no two-sheet spelling in A1 semantics.
    let sheet = sheets[0].clone();
    if sheets.iter().any(|name| *name != sheet) {
        return None;
    }
    let mut out = sheet.map_or_else(String::new, |name| format!("{name}!"));
    out.push_str(&references.join(":"));
    Some((out, close + 1))
}

/// Split `[sheet.]cell` at the last dot: ODF sheet names may contain one, and a
/// name is only usable when what follows it still parses as a reference.
fn split_reference(part: &str) -> Option<(Option<String>, String)> {
    let dot = part.rfind('.')?;
    let sheet = sheet_name(&part[..dot])?;
    Some((sheet, cell_reference(&part[dot + 1..])?))
}

/// The sheet token of a reference: an optional absolute-sheet `$`, then either
/// a bare name or a `'`-quoted one whose embedded quotes are doubled.
///
/// The `$` and the quotes are ODF addressing syntax, not part of the name:
/// LibreOffice writes every other-sheet reference in that absolute form
/// (`[$Other.A1]`, `[$'My Sheet'.A1]`), and passing either through unstripped
/// yields an A1 reference to a sheet that does not exist. Quoting that survives
/// is re-decided by `quoted_sheet_name`, since the name may still need it.
fn sheet_name(token: &str) -> Option<Option<String>> {
    let bare = token.strip_prefix('$').unwrap_or(token);
    if bare.is_empty() {
        // `[.A1]` — the self sheet, which A1 leaves unqualified.
        return Some(None);
    }
    let name = match bare.strip_prefix('\'') {
        None => bare.to_string(),
        Some(quoted) => {
            let body = quoted.strip_suffix('\'')?;
            // Every quote inside a quoted name belongs to a `''` pair; a lone
            // one means this was never a quoted name.
            let mut name = String::with_capacity(body.len());
            let mut rest = body;
            while let Some(at) = rest.find('\'') {
                name.push_str(&rest[..at]);
                if !rest[at..].starts_with("''") {
                    return None;
                }
                name.push('\'');
                rest = &rest[at + 2..];
            }
            name.push_str(rest);
            name
        }
    };
    Some(Some(quoted_sheet_name(&name)?))
}

/// `A1`, `$A$1`, `A$1`, a whole column (`A`) or a whole row (`1`).
fn cell_reference(text: &str) -> Option<String> {
    let chars: Vec<char> = text.chars().collect();
    let mut index = 0usize;
    let leading_marker = eat_dollar(&chars, &mut index);
    let letters_start = index;
    let mut column = 0u32;
    while index < chars.len() && chars[index].is_ascii_alphabetic() {
        column = column * 26 + (chars[index].to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        index += 1;
    }
    let letters = index - letters_start;
    if letters > 3 || column > 16_384 {
        return None;
    }
    let row_marker = eat_dollar(&chars, &mut index);
    let digits_start = index;
    while index < chars.len() && chars[index].is_ascii_digit() {
        index += 1;
    }
    let digits = index - digits_start;
    if index != chars.len() || digits > 7 || (letters == 0 && digits == 0) {
        return None;
    }
    if digits > 0 && row_number(&chars[digits_start..index]) > 1_048_576 {
        return None;
    }
    let mut out = String::with_capacity(chars.len());
    if letters == 0 {
        // A whole row: the leading `$` can only have marked the row.
        if leading_marker || row_marker {
            out.push('$');
        }
    } else {
        if leading_marker {
            out.push('$');
        }
        for letter in &chars[letters_start..letters_start + letters] {
            out.push(letter.to_ascii_uppercase());
        }
        if row_marker {
            out.push('$');
        }
    }
    out.extend(chars[digits_start..index].iter().copied());
    Some(out)
}

fn row_number(digits: &[char]) -> u64 {
    digits
        .iter()
        .fold(0u64, |value, digit| value * 10 + u64::from(*digit as u8 - b'0'))
}

fn eat_dollar(chars: &[char], index: &mut usize) -> bool {
    if chars.get(*index) == Some(&'$') {
        *index += 1;
        return true;
    }
    false
}

/// Quote a sheet name only when it needs it, doubling embedded quotes.
fn quoted_sheet_name(name: &str) -> Option<String> {
    if name.is_empty() || name.chars().any(char::is_control) {
        return None;
    }
    let plain = name
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || c == '_')
        && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
    if plain {
        return Some(name.to_string());
    }
    Some(format!("'{}'", name.replace('\'', "''")))
}

/// Copy a `"..."` literal through, including its doubled-quote escapes.
fn read_string(chars: &[char], open: usize) -> Option<(String, usize)> {
    let mut index = open + 1;
    let mut out = String::from("\"");
    loop {
        let c = *chars.get(index)?;
        if c == '"' {
            // `""` inside a literal is an escaped quote, not the terminator.
            if chars.get(index + 1) == Some(&'"') {
                out.push_str("\"\"");
                index += 2;
                continue;
            }
            out.push('"');
            return Some((out, index + 1));
        }
        out.push(c);
        index += 1;
    }
}

#[cfg(test)]
mod tests {
    use super::translate_formula;

    fn translated(formula: &str) -> Option<String> {
        translate_formula(&format!("of:={formula}"))
    }

    #[test]
    fn rewrites_self_sheet_references_and_keeps_operators() {
        assert_eq!(translated("[.A1]+[.B1]*2").as_deref(), Some("A1+B1*2"));
        assert_eq!(translated("[.A1]*3.14").as_deref(), Some("A1*3.14"));
        assert_eq!(translated("SUM([.A1:.B1])").as_deref(), Some("SUM(A1:B1)"));
        assert_eq!(
            translated("[.$A$1]+[.A$2]+[.$A3]").as_deref(),
            Some("$A$1+A$2+$A3")
        );
    }

    #[test]
    fn qualifies_sheet_names_and_quotes_the_ones_needing_it() {
        assert_eq!(translated("[Other.B1]*2").as_deref(), Some("Other!B1*2"));
        assert_eq!(translated("[My Sheet.A1]").as_deref(), Some("'My Sheet'!A1"));
        assert_eq!(
            translated("[Q1.Revenue.A1]").as_deref(),
            Some("'Q1.Revenue'!A1")
        );
        assert_eq!(
            translated("SUM([Data.A1:Data.B2])").as_deref(),
            Some("SUM(Data!A1:B2)")
        );
    }

    // LibreOffice writes every other-sheet reference in the absolute-sheet
    // form: the `$` and the apostrophes are addressing syntax the A1 dialect
    // spells differently, and keeping them produces references to sheets that
    // do not exist.
    #[test]
    fn unwraps_absolute_and_quoted_sheet_names() {
        assert_eq!(translated("[$Other.A1]*2").as_deref(), Some("Other!A1*2"));
        assert_eq!(
            translated("[$'My Sheet'.A1]+1").as_deref(),
            Some("'My Sheet'!A1+1")
        );
        assert_eq!(
            translated("SUM([$Data.A1:$Data.B2])").as_deref(),
            Some("SUM(Data!A1:B2)")
        );
        // the `$` marks the sheet only; cell-level absolute markers survive
        assert_eq!(
            translated("SUM([$Data.$A$1:$Data.$B$2])").as_deref(),
            Some("SUM(Data!$A$1:$B$2)")
        );
        // a quoted name that escapes an embedded quote
        assert_eq!(
            translated("[$'Bob''s data'.A1]").as_deref(),
            Some("'Bob''s data'!A1")
        );
        // an unterminated or unescaped quote is not a name this can prove
        assert_eq!(translated("[$'My Sheet.A1]"), None);
        assert_eq!(translated("[$'My'Sheet'.A1]"), None);
        // a `$` that is not a sheet marker still reaches the cell reader
        assert_eq!(translated("[$A$3]"), None);
    }

    #[test]
    fn maps_whole_column_and_whole_row_ranges() {
        assert_eq!(translated("SUM([.A:.C])").as_deref(), Some("SUM(A:C)"));
        assert_eq!(translated("SUM([.1:.5])").as_deref(), Some("SUM(1:5)"));
    }

    #[test]
    fn separates_arguments_with_commas_but_not_inside_literals() {
        assert_eq!(
            translated(r#"CONCATENATE([.A2];"!;")"#).as_deref(),
            Some(r#"CONCATENATE(A2,"!;")"#)
        );
        // A doubled quote stays an escaped quote rather than ending the literal.
        assert_eq!(
            translated(r#"CONCATENATE("a""b";[.A1])"#).as_deref(),
            Some(r#"CONCATENATE("a""b",A1)"#)
        );
    }

    #[test]
    fn refuses_references_outside_the_ooxml_grid() {
        assert_eq!(translated("[.A1048577]"), None);
        assert_eq!(translated("[.ZZZ1]"), None);
        assert_eq!(translated("[.A1x]"), None);
        assert_eq!(translated("[]"), None);
        // ODF spells every reference `[.A1]`; a bracket without the self-sheet
        // dot is a dialect this translator cannot identify, so it gives up.
        assert_eq!(translated("[$A3]"), None);
        assert_eq!(translated("[A1]"), None);
    }

    #[test]
    fn passes_a_non_odf_formula_through() {
        assert_eq!(translate_formula("B1*2").as_deref(), Some("B1*2"));
        // An unprefixed bracket is an external-workbook reference, not a miss.
        assert_eq!(translate_formula("[1]Sheet!A1"), None);
    }

    #[test]
    fn gives_up_on_constructs_with_no_ooxml_spelling() {
        // named ranges, array constants, operators OOXML lacks, vendor names
        assert_eq!(translated("$Sales"), None);
        assert_eq!(translated("{1;2;3}"), None);
        assert_eq!(translated("[.A1]![.B1]"), None);
        assert_eq!(translated("COM.MICROSOFT.SUM([.A1])"), None);
        assert_eq!(translated("[#ALL]"), None);
        assert_eq!(translated("[.A1"), None);
        assert_eq!(translated("SUM([.A1]"), None);
        assert_eq!(translated("SUM([.A1:Other.B2])"), None);
    }
}

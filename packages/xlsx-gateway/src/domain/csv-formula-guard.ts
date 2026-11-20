/**
 * CSV formula-injection guard, shared by the two CSV writers (the CLI's
 * sheet-to-csv and the sheets renderer's csv-export).
 *
 * Spreadsheet apps evaluate some cell-leading shapes as a formula or DDE
 * instruction when the CSV is opened — the standard CSV injection chain. The
 * writers quote for RFC 4180 only; this guard adds the prefix neutralization:
 * an affected field is written with a leading `'`, which spreadsheets treat as
 * a text marker and hide.
 *
 * Only the shapes that actually evaluate neutralize, and the writers apply the
 * guard as an opt-in: a leading `=`, a leading `@` or `+`/`-` followed by a
 * function call anywhere behind it (`+SUM(A1:A9)` but also
 * `+1+WEBSERVICE(...)`, `-(HYPERLINK(...))`, `+A1&WEBSERVICE(...)`), or the DDE
 * form (`program|topic!cell`, the payload behind `+cmd|' /C calc'!A0`).
 * Ordinary data never gains the apostrophe — `-`, `-$1,234.50`, `- item`,
 * `@handle`, `+86 138` keep their bytes, because Excel does not hide the
 * apostrophe on CSV open and pandas or a database import reads it as data.
 */
export function startsLikeSpreadsheetFormula(field: string): boolean {
  // spaces only: a tab prefix is itself one of the injection prefixes
  const trimmed = field.replace(/^ +/, '')
  const first = trimmed[0]
  if (first === '=' || first === '\t') return true
  if (first === '@' || first === '+' || first === '-') {
    const rest = trimmed.slice(1)
    // a function call anywhere behind the prefix evaluates — it need not sit
    // directly after the sign: "+1+WEBSERVICE(...)", "-1+HYPERLINK(...)",
    // "-(HYPERLINK(...))" and "+A1&WEBSERVICE(...)" all execute. The name(
    // shape is written without a space, which keeps prose ("see note (below)")
    // out of the net.
    if (/[A-Za-z_][A-Za-z0-9_.]*\(/.test(rest)) return true
    // the DDE shape: program|topic!cell
    if (rest.includes('|') && rest.includes('!')) return true
    return false
  }
  return false
}

/** The field with a text-marker apostrophe prepended when it reads as a formula. */
export function neutralizeCsvFormulaField(field: string): string {
  return startsLikeSpreadsheetFormula(field) ? `'${field}` : field
}

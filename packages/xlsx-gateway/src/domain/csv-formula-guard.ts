/**
 * CSV formula-injection guard, shared by the two CSV writers (the CLI's
 * sheet-to-csv and the sheets renderer's csv-export).
 *
 * Spreadsheet apps evaluate a cell's leading `=`, `+`, `-` or `@` (and tab /
 * carriage-return prefixes) as a formula or DDE instruction when the CSV is
 * opened, so a hostile workbook's cell text that survives export verbatim
 * executes in whatever opens the file next — the standard CSV injection
 * chain. The writers quote for RFC 4180 only; this guard adds the prefix
 * neutralization: an affected field is written with a leading `'`, which
 * spreadsheets treat as a text marker and hide.
 *
 * A leading `+`/`-` directly followed by a digit or point is a plain number
 * and stays untouched, so legitimate numeric exports keep their shape.
 */
export function startsLikeSpreadsheetFormula(field: string): boolean {
  // spaces only: a tab prefix is itself one of the injection prefixes
  const trimmed = field.replace(/^ +/, '')
  const first = trimmed[0]
  if (first === '=' || first === '@' || first === '\t') return true
  if (first === '+' || first === '-') {
    const next = trimmed[1]
    return !(next !== undefined && /[\d.]/.test(next))
  }
  return false
}

/** The field with a text-marker apostrophe prepended when it reads as a formula. */
export function neutralizeCsvFormulaField(field: string): string {
  return startsLikeSpreadsheetFormula(field) ? `'${field}` : field
}

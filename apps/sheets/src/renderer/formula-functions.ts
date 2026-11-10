/// Function-call extraction for file formulas. Pure text analysis: the
/// keep-cached-value decision needs to know WHICH functions a formula calls
/// so it can ask the running engine whether it implements every one of them.

/// Excel stores post-2007 functions behind `_xlfn.` (worksheet-scope
/// dynamic-array functions behind `_xlfn._xlws.`); the registry knows the
/// plain name. The sidecar already strips these on read, so this is only a
/// guard for formulas that reach the renderer through another path.
const STORAGE_MARKERS = /^(?:_XLFN\.|_XLWS\.)+/

function isIdentifierStart(character: string): boolean {
  return /[\p{L}_]/u.test(character)
}

function isIdentifierPart(character: string): boolean {
  return /[\p{L}\p{N}_.]/u.test(character)
}

function isBlank(character: string): boolean {
  return character === ' ' || character === '\t' || character === '\r' || character === '\n'
}

/// Skips a quoted run (`"..."` string literal or `'...'` sheet name) whose
/// opening quote sits at `index`; a doubled quote is an escaped quote.
/// Returns the index just past the closing quote (or the formula's end).
function skipQuoted(formula: string, index: number, quote: string): number {
  let cursor = index + 1
  while (cursor < formula.length) {
    if (formula[cursor] !== quote) {
      cursor += 1
      continue
    }
    if (formula[cursor + 1] === quote) {
      cursor += 2
      continue
    }
    return cursor + 1
  }
  return cursor
}

/// Skips a bracketed run (structured reference `Table[[#Headers],[Col]]` or
/// external-workbook index `[1]`) whose opening bracket sits at `index`.
/// Column names inside may contain "(" and would otherwise read as calls.
function skipBracketed(formula: string, index: number): number {
  let depth = 0
  let cursor = index
  while (cursor < formula.length) {
    const character = formula[cursor]
    if (character === '[') depth += 1
    else if (character === ']') {
      depth -= 1
      if (depth === 0) return cursor + 1
    }
    cursor += 1
  }
  return cursor
}

/// Canonical registry key for a called identifier: uppercase, storage
/// markers stripped (`_xlfn._xlws.FILTER` → `FILTER`).
export function canonicalFunctionName(identifier: string): string {
  return identifier.toUpperCase().replace(STORAGE_MARKERS, '')
}

/// Every function the formula calls, as canonical names, in first-seen
/// order. An identifier counts as a call only when the next non-blank
/// character is "(" — a reference or a defined name is never followed by
/// one, a call always is — so `Sheet1!A1` and `LOG10(A1)` sort themselves
/// out. Identifiers inside string literals, quoted sheet names, and
/// bracketed (structured / external) references are skipped.
export function extractFunctionNames(formula: string): string[] {
  const names = new Set<string>()
  const length = formula.length
  let index = 0
  while (index < length) {
    const character = formula[index] ?? ''
    if (character === '"' || character === "'") {
      index = skipQuoted(formula, index, character)
      continue
    }
    if (character === '[') {
      index = skipBracketed(formula, index)
      continue
    }
    if (!isIdentifierStart(character)) {
      index += 1
      continue
    }
    let end = index + 1
    while (end < length && isIdentifierPart(formula[end] ?? '')) end += 1
    // A letter glued to a number literal (1E5, 2.5E-3) is an exponent, not
    // an identifier; the digits before it were skipped one by one above.
    const previous = index === 0 ? '' : (formula[index - 1] ?? '')
    if (!/[\p{N}.]/u.test(previous)) {
      let after = end
      while (after < length && isBlank(formula[after] ?? '')) after += 1
      if (formula[after] === '(') names.add(canonicalFunctionName(formula.slice(index, end)))
    }
    index = end
  }
  return [...names]
}

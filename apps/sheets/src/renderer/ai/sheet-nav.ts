/**
 * In-answer citations: the model links cells as [B12](sheetnav://B12) or
 * [Data!B2:D9](sheetnav://Data!B2:D9), and clicking one jumps the grid there
 * through the same Go To path the Name Box uses.
 *
 * A citation is a link the user chooses to follow, so pointing at a cell no
 * longer costs a select_range call that takes the selection away from them.
 */

export const SHEET_NAV_SCHEME = 'sheetnav://'

/**
 * sheetnav://Data!B2:D9 -> "Data!B2:D9"; null for anything else. The markdown
 * link syntax stops at the first space, so sheet names containing one arrive
 * percent-escaped — and once decoded they need Excel's quoting before the Go To
 * parser will take them.
 */
export function parseSheetNavHref(href: string): string | null {
  if (!href.startsWith(SHEET_NAV_SCHEME)) return null
  const raw = href.slice(SHEET_NAV_SCHEME.length).trim()
  if (raw === '') return null
  let ref: string
  try {
    ref = decodeURIComponent(raw)
  } catch {
    // a stray '%' is not worth dropping the citation over
    ref = raw
  }
  return quoteSheetPrefix(ref.trim()) || null
}

/** `My Summary!B2` -> `'My Summary'!B2`; already-quoted and bare refs pass through. */
function quoteSheetPrefix(ref: string): string {
  const bang = ref.lastIndexOf('!')
  if (bang <= 0) return ref
  const sheet = ref.slice(0, bang)
  const body = ref.slice(bang + 1)
  if (sheet.startsWith("'") || !/[\s'()]/.test(sheet)) return ref
  return `'${sheet.replace(/'/g, "''")}'!${body}`
}

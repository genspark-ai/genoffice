/**
 * Name-addressed worksheet references for the MCP surface.
 *
 * The app's DSL addresses worksheets by id. A sheet id is a session artifact —
 * `read_sheet` mints `sheet-3` today and `jg_nIT-Asw-…` for the same sheet
 * tomorrow — so an agent that cached one from an earlier call writes into
 * nothing ("Unknown sheet") or, worse, into whatever now carries that id. The
 * bundled CLI has always addressed sheets by NAME (`"sheet": "Sheet1"`), which
 * is stable across sessions and is what the user sees on the tab.
 *
 * The MCP tools therefore accept the CLI's name fields (`sheet`,
 * `sourceSheet`, `targetSheet`) alongside the id fields, and this module
 * translates names to ids in the renderer — the only process that knows the
 * open workbook's sheets. A caller that still sends ids keeps working
 * unchanged.
 */

/** one worksheet of the open workbook, as the read context reports it */
export interface SheetRef {
  readonly id: string
  readonly name: string
}

/** the reference fields a name may be given in, and the id each maps to */
const NAME_TO_ID_FIELDS: ReadonlyArray<readonly [name: string, id: string]> = [
  ['sheet', 'sheetId'],
  ['sourceSheet', 'sourceSheetId'],
  ['targetSheet', 'targetSheetId'],
]

const NAME_FIELDS = NAME_TO_ID_FIELDS.map(([name]) => name)
const ID_FIELDS = NAME_TO_ID_FIELDS.map(([, id]) => id)

export type SheetRefRewrite =
  { readonly ok: true; readonly ops: unknown[] } | { readonly ok: false; readonly error: string }

/**
 * Resolve every `sheet` / `sourceSheet` / `targetSheet` in a batch to the id
 * field the DSL expects, recursively (chart series address sheets one level
 * down). Ops that carry no name are returned as they came.
 *
 * An unknown name fails the whole batch before anything applies: dropping the
 * reference would let the op fall back to the active sheet and write into the
 * wrong place, which is the failure this addressing exists to prevent.
 */
export function normalizeSheetRefs(
  ops: readonly unknown[],
  sheets: readonly SheetRef[],
): SheetRefRewrite {
  if (!ops.some((op) => namesASheet(op))) return { ok: true, ops: [...ops] }
  const byName = new Map<string, string>()
  const byId = new Set<string>()
  for (const sheet of sheets) {
    // First tab wins on duplicate names (case/space-insensitive): tab order is
    // the user's visible order, so the leftmost match is the least surprising
    // target, and keeping the first keeps resolution deterministic.
    const key = sheet.name.trim().toLowerCase()
    if (!byName.has(key)) byName.set(key, sheet.id)
    byId.add(sheet.id)
  }
  const known = sheets.length === 0 ? 'none' : sheets.map((sheet) => sheet.name).join(', ')
  const resolved: unknown[] = []
  for (let index = 0; index < ops.length; index += 1) {
    const rewritten = rewriteOne(ops[index], byName, byId, known)
    if (!rewritten.ok) return { ok: false, error: `op #${index} ${rewritten.error}` }
    resolved.push(rewritten.value)
  }
  return { ok: true, ops: resolved }
}

/** true when any object below this one carries one of the name fields */
function namesASheet(value: unknown): boolean {
  if (Array.isArray(value)) return value.some(namesASheet)
  if (typeof value !== 'object' || value === null) return false
  const record = value as Record<string, unknown>
  // Empty refs still count here so rewriteOne can strip them; they never fail.
  if (NAME_FIELDS.some((field) => typeof record[field] === 'string')) return true
  return Object.values(record).some(namesASheet)
}

type RewriteStep =
  { readonly ok: true; readonly value: unknown } | { readonly ok: false; readonly error: string }

/** the translated value, or the reason the batch cannot be addressed */
function rewriteOne(
  value: unknown,
  byName: ReadonlyMap<string, string>,
  byId: ReadonlySet<string>,
  known: string,
): RewriteStep {
  if (Array.isArray(value)) {
    const items: unknown[] = []
    for (const item of value) {
      const step = rewriteOne(item, byName, byId, known)
      if (!step.ok) return step
      items.push(step.value)
    }
    return { ok: true, value: items }
  }
  if (typeof value !== 'object' || value === null) return { ok: true, value }
  const record = { ...(value as Record<string, unknown>) }
  for (const [nameField, idField] of NAME_TO_ID_FIELDS) {
    const given = record[nameField]
    if (typeof given !== 'string') continue
    // An empty or whitespace-only ref is ignored like an absent one: strip it
    // so an agent passing "" does not fail the whole batch.
    if (given.trim() === '') {
      delete record[nameField]
      continue
    }
    delete record[nameField]
    // An id passed in the name slot (an agent reusing an earlier read) is
    // accepted as-is rather than reported as an unknown worksheet name.
    if (byId.has(given)) {
      record[idField] ??= given
      continue
    }
    const id = byName.get(given.trim().toLowerCase())
    if (id === undefined) {
      return {
        ok: false,
        error: `addresses worksheet "${given}", which is not in this workbook (sheets: ${known}); call read_sheet for the current sheets`,
      }
    }
    // An explicit id wins: a batch may legitimately carry both (an agent
    // correcting itself), and the id is the more precise reference.
    record[idField] ??= id
  }
  for (const [key, nested] of Object.entries(record)) {
    if (key === 'op') continue
    const step = rewriteOne(nested, byName, byId, known)
    if (!step.ok) return step
    record[key] = step.value
  }
  return { ok: true, value: record }
}

/**
 * The worksheet a batch primarily addresses: the first op naming one. Focus
 * follows it, so an edit to a sheet the view is not showing cannot land
 * silently off-screen. Nested sheet ids (chart series / seriesData one level
 * down, or deeper) count: an edit_chart batch names no top-level sheet but
 * still edits one.
 */
export function primarySheetId(ops: readonly unknown[]): string | undefined {
  for (const op of ops) {
    const found = findSheetId(op)
    if (found !== undefined) return found
  }
  return undefined
}

/** depth-first id search within one op: top-level ids first, then nested */
function findSheetId(value: unknown): string | undefined {
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findSheetId(item)
      if (found !== undefined) return found
    }
    return undefined
  }
  if (typeof value !== 'object' || value === null) return undefined
  const record = value as Record<string, unknown>
  for (const idField of ID_FIELDS) {
    const id = record[idField]
    if (typeof id === 'string' && id !== '') return id
  }
  for (const [key, nested] of Object.entries(record)) {
    if (key === 'op') continue
    const found = findSheetId(nested)
    if (found !== undefined) return found
  }
  return undefined
}

/**
 * The cell the batch's first addressing op touches — a single-cell `address`,
 * or the top-left of a `range` / `target` — used to scroll focus onto the
 * edit rather than merely opening the sheet. Ops addressed only by
 * sourceSheet / targetSheet still count. Whole-column ranges ("A:C") focus
 * the top of the first column ("A1"); named ranges and whole rows have no
 * single cell to focus, so they yield undefined (sheet-only focus).
 */
export function primaryCellOf(ops: readonly unknown[]): string | undefined {
  for (const op of ops) {
    if (typeof op !== 'object' || op === null) continue
    const record = op as Record<string, unknown>
    if (!addressesAnySheet(record)) continue
    for (const key of ['address', 'range', 'target', 'source']) {
      const value = record[key]
      if (typeof value !== 'string' || value.trim() === '') continue
      const start = primaryCellStart(value)
      if (start !== undefined) return start
    }
  }
  return undefined
}

/** true when the op carries any sheet ref (sheet / source / target, name or id) */
function addressesAnySheet(record: Record<string, unknown>): boolean {
  for (const field of [
    'sheet',
    'sheetId',
    'sourceSheet',
    'sourceSheetId',
    'targetSheet',
    'targetSheetId',
  ]) {
    const value = record[field]
    if (typeof value === 'string' && value.trim() !== '') return true
  }
  return false
}

/** top-left cell of a range string, or undefined when there is no single cell */
function primaryCellStart(value: string): string | undefined {
  const trimmed = value.trim()
  if (trimmed === '') return undefined
  const parts = trimmed.split(':')
  const firstRaw = parts[0]?.trim() ?? ''
  const first = firstRaw.includes('!') ? (firstRaw.split('!').pop()?.trim() ?? '') : firstRaw
  if (first === '') return undefined
  // A plain cell ("B2", "$A$1").
  if (/^\$?[A-Za-z]+\$?\d+$/.test(first)) return first
  // A whole-column range ("A:C", "$B:$D"): focus the top of the first column.
  // Both sides must be short column labels; a bare defined name without a
  // colon has no single cell, so it yields undefined.
  if (parts.length === 2) {
    const secondRaw = parts[1]?.trim() ?? ''
    const second = secondRaw.includes('!') ? (secondRaw.split('!').pop()?.trim() ?? '') : secondRaw
    const firstCol = /^\$?([A-Za-z]{1,3})$/.exec(first)?.[1]
    const secondCol = /^\$?([A-Za-z]{1,3})$/.exec(second)?.[1]
    if (firstCol && secondCol) return `${firstCol}1`
  }
  return undefined
}

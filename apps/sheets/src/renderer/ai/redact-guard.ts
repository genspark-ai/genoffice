/**
 * Stop a model write from landing on a cell the reader withheld.
 *
 * ## Damage, not disclosure
 *
 * The read path is what keeps the values from the model (see `./redact`). What
 * is left is the model **writing over** them, and that failure is silent: the
 * file still opens, it just no longer holds what the reader kept. So this is a
 * damage guard — the one other apps already keep in the same place
 * (`apps/slides/.../redact-guard.ts`, `apps/html/.../redact-guard.ts`).
 *
 * ## Why it is cheap here
 *
 * Every write op names the cells it touches in A1 notation, and a mark is a
 * rectangle in the same zero-based coordinates, so "does this op reach a
 * withheld cell" is a rectangle intersection — no cell reads, no before/after
 * comparison, and no cost at all for a workbook that withholds nothing (the
 * shared empty index answers `isEmpty` first).
 *
 * ## What is deliberately still allowed
 *
 * Formatting, colours, notes, filters, validation, conditional formats and
 * moving/resizing a chart do not change a cell's value, and refusing them would
 * make the feature unusable for ordinary work. Inserting a *blank* row or
 * column is allowed too; what is refused is anything that would move or remove
 * a withheld cell, because the mark does not travel with it — the value would
 * reappear in an unmarked cell on the next open.
 */
import {
  columnIndex,
  formatAddress,
  parseAddress,
  parseRange,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { placeholderSource, type CellMark, type RedactionIndex } from './redact'

export interface RedactionRefusal {
  /** the whole message the model sees, naming the cell and the label */
  readonly reason: string
  readonly label: string
  /** A1 address of the first withheld cell the op reached */
  readonly where: string
}

/** op name → the field holding the cells it writes, and how to read them. */
const WRITES: Record<string, (op: Record<string, unknown>) => RangeBounds | null> = {
  set_cell: (op) => cellOf(op.address),
  set_formula: (op) => cellOf(op.address),
  clear_cell: (op) => cellOf(op.address),
  // set_range writes a grid from `start` (or an explicit `range`, whose size
  // must match the values). The grid's own extent is the truth when present.
  set_range: (op) => gridOf(op),
  clear_range: (op) => rangeOf(op.range),
  fill_range: (op) => rangeOf(op.target),
  copy_range: (op) => rangeOf(op.target),
  // a move clears its source and writes its target; both are checked
  move_range: (op) => rangeOf(op.target),
  convert_to_values: (op) => rangeOf(op.range),
  sort_range: (op) => rangeOf(op.range),
  find_replace: (op) => rangeOf(op.range),
  // Merging keeps only the top-left value, so every other cell in the band is
  // destroyed. A one-cell merge changes nothing and stays allowed.
  merge_cells: (op) => {
    const bounds = rangeOf(op.range)
    if (!bounds) return null
    return bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn
      ? null
      : bounds
  },
  add_table: (op) => rangeOf(op.range),
  // Deleting rows/columns takes the withheld cells with them; the marks stay
  // where they are and would then describe whatever shifted into place.
  delete_rows: (op) => rowBand(op.row, op.count),
  delete_cols: (op) => colBand(op.column, op.count),
  // Dropping the whole sheet drops every value on it.
  delete_sheet: () => ENTIRE_SHEET,
  // A duplicate copies the withheld values into a new sheet that carries no
  // marks, which publishes them.
  duplicate_sheet: () => ENTIRE_SHEET,
}

/**
 * Where an op takes its values from. A withheld value copied, moved or
 * referenced by a formula lands in an unmarked cell (or in the readback of
 * the formula's result), which publishes it — so the source is refused too.
 */
interface ReadSite {
  /** the sheet the read happens on; undefined = the op's own sheet */
  sheetId?: string | undefined
  /** a sheet named in a formula, resolved by the caller */
  sheetName?: string | undefined
  bounds: RangeBounds
}

const READS: Record<string, (op: Record<string, unknown>) => ReadSite[]> = {
  copy_range: (op) => sourceOf(op),
  move_range: (op) => sourceOf(op),
  fill_range: (op) => sourceOf(op),
  set_formula: (op) => formulaRefs(op.formula),
  set_cell: (op) => formulaRefs(op.value),
  set_range: (op) => {
    const values = Array.isArray(op.values) ? op.values : []
    return values.flatMap((row) => (Array.isArray(row) ? row.flatMap(formulaRefs) : []))
  },
}

function sourceOf(op: Record<string, unknown>): ReadSite[] {
  const bounds = rangeOf(op.source)
  if (!bounds) return []
  return [{ sheetId: typeof op.sourceSheetId === 'string' ? op.sourceSheetId : undefined, bounds }]
}

/** A1 references of a formula, string literals set aside; a sheet prefix stays a name */
const FORMULA_REF =
  /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![\w(])/g

function formulaRefs(value: unknown): ReadSite[] {
  if (typeof value !== 'string' || !value.startsWith('=')) return []
  const bare = value.replace(/"(?:[^"]|"")*"/g, '""')
  const out: ReadSite[] = []
  for (const m of bare.matchAll(FORMULA_REF)) {
    const ref = m[3]!.replace(/\$/g, '')
    const bounds = refBounds(ref)
    if (!bounds) continue
    const sheetName = m[1]?.replace(/''/g, "'") ?? m[2]
    out.push({ sheetName, bounds })
  }
  return out
}

function refBounds(ref: string): RangeBounds | null {
  const cols = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(ref)
  if (cols) {
    const a = columnIndex(cols[1]!)
    const b = columnIndex(cols[2]!)
    return { startRow: 0, endRow: MAX_ROW, startColumn: Math.min(a, b), endColumn: Math.max(a, b) }
  }
  const rows = /^(\d+):(\d+)$/.exec(ref)
  if (rows) {
    const a = Number(rows[1]) - 1
    const b = Number(rows[2]) - 1
    return {
      startRow: Math.min(a, b),
      endRow: Math.max(a, b),
      startColumn: 0,
      endColumn: MAX_COLUMN,
    }
  }
  return ref.includes(':') ? rangeOf(ref) : cellOf(ref)
}

/**
 * Ops that shift every cell at or below a point rather than writing one.
 * Refused only when the shift would carry a withheld cell away from its mark.
 */
const SHIFTS: Record<string, (op: Record<string, unknown>) => (mark: CellMark) => boolean> = {
  insert_rows: (op) => {
    const from = rowStart(op.row)
    return (mark) => mark.startRow >= from
  },
  insert_cols: (op) => {
    const from = colStart(op.column)
    return (mark) => mark.startColumn >= from
  },
}

/** Excel's grid limits, used where an op covers a whole row or column band. */
const MAX_ROW = 1_048_575
const MAX_COLUMN = 16_383

const ENTIRE_SHEET: RangeBounds = {
  startRow: 0,
  endRow: Number.MAX_SAFE_INTEGER,
  startColumn: 0,
  endColumn: Number.MAX_SAFE_INTEGER,
}

function cellOf(address: unknown): RangeBounds | null {
  if (typeof address !== 'string') return null
  try {
    const { row, column } = parseAddress(address)
    return { startRow: row, endRow: row, startColumn: column, endColumn: column }
  } catch (error) {
    // a malformed model-supplied address is a normal refusal, but a silent
    // swallow would also hide a genuine parse bug
    console.debug('[redact-guard] unparsable address', address, error)
    return null
  }
}

function rangeOf(range: unknown): RangeBounds | null {
  if (typeof range !== 'string') return null
  try {
    return parseRange(range)
  } catch (error) {
    console.debug('[redact-guard] unparsable range', range, error)
    return null
  }
}

/**
 * The cells a `set_range` writes.
 *
 * The anchor cell alone is not the answer: the op paints a whole grid of values
 * down and to the right of it, and checking only the anchor would wave a write
 * straight through a mark one column further on. `range` is accepted as an
 * alternative anchor and the schema requires it to match the grid's size, so
 * either gives the same rectangle.
 */
function gridOf(op: Record<string, unknown>): RangeBounds | null {
  const explicit = rangeOf(op.range)
  const start = cellOf(op.start)
  if (!start) return explicit
  if (explicit) return explicit
  const values = op.values
  if (!Array.isArray(values) || values.length === 0) return start
  const width = values.reduce(
    (widest, row) => (Array.isArray(row) && row.length > widest ? row.length : widest),
    0,
  )
  return {
    startRow: start.startRow,
    endRow: start.startRow + values.length - 1,
    startColumn: start.startColumn,
    endColumn: start.startColumn + width - 1,
  }
}

/** `row` is 1-based on every row/column op; the marks are zero-based. */
function rowStart(row: unknown): number {
  return typeof row === 'number' && Number.isInteger(row) && row > 0 ? row - 1 : 0
}

function rowBand(row: unknown, count: unknown): RangeBounds | null {
  const start = rowStart(row)
  return {
    startRow: start,
    endRow: start + spanOf(count) - 1,
    startColumn: 0,
    endColumn: MAX_COLUMN,
  }
}

function colStart(column: unknown): number {
  if (typeof column !== 'string' || column.length === 0) return 0
  try {
    return columnIndex(column)
  } catch (error) {
    console.debug('[redact-guard] unparsable column', column, error)
    return 0
  }
}

function colBand(column: unknown, count: unknown): RangeBounds | null {
  const start = colStart(column)
  return { startRow: 0, endRow: MAX_ROW, startColumn: start, endColumn: start + spanOf(count) - 1 }
}

function spanOf(count: unknown): number {
  return typeof count === 'number' && count > 0 ? Math.floor(count) : 1
}

function intersects(mark: CellMark, bounds: RangeBounds): boolean {
  return (
    mark.startRow <= bounds.endRow &&
    mark.endRow >= bounds.startRow &&
    mark.startColumn <= bounds.endColumn &&
    mark.endColumn >= bounds.startColumn
  )
}

function refusalAt(
  mark: CellMark,
  bounds: RangeBounds,
  sheetName: string | undefined,
  reason: string,
): RedactionRefusal {
  // The first cell of the overlap in reading order: the one the model would
  // have to deal with first, and the one the user can go and look at.
  const row = Math.max(mark.startRow, bounds.startRow)
  const column = Math.max(mark.startColumn, bounds.startColumn)
  const address = formatAddress(row, column)
  const where = sheetName ? `${sheetName}!${address}` : address
  return {
    label: mark.label,
    where,
    reason: `${reason} (${where}, ${placeholderSource(mark.label)})`,
  }
}

/**
 * Check one batch of proposed operations. Returns null when the batch may
 * proceed, or the first refusal to hand back.
 *
 * The whole batch is checked before any of it runs, so a refusal leaves the
 * workbook exactly as it was — `propose_operations` is all-or-nothing, and a
 * partially applied batch would leave the reader unable to tell what happened.
 */
export function redactGuardForOps(
  ops: readonly unknown[],
  index: RedactionIndex,
  sheetNameOf: (sheetId: string) => string | undefined = () => undefined,
  sheetIdOf: (sheetName: string) => string | undefined = () => undefined,
): RedactionRefusal | null {
  if (index.isEmpty) return null
  for (const raw of ops) {
    if (typeof raw !== 'object' || raw === null) continue
    const op = raw as Record<string, unknown>
    const name = String(op.op ?? '')
    if (!name) continue
    const sheetId = typeof op.sheetId === 'string' ? op.sheetId : ''
    if (!sheetId) continue

    const reads = READS[name]
    if (reads) {
      for (const site of reads(op)) {
        const readSheetId =
          site.sheetId ?? (site.sheetName !== undefined ? sheetIdOf(site.sheetName) : sheetId)
        if (!readSheetId) continue
        for (const mark of index.marksFor(readSheetId)) {
          if (!intersects(mark, site.bounds)) continue
          return refusalAt(
            mark,
            site.bounds,
            sheetNameOf(readSheetId),
            `${name} reads a cell withheld from the model — copying, moving or referencing ` +
              `it would put its value where the model can see it; leave it out`,
          )
        }
      }
    }

    const marks = index.marksFor(sheetId)
    if (marks.length === 0) continue
    const sheetName = sheetNameOf(sheetId)

    const shift = SHIFTS[name]
    if (shift) {
      const shifts = shift(op)
      for (const mark of marks) {
        if (!shifts(mark)) continue
        return refusalAt(
          mark,
          cellBounds(mark),
          sheetName,
          `${name} would shift a cell withheld from the model out from under its mark`,
        )
      }
      continue
    }

    const write = WRITES[name]
    if (!write) continue
    const bounds = write(op)
    if (!bounds) continue
    for (const mark of marks) {
      if (!intersects(mark, bounds)) continue
      return refusalAt(
        mark,
        bounds,
        sheetName,
        `${name} would write into a cell withheld from the model — it was never shown ` +
          `those values, so it cannot have meant to replace them; write around it instead`,
      )
    }
  }
  return null
}

/// The address a refusal about a shifted mark names: the mark's own top-left.
function cellBounds(mark: CellMark): RangeBounds {
  return {
    startRow: mark.startRow,
    endRow: mark.startRow,
    startColumn: mark.startColumn,
    endColumn: mark.startColumn,
  }
}

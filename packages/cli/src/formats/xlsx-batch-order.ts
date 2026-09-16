import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  shiftIndex,
  shiftSpecForOp,
  type ShiftSpec,
} from '@genoffice/xlsx-gateway/domain/formula-shift'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { StructuralOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { CliError, EXIT, type ErrorHints } from '../result'

export type BatchOp = Record<string, unknown> & { op: string; __index?: number }

/** Row/column ops move every address after them; the batch is read in post-shift coordinates. */
export const SHIFTING_OPS = new Set(['insert_rows', 'delete_rows', 'insert_cols', 'delete_cols'])
/** Ops the in-memory workbook replays to reproduce the gateway's structural pass. */
export const REPLAYED_OPS = new Set([...SHIFTING_OPS, 'add_sheet', 'delete_sheet'])

const AFTER_SHIFT: ErrorHints = {
  reason: 'op_rejected',
  suggestion:
    'place the op after the row/column op it follows and address the shifted cells, or send it in a later batch',
}
const TWO_BATCHES: ErrorHints = {
  reason: 'op_rejected',
  suggestion: 'split the ops into two `sheet apply` calls',
}

/**
 * The gateway replays the structural ops first and then writes the cell
 * edits, which are therefore expressed in final coordinates. Running only
 * the structural ops on a second in-memory workbook gives the state the
 * gateway's replay produces, so the difference to the fully applied batch
 * is exactly the content that was written, already at its shifted address.
 */
export function replayStructure(
  before: WorkbookSnapshot,
  ops: readonly BatchOp[],
): WorkbookSnapshot {
  const adapter = new InMemoryWorkbookAdapter(before)
  for (const [index, op] of ops.entries()) {
    // renames keep the replay's name space in step so add_sheet may take a freed name
    if (!REPLAYED_OPS.has(op.op) && op.op !== 'rename_sheet') continue
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: `genoffice-structure-${index}`,
      baseRevision: adapter.getSnapshot().revision,
      summary: 'structural replay',
      operations: [op as unknown as StructuralOperation],
    })
    adapter.apply(plan)
  }
  return adapter.getSnapshot()
}

/** Shift specs of the row/column ops after `index` on `sheetId`, in batch order. */
export function shiftsAfter(ops: readonly BatchOp[], index: number, sheetId: unknown): ShiftSpec[] {
  const specs: ShiftSpec[] = []
  for (let i = index + 1; i < ops.length; i++) {
    const op = ops[i]!
    if (!SHIFTING_OPS.has(op.op) || op.sheetId !== sheetId) continue
    const spec = shiftSpecForOp(op as unknown as StructuralOperation)
    if (spec) specs.push(spec)
  }
  return specs
}

/** A range after the shifts, or null when it fell entirely inside deleted rows/columns. */
export function shiftBounds(bounds: RangeBounds, specs: readonly ShiftSpec[]): RangeBounds | null {
  let current: RangeBounds | null = bounds
  for (const spec of specs) {
    if (!current) return null
    const row = spec.axis === 'row'
    let start = shiftIndex(row ? current.startRow : current.startColumn, spec)
    let end = shiftIndex(row ? current.endRow : current.endColumn, spec)
    if (start === null && end === null) return null
    if (start === null) start = spec.start
    if (end === null) end = spec.start - 1
    if (end < start) return null
    current = row
      ? { ...current, startRow: start, endRow: end }
      : { ...current, startColumn: start, endColumn: end }
  }
  return current
}

/**
 * Declarative payloads (rules, links, notes, filters, visuals, hidden rows)
 * are saved after the structural pass with the addresses they were given, so
 * one that precedes a row/column op on its sheet would land pre-shift.
 */
export function assertPayloadsFollowShifts(
  ops: readonly BatchOp[],
  payloadOps: ReadonlySet<string>,
  sheetNameOf: (op: BatchOp) => string,
): void {
  for (const [index, op] of ops.entries()) {
    if (!payloadOps.has(op.op) || op.sheetId === undefined) continue
    const later = ops.findIndex(
      (other, j) => j > index && SHIFTING_OPS.has(other.op) && other.sheetId === op.sheetId,
    )
    if (later < 0) continue
    const at = op.__index ?? index
    const laterAt = ops[later]!.__index ?? later
    throw new CliError(
      EXIT.usage,
      `ops[${at}] (${op.op}) addresses ${sheetNameOf(op)} before ops[${laterAt}] (${ops[later]!.op}) shifts its rows/columns`,
      { failures: [{ index: at, op: op.op, error: `must follow ops[${laterAt}]` }] },
      AFTER_SHIFT,
    )
  }
}

/** The copy is cloned from the file's part, so edits to the source sheet in the same batch would be missing from it. */
export function assertDuplicateSourceUntouched(
  ops: readonly BatchOp[],
  contentOp: (op: BatchOp) => boolean,
  sheetNameOf: (op: BatchOp) => string,
): void {
  for (const dup of ops.filter((op) => op.op === 'duplicate_sheet')) {
    const edit = ops.find((op) => contentOp(op) && op.sheetId === dup.sheetId)
    if (!edit) continue
    throw new CliError(
      EXIT.usage,
      `ops rejected: duplicate_sheet of ${sheetNameOf(dup)} cannot share a batch with content ops on that sheet (${edit.op}); the copy is cloned from the file`,
      undefined,
      TWO_BATCHES,
    )
  }
}

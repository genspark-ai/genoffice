import {
  expandToPrimitiveOps,
  formatOpLabel,
  isLayoutOp,
  isStructuralOp,
  layoutOpLabel,
  structuralOpLabel,
  type WorkbookCommandBatch,
} from '../domain/workbook-dsl'
import type { CellState, ChangePlan } from '../domain/workbook.types'

/// Builds an AI change preview against a live (imported) workbook: "before"
/// states come from the current on-screen cells, and the same reader is used
/// at apply time to detect drift since the preview. Operations carry their
/// own sheetId (which may differ from the active sheet), so the reader and
/// the sheet-name lookup are both sheet-addressed.

export type CellReader = (address: string, sheetId: string) => CellState

export function buildLazyChangePlan(
  batch: WorkbookCommandBatch,
  readCell: CellReader,
  sheetNameById: (sheetId: string) => string,
): ChangePlan {
  const cellChanges: ChangePlan['cellChanges'][number][] = []
  const sheetRenames: ChangePlan['sheetRenames'][number][] = []
  const structuralChanges: ChangePlan['structuralChanges'][number][] = []
  const formatChanges: ChangePlan['formatChanges'][number][] = []
  for (const operation of expandToPrimitiveOps(batch.operations, readCell)) {
    if (operation.op === 'format_range') {
      formatChanges.push({
        sheetId: operation.sheetId,
        range: operation.range,
        format: operation.format,
        label: formatOpLabel(operation),
      })
    } else if (isLayoutOp(operation)) {
      structuralChanges.push({ op: operation, label: layoutOpLabel(operation) })
    } else if (isStructuralOp(operation)) {
      structuralChanges.push({ op: operation, label: structuralOpLabel(operation) })
    } else if (operation.op === 'set_cell') {
      cellChanges.push({
        sheetId: operation.sheetId,
        address: operation.address,
        before: readCell(operation.address, operation.sheetId),
        after: { value: operation.value },
      })
    } else if (operation.op === 'set_formula') {
      cellChanges.push({
        sheetId: operation.sheetId,
        address: operation.address,
        before: readCell(operation.address, operation.sheetId),
        after: { value: null, formula: operation.formula },
      })
    } else if (operation.op === 'clear_cell') {
      cellChanges.push({
        sheetId: operation.sheetId,
        address: operation.address,
        before: readCell(operation.address, operation.sheetId),
        after: { value: null },
      })
    } else {
      sheetRenames.push({
        sheetId: operation.sheetId,
        before: sheetNameById(operation.sheetId),
        after: operation.name.trim(),
      })
    }
  }
  return {
    transactionId: batch.transactionId,
    baseRevision: batch.baseRevision,
    cellChanges,
    sheetRenames,
    structuralChanges,
    formatChanges,
    warnings: [],
  }
}

/// True when every planned cell still holds its previewed "before" content.
/// Structural changes have no per-cell before-state to verify; they rely on
/// the sheet-existence check and Univer's own command gating at apply time.
export function planStillMatches(plan: ChangePlan, readCell: CellReader): boolean {
  return plan.cellChanges.every((change) => {
    const current = readCell(change.address, change.sheetId)
    if ((current.formula ?? undefined) !== (change.before.formula ?? undefined)) return false
    // Formula cells only compare formula text: computed values fluctuate with
    // recalcs/dependency changes and are not user-edit conflicts
    if (change.before.formula) return true
    return current.value === change.before.value
  })
}

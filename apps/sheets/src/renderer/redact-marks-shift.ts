import type { CellMark } from './ai/redact'
import { fileToScreen, type Axis } from './view-transform'
import type { StructuralJournalOp } from './edit-journal'

/**
 * Marks are absolute rectangles, so a row or column inserted above one leaves
 * it describing whatever shifted down into its place — and the withheld value
 * it was meant to cover is now somewhere the mark does not reach. A save
 * writes the shifted rectangle, so the leak is not visible in the app: it
 * reappears in the file, on the next open, on someone else's machine.
 *
 * The AI's own `insert_rows` / `insert_cols` are refused while a mark would
 * move (see `ai/redact-guard.ts`). This covers the other way the same thing
 * happens: a person inserting a row from the grid, which no guard sees. The
 * journal already records both — it is how streamed viewports survive a shift —
 * so the marks are replayed through the same stream rather than watched for.
 *
 * A mark whose row or column was deleted is dropped: the value it withheld is
 * gone, and a rectangle clamped onto a neighbour would withhold something the
 * reader never hid.
 */
export function shiftMarksThroughOps(
  marks: readonly CellMark[],
  ops: readonly StructuralJournalOp[],
): CellMark[] {
  const relevant = ops.filter((op) => 'index' in op) as Extract<
    StructuralJournalOp,
    { index: number }
  >[]
  if (relevant.length === 0) return [...marks]
  const out: CellMark[] = []
  for (const mark of marks) {
    const startRow = fileToScreen(relevant, 'row' satisfies Axis, mark.startRow)
    const endRow = fileToScreen(relevant, 'row' satisfies Axis, mark.endRow)
    const startColumn = fileToScreen(relevant, 'column' satisfies Axis, mark.startColumn)
    const endColumn = fileToScreen(relevant, 'column' satisfies Axis, mark.endColumn)
    if (startRow === null || endRow === null) continue
    if (startColumn === null || endColumn === null) continue
    out.push({ ...mark, startRow, endRow, startColumn, endColumn })
  }
  return out
}

/** True when any op would move a mark, so a caller can refuse instead of shifting. */
export function opsWouldMoveMarks(
  marks: readonly CellMark[],
  ops: readonly StructuralJournalOp[],
): boolean {
  const shifted = shiftMarksThroughOps(marks, ops)
  return (
    shifted.length !== marks.length ||
    shifted.some((m, i) => {
      const was = marks[i]
      return (
        was !== undefined &&
        (was.startRow !== m.startRow ||
          was.endRow !== m.endRow ||
          was.startColumn !== m.startColumn ||
          was.endColumn !== m.endColumn)
      )
    })
  )
}

/** The states as the grid shows them, from a sheet list that maps names to ids. */
export function statesOnScreen(
  states: readonly { sheetName: string; marks: readonly CellMark[] }[],
  opsFor: (sheetName: string) => readonly StructuralJournalOp[],
): { sheetName: string; marks: CellMark[] }[] {
  return states.map((state) => {
    const ops = opsFor(state.sheetName)
    return ops.length === 0
      ? { sheetName: state.sheetName, marks: [...state.marks] }
      : { sheetName: state.sheetName, marks: shiftMarksThroughOps(state.marks, ops) }
  })
}

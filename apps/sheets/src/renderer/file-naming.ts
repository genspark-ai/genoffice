/**
 * Model file naming for the sheets editor: the app-side half of
 * `suggestFileName` (`@genoffice/ui`).
 *
 * ## A name only when the reader asks for one
 *
 * `modelSaveAsName` is the whole of it: the ribbon's "Name with AI…" button
 * calls it, and it is the only path from this app into the model for a file
 * name. There is no automatic path — a save never waits on a model round-trip,
 * so a workbook keeps the name it already had and the reader is never asked to
 * approve a name nobody asked for.
 *
 * ## The text is the projected text
 *
 * The call hands `suggestFileName` the output of `modelTextOf`, the workbook's
 * own model-facing projection, and never a raw cell value. A cell the reader
 * hid from the model stays hidden here: the file name is the one thing that
 * reaches the model-bearing path *after* the rest of the workbook, so a value
 * that leaked into it would outlive the redaction meant to contain it — the
 * name is written to disk, synced, and emailed, long after the conversation is
 * gone.
 *
 * ## The index is the app's, not a second one
 *
 * The projection needs a `RedactionIndex`, and a caller that passed the wrong
 * one — or none — would hand the model every value the reader hid. So the
 * index is read through the very `WorkbookReadContext` the AI panel reads
 * (`redactionsOf`), the same ref it keeps current. A second index built for
 * naming would drift from that one the moment a mark was added or cleared, and
 * the drift is silent: the file would simply start carrying a secret.
 */
import { suggestFileName } from '@genoffice/ui'
import {
  modelTextOf,
  type ModelViewCell,
  type ModelViewSheet,
  type ModelViewWorkbook,
} from './ai/redact-view'
import { redactionsOf, type WorkbookReadContext } from './ai/workbook-readers'
import { lazyCellReader } from './univer-sync'
import { lazySheetScreenExtent } from './univer-state'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

/**
 * A file name is read off the first sheets: a sketch of a few hundred cells
 * says what the workbook *is*, and more says nothing a name could use. The
 * budget keeps the projection bounded on the UI thread — a
 * hundred-thousand-cell workbook stops paying for its cells after the first
 * sheet, instead of building tens of megabytes of text that the shell then
 * truncates to 4000 characters anyway.
 */
const NAME_TEXT_BUDGET = { maxLines: 200, maxChars: 8000 } as const

/** Grid cells read for a name, at most: a sketch, not the sheet. */
const NAME_GRID_CAP = { rows: 60, columns: 30, cells: 400 } as const

/**
 * The workbook as a `ModelViewWorkbook`, in the shape the projection reads.
 *
 * An opened file lives in the grid and `lazyWorkbookRef`, not in the adapter,
 * whose snapshot stays the blank boot workbook until an AI plan writes to it.
 * So with a file open the view is read off the grid, bounded by
 * `NAME_GRID_CAP`; the adapter snapshot is the model only for a session
 * workbook, where `WorksheetState` is already `{ id, name, cells }`.
 */
function modelViewOf(ctx: WorkbookReadContext): ModelViewWorkbook {
  const state = ctx.lazyWorkbookRef.current
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  if (!state || !workbook) return ctx.adapterRef.current.getSnapshot()
  let budget = NAME_GRID_CAP.cells
  const sheets: ModelViewSheet[] = []
  for (const worksheet of workbook.getSheets()) {
    const id = worksheet.getSheetId()
    const extent = lazySheetScreenExtent(state, id)
    const rows = Math.min(extent?.rows ?? NAME_GRID_CAP.rows, NAME_GRID_CAP.rows)
    const columns = Math.min(extent?.columns ?? NAME_GRID_CAP.columns, NAME_GRID_CAP.columns)
    const read = lazyCellReader(worksheet)
    const cells: Record<string, ModelViewCell> = {}
    for (let row = 0; row < rows && budget > 0; row += 1) {
      for (let column = 0; column < columns && budget > 0; column += 1) {
        const cell = read(`${columnLabel(column)}${row + 1}`)
        // rich-text cells carry their text only in rawValue
        const value = cell.value ?? cell.rawValue ?? null
        if (!cell.formula && (value === null || value === '')) continue
        cells[`${columnLabel(column)}${row + 1}`] = { value }
        budget -= 1
      }
    }
    sheets.push({ id, name: worksheet.getSheetName(), cells })
    if (budget <= 0) break
  }
  return { sheets }
}

/** The workbook's own cells, as the model may read them. Never a raw value. */
function modelTextOfWorkbook(ctx: WorkbookReadContext): string {
  return modelTextOf(modelViewOf(ctx), redactionsOf(ctx), NAME_TEXT_BUDGET)
}

/**
 * Ask the model to name this workbook, on request.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 *
 * An empty answer means "no name to offer" — no provider, an unreachable
 * provider, a refusal, or a reply that sanitised away to nothing. The caller
 * keeps the name it already had, which is the only acceptable outcome.
 */
export async function modelSaveAsName(ctx: WorkbookReadContext): Promise<string> {
  return suggestFileName({ content: modelTextOfWorkbook(ctx) })
}

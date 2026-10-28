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
import { modelTextOf, type ModelViewWorkbook } from './ai/redact-view'
import { redactionsOf, type WorkbookReadContext } from './ai/workbook-readers'

/**
 * A file name is read off the first sheets: a sketch of a few hundred cells
 * says what the workbook *is*, and more says nothing a name could use. The
 * budget keeps the projection bounded on the UI thread — a
 * hundred-thousand-cell workbook stops paying for its cells after the first
 * sheet, instead of building tens of megabytes of text that the shell then
 * truncates to 4000 characters anyway.
 */
const NAME_TEXT_BUDGET = { maxLines: 200, maxChars: 8000 } as const

/**
 * The workbook as a `ModelViewWorkbook`, in the shape the projection reads.
 *
 * The adapter snapshot is the in-memory model, and `WorksheetState` is already
 * `{ id, name, cells }` — so the snapshot is handed over as-is rather than
 * translated into a second copy that could fall out of step with the grid.
 */
function modelViewOf(ctx: WorkbookReadContext): ModelViewWorkbook {
  return ctx.adapterRef.current.getSnapshot()
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

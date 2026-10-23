/**
 * Model file naming for the sheets editor: the app-side half of `nameForSave`
 * (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `nameForSave` the output of `modelTextOf`, the workbook's own
 * model-facing projection, and never a raw cell value. A cell the reader hid
 * from the model stays hidden here: the file name is the one thing that reaches
 * the model-bearing path *after* the rest of the workbook, so a value that
 * leaked into it would outlive the redaction meant to contain it — the name is
 * written to disk, synced, and emailed, long after the conversation is gone.
 *
 * ## The index is the app's, not a second one
 *
 * The projection needs a `RedactionIndex`, and a caller that passed the wrong
 * one — or none — would hand the model every value the reader hid. So the
 * index is read through the very `WorkbookReadContext` the AI panel reads
 * (`redactionsOf`), the same ref it keeps current. A second index built for
 * naming would drift from that one the moment a mark was added or cleared, and
 * the drift is silent: the file would simply start carrying a secret.
 *
 * ## One attempt per document
 *
 * `firstSaveName` latches before it awaits, so a second ⌘S while the model is
 * still thinking does not start a second turn for the same workbook. The latch
 * is released by `noteDocumentSwapped` when a different workbook is loaded.
 */
import { nameForSave } from '@genoffice/ui'
import { modelTextOf, type ModelViewWorkbook } from './ai/redact-view'
import { redactionsOf, type WorkbookReadContext } from './ai/workbook-readers'

/**
 * One naming turn per workbook, already spent or not.
 *
 * A save that re-asked on every ⌘S would put a model call behind every
 * keystroke the reader made after the first one, which is neither affordable
 * nor wanted: the name is a one-time decision about the workbook, not a
 * response to the save key.
 */
let firstSaveNameSpent = false

/** A different workbook is now open: the next first save may ask again. */
export function noteDocumentSwapped(): void {
  firstSaveNameSpent = false
}

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
  return modelTextOf(modelViewOf(ctx), redactionsOf(ctx))
}

/**
 * The name a workbook that has never been saved should use.
 *
 * The model's name comes first when the reader asked for one. The fallback is
 * empty: sheets' own name for an untitled workbook is the shell's untitled
 * default, retargeted through `autoRenameWorkbook` — a stem this returns as ''
 * leaves the save under exactly the name it would have had anyway, so a
 * declined attempt costs the workbook nothing.
 */
export async function firstSaveName(ctx: WorkbookReadContext): Promise<string> {
  if (firstSaveNameSpent) return ''
  // Latched before the await, not after: a second ⌘S arriving while the model
  // is still thinking must not start a second call for the same workbook.
  firstSaveNameSpent = true
  return nameForSave({
    content: modelTextOfWorkbook(ctx),
    trigger: 'first-save',
    fallback: '',
  })
}

/**
 * Ask the model to name this workbook, on request, for one that already has a
 * name.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 */
export async function modelSaveAsName(
  ctx: WorkbookReadContext,
  filePath: string | null,
): Promise<string> {
  return nameForSave({
    content: modelTextOfWorkbook(ctx),
    trigger: 'manual',
    filePath,
    fallback: '',
  })
}

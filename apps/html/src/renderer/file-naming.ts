/**
 * Model file naming for the HTML editor: the app-side half of `nameForSave`
 * (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `nameForSave` the output of `modelTextOf`, the page's own
 * model-facing projection, and never the raw source. A span the reader hid from
 * the model stays hidden here: the file name is the one thing that reaches the
 * model-bearing path *after* the rest of the document, so a secret that leaked
 * into it would outlive the redaction meant to contain it — the name is written
 * to disk, synced, and emailed, long after the conversation is gone.
 *
 * The parse map travels with the source on purpose. Without it the projection
 * cannot tell which element a mark belongs to and answers "everything is
 * withheld" (`redact-view.ts`); that is the safe direction to fail in, and this
 * module passes the map the editor already has rather than a second one.
 *
 * ## One attempt per document
 *
 * `firstSaveName` latches before it awaits, so a second ⌘S while the model is
 * still thinking does not start a second turn for the same document. The latch
 * is released by `noteDocumentSwapped` when a different document is loaded.
 */
import { nameForSave } from '@genoffice/ui'
import { modelTextOf } from './document/redact-view'
import { deriveAutoFileName, derivePageTitleName } from './document/auto-name'
import type { ParseMap } from './document/parse-map'

/**
 * One naming turn per document, already spent or not.
 *
 * A save that re-asked on every ⌘S would put a model call behind every
 * keystroke the reader made after the first one, which is neither affordable
 * nor wanted: the name is a one-time decision about the document, not a
 * response to the save key.
 */
let firstSaveNameSpent = false

/** A different document is now open: the next first save may ask again. */
export function noteDocumentSwapped(): void {
  firstSaveNameSpent = false
}

/** The page's own text, as the model may read it. Never the raw source. */
function modelTextOfDocument(source: string, map: ParseMap): string {
  return modelTextOf(source, map)
}

/**
 * The name this page's own content would take, in the order the app already
 * prefers: its `<title>`, else its first heading, else the first words. The
 * prompt the reader typed is the caller's business and rides alongside this
 * only as a fallback for a page that has no title of its own.
 */
function localName(source: string, promptName: string | null): string {
  return derivePageTitleName(source) || promptName || deriveAutoFileName(source)
}

/**
 * The name a never-saved page's first save should use.
 *
 * The model's name comes first when the reader asked for one; the page's own
 * title stays the fallback either way. `nameForSave` returns the fallback
 * unchanged when naming is switched off, declines, times out, or has nothing
 * to work from, so a failed attempt costs the document nothing — the save
 * proceeds under the name it already had.
 */
export async function firstSaveName(
  source: string,
  map: ParseMap,
  promptName: string | null,
): Promise<string> {
  const local = localName(source, promptName)
  if (firstSaveNameSpent) return local
  // Latched before the await, not after: a second ⌘S arriving while the model
  // is still thinking must not start a second call for the same document.
  firstSaveNameSpent = true
  const stem = await nameForSave({
    content: modelTextOfDocument(source, map),
    trigger: 'first-save',
    fallback: '',
  })
  return stem || local
}

/**
 * Ask the model to name this page, on request, for a page that already has a
 * name.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 */
export async function modelSaveAsName(
  source: string,
  map: ParseMap,
  filePath: string | null,
): Promise<string> {
  return nameForSave({
    content: modelTextOfDocument(source, map),
    trigger: 'manual',
    filePath,
    fallback: '',
  })
}

/**
 * Model file naming for the markdown editor: the app-side half of
 * `nameForSave` (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `nameForSave` the output of `modelTextOf`, the editor's own
 * model-facing projection, and never `textContent` or a serialized document. A
 * selection the reader hid from the model stays hidden here: the file name is
 * the one thing that reaches the model-bearing path *after* the rest of the
 * document, so a secret that leaked into it would outlive the redaction meant
 * to contain it — the name is written to disk, synced, and emailed, long after
 * the conversation is gone.
 *
 * ## One attempt per document
 *
 * `firstSaveName` latches before it awaits, so a second ⌘S while the model is
 * still thinking does not start a second turn for the same document. The latch
 * is released by `noteDocumentSwapped` when a different document is loaded.
 */
import type { Editor } from '@tiptap/core'
import { nameForSave } from '@genoffice/ui'
import { modelTextOf } from './editor/redact'
import { deriveAutoFileName } from './auto-file-name'

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

/** The document's own text, as the model may read it. Never the raw document. */
function modelTextOfDocument(editor: Editor): string {
  return modelTextOf(editor.state.doc.toJSON() as never)
}

/**
 * The name a never-saved document's first save should use.
 *
 * The model's name comes first when the reader asked for one; the document's
 * own first heading stays the fallback either way. `nameForSave` returns the
 * fallback unchanged when naming is switched off, declines, times out, or has
 * nothing to work from, so a failed attempt costs the document nothing — the
 * save proceeds under the name it already had.
 */
export async function firstSaveName(editor: Editor): Promise<string> {
  const local = deriveAutoFileName(editor)
  if (firstSaveNameSpent) return local
  // Latched before the await, not after: a second ⌘S arriving while the model
  // is still thinking must not start a second call for the same document.
  firstSaveNameSpent = true
  const stem = await nameForSave({
    content: modelTextOfDocument(editor),
    trigger: 'first-save',
    fallback: '',
  })
  return stem || local
}

/**
 * Ask the model to name this document, on request, for a document that already
 * has a name.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 */
export async function modelSaveAsName(editor: Editor, filePath: string | null): Promise<string> {
  return nameForSave({
    content: modelTextOfDocument(editor),
    trigger: 'manual',
    filePath,
    fallback: '',
  })
}

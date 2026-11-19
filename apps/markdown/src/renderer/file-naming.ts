/**
 * Model file naming for the markdown editor: the app-side half of
 * `suggestFileName` (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `suggestFileName` the output of `modelTextOf`, the editor's
 * own model-facing projection, and never `textContent` or a serialized document.
 * A selection the reader hid from the model stays hidden here: the file name is
 * the one thing that reaches the model-bearing path *after* the rest of the
 * document, so a secret that leaked into it would outlive the redaction meant
 * to contain it — the name is written to disk, synced, and emailed, long after
 * the conversation is gone.
 *
 * ## A name is asked for, never taken
 *
 * There is no automatic path: a save never waits on a model round-trip, and no
 * latch decides whether a document is allowed to be named. A save offers the
 * document's own name — its first heading, else its first words — and only the
 * reader's explicit request reaches the model.
 */
import type { Editor } from '@tiptap/core'
import { suggestFileName } from '@genoffice/ui'
import { modelTextOf } from './editor/redact'
import { deriveAutoFileName } from './auto-file-name'

/** The document's own text, as the model may read it. Never the raw document. */
function modelTextOfDocument(editor: Editor): string {
  return modelTextOf(editor.state.doc.toJSON() as never)
}

/**
 * Ask the model to name this document, on request.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 *
 * An empty answer is not a failure to report: the model may have no provider,
 * may be unreachable, may decline, or may have nothing to work from, and in
 * every one of those cases the reader keeps the name the document already had.
 */
export async function modelSaveAsName(editor: Editor): Promise<string> {
  return suggestFileName({ content: modelTextOfDocument(editor) })
}

/**
 * The name a save should offer, or nothing when it should offer none.
 *
 * This is the whole decision, in one place, because it is the rule the reader
 * feels: a document that already has a name keeps it, and only the reader's own
 * request for a name moves it.
 *
 * - a name the caller proposed outright (the ribbon's "Name with AI") wins:
 *   the reader asked for that name, so nothing else gets a say;
 * - a document that already has a path keeps its name — the save path never
 *   renames a file that is open elsewhere, referenced from a note, or synced;
 * - a save that already carries a name of its own (the AI auto-naming path)
 *   is not asked again;
 * - otherwise the document names itself, from its own first heading, else its
 *   first words. The model is not consulted: a save must not wait on a model
 *   round-trip to learn what the document already says about itself.
 */
export async function proposedSaveName(
  editor: Editor,
  options: { filePath: string | null; suggestedName?: string; proposedName?: string },
): Promise<string | undefined> {
  if (options.proposedName) return options.proposedName
  if (options.filePath) return undefined
  if (options.suggestedName) return undefined
  return deriveAutoFileName(editor) || undefined
}

/**
 * Model file naming for the slides editor: the app-side half of `nameForSave`
 * (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `nameForSave` the output of `textForModel`, the deck's own
 * model-facing projection, and never a raw run of glyph text. A span the reader
 * hid from the model stays hidden here: the file name is the one thing that
 * reaches the model-bearing path *after* the rest of the deck, so a secret that
 * leaked into it would outlive the redaction meant to contain it — the name is
 * written to disk, synced, and emailed, long after the conversation is gone.
 *
 * `textForModel`'s `forDisplay` defaults to false, so this module gets the
 * redacted view by omitting the argument. That default is the whole safety
 * property here: passing `true` would be the one way to leak, and nothing in
 * this file does.
 *
 * ## One attempt per document
 *
 * `firstSaveName` latches before it awaits, so a second ⌘S while the model is
 * still thinking does not start a second turn for the same deck. The latch is
 * released by `noteDocumentSwapped` when a different deck is loaded.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import { nameForSave } from '@genoffice/ui'
import { textForModel } from './ai/redact-view'

/**
 * One naming turn per deck, already spent or not.
 *
 * A save that re-asked on every ⌘S would put a model call behind every
 * keystroke the reader made after the first one, which is neither affordable
 * nor wanted: the name is a one-time decision about the deck, not a response to
 * the save key.
 */
let firstSaveNameSpent = false

/** A different deck is now open: the next first save may ask again. */
export function noteDocumentSwapped(): void {
  firstSaveNameSpent = false
}

/**
 * The deck's own text, as the model may read it, one slide per paragraph.
 *
 * Every node of every slide goes through the same projection, so a withheld
 * shape, a withheld table cell and a withheld picture are all swapped for their
 * markers rather than dropping out — the model can see that something is there
 * and write around it.
 */
function modelTextOfDeck(slides: readonly RenderSlide[]): string {
  const blocks: string[] = []
  for (const [index, slide] of slides.entries()) {
    const text = slide.nodes
      .map((node) => textForModel(node))
      .filter((line) => line.trim())
      .join('\n')
    blocks.push(`## Slide ${index + 1}\n${text}`)
  }
  return blocks.join('\n\n')
}

/**
 * The name a never-saved deck's first save should use.
 *
 * The model's name comes first when the reader asked for one. The fallback is
 * empty: slides' own deterministic name for an untitled deck is the localized
 * "Untitled presentation", and that lives in the main process's `pickDraftPath`
 * — a stem this returns as '' leaves the save under exactly the name it would
 * have had anyway, so a declined attempt costs the deck nothing.
 */
export async function firstSaveName(slides: readonly RenderSlide[]): Promise<string> {
  if (firstSaveNameSpent) return ''
  // Latched before the await, not after: a second ⌘S arriving while the model
  // is still thinking must not start a second call for the same deck.
  firstSaveNameSpent = true
  return nameForSave({
    content: modelTextOfDeck(slides),
    trigger: 'first-save',
    fallback: '',
  })
}

/**
 * Ask the model to name this deck, on request, for a deck that already has a
 * name.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 */
export async function modelSaveAsName(
  slides: readonly RenderSlide[],
  filePath: string | null,
): Promise<string> {
  return nameForSave({
    content: modelTextOfDeck(slides),
    trigger: 'manual',
    filePath,
    fallback: '',
  })
}

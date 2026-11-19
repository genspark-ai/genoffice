/**
 * Model file naming for the slides editor: the app-side half of
 * `suggestFileName` (`@genoffice/ui`).
 *
 * ## The text is the projected text
 *
 * Every call hands `suggestFileName` the output of `textForModel`, the deck's
 * own model-facing projection, and never a raw run of glyph text. A span the
 * reader hid from the model stays hidden here: the file name is the one thing
 * that reaches the model-bearing path *after* the rest of the deck, so a secret
 * that leaked into it would outlive the redaction meant to contain it — the name
 * is written to disk, synced, and emailed, long after the conversation is gone.
 *
 * `textForModel`'s `forDisplay` defaults to false, so this module gets the
 * redacted view by omitting the argument. That default is the whole safety
 * property here: passing `true` would be the one way to leak, and nothing in
 * this file does.
 *
 * ## Only when the reader asks
 *
 * There is no automatic path here. A deck is named because the reader asked for
 * it — the ribbon's "Name with AI…" — and the answer lands in the Save As
 * dialog, where it can be read and edited before anything is written. A first
 * save never waits on a model round-trip: an untitled deck keeps the app's own
 * name, the localized "Untitled presentation" the main process produces in
 * `pickDraftPath`.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import { suggestFileName } from '@genoffice/ui'
import { textForModel } from './ai/redact-view'

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
 * Ask the model to name this deck, on request.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 *
 * An empty answer means there is no name to offer — no provider, an
 * unreachable one, a refusal, or a reply that sanitised away to nothing — and
 * the reader keeps the name the deck already had.
 */
export async function modelSaveAsName(slides: readonly RenderSlide[]): Promise<string> {
  return suggestFileName({ content: modelTextOfDeck(slides) })
}

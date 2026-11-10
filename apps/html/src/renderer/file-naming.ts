/**
 * Model file naming for the HTML editor: the app-side half of
 * `suggestFileName` (`@genoffice/ui`).
 *
 * ## The reader asks, the model answers
 *
 * The model is asked for a name in exactly one place: the ribbon's
 * "Name with AI…" button. The answer prefills the Save As dialog, where the
 * reader can read it, change it or cancel — nothing is written before they
 * decide. There is no automatic path: a first save never waits on a model
 * round-trip, it takes the page's own name (`localFileName`) and is done.
 *
 * ## The text is the projected text
 *
 * Every call hands `suggestFileName` the output of `modelTextOf`, the page's
 * own model-facing projection, and never the raw source. A span the reader hid
 * from the model stays hidden here: the file name is the one thing that reaches
 * the model-bearing path *after* the rest of the document, so a secret that
 * leaked into it would outlive the redaction meant to contain it — the name is
 * written to disk, synced, and emailed, long after the conversation is gone.
 *
 * The parse map travels with the source on purpose. Without it the projection
 * cannot tell which element a mark belongs to and answers "everything is
 * withheld" (`redact-view.ts`); that is the safe direction to fail in, and this
 * module passes the map the editor already has rather than a second one.
 *
 * ## Nothing to offer is not a failure
 *
 * `suggestFileName` answers '' when there is no name to offer: no provider, an
 * unreachable one, a refusal, or a reply that sanitised away to nothing. The
 * caller keeps the name the document already had, which is the only acceptable
 * outcome — a rename that costs the reader their file name is worse than no
 * rename at all.
 */
import { suggestFileName } from '@genoffice/ui'
import { modelTextOf } from './document/redact-view'
import { deriveAutoFileName, derivePageTitleName } from './document/auto-name'
import type { ParseMap } from './document/parse-map'

/** The page's own text, as the model may read it. Never the raw source. */
function modelTextOfDocument(source: string, map: ParseMap): string {
  return modelTextOf(source, map)
}

/**
 * The name this page's own content would take, in the order the app already
 * prefers: its `<title>`, else the name the reader's first request suggested,
 * else the first words of the page.
 *
 * This is the name a save uses when the reader did not ask the model for one.
 * It costs nothing and waits for nobody, which is what a save needs: a save
 * that waited on a model round-trip would put a network call behind every
 * ⌘S the reader made.
 */
export function localFileName(source: string, promptName: string | null): string {
  return derivePageTitleName(source) || promptName || deriveAutoFileName(source)
}

/**
 * Ask the model to name this page, on request.
 *
 * The result is the stem for the Save As dialog rather than a rename: moving a
 * file that may be open elsewhere, referenced from a note, or synced is not
 * something to do behind the reader's back. The dialog is the confirmation.
 */
export async function modelSaveAsName(source: string, map: ParseMap): Promise<string> {
  return suggestFileName({ content: modelTextOfDocument(source, map) })
}

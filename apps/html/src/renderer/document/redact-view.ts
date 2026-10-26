/**
 * The model-facing view of a page: the source as the model may read it, with
 * every withheld span written out as its `{{label}}` instead of its contents.
 *
 * This is a thin, total front for `RedactionProjection`, which is where the
 * projection itself lives. That module is also the guard's: `redact-guard.ts`
 * needs offsets in both directions, because an edit composed against the view
 * has to be checked against the real source. A caller that only wants the text
 * should not have to know that, and should not be able to reach for the raw
 * string while reaching for it.
 *
 * ## Why it fails closed
 *
 * `collectWithheld` needs a parse map: it is what tells a mark which element it
 * belongs to, and what keeps a `/*gx:redact:*\/` in ordinary prose from being
 * honoured. A map that is missing or damaged cannot answer that, so the
 * withheld spans are unknown — and unknown is not "none". Returning the source
 * here would hand the model exactly the values the reader hid, which is why the
 * unrecognised case answers with one `{{private}}` instead. It reads as a
 * refusal, and it is visible.
 */

import { buildProjection, placeholderSource } from '../ai/redact'
import type { ParseMap } from './parse-map'

/** The label shown when a file's withheld spans could not be established. */
export const UNKNOWN_LABEL = 'private'

/**
 * The text of a page as the model may read it.
 *
 * Total by construction: it returns for every input, and never the raw source
 * where a mark might have stood.
 */
export function modelTextOf(source: string, map: ParseMap | null | undefined): string {
  if (typeof source !== 'string' || source === '') return ''
  try {
    return buildProjection(source, map as ParseMap).view
  } catch {
    // No map, no parse, or a map whose shape the collector does not recognise.
    // The withheld spans are unknown rather than absent, so the whole document
    // reads as withheld.
    return placeholderSource(UNKNOWN_LABEL)
  }
}

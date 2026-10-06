/**
 * Withheld spans for the document editor: text the reader keeps but does not
 * want a model to read.
 *
 * A span is an ordinary inline mark over the real words, not a replacement for
 * them. The document and the .docx on disk still hold the actual text — losing
 * the reader's own data to hide it from a model would be the worse failure. On
 * disk the mark becomes a character border with a custom run property, so Word
 * renders the words with a line under them and the label rides along inside
 * the file. What changes is the *model's* view: when a block is serialized for
 * a request, the span is written as `{{label}}` instead of its contents.
 *
 * Two consequences follow from the mark sitting on real text:
 *
 * - the model can split it, the way it could not split an atom. A model
 *   answering `{{cli}}ent pho{{ne}}` is plausible, so every model edit is
 *   checked before it is allowed into the document.
 * - an image can be withheld too. Images carry no text, so there is nothing to
 *   redact — the point is that the model is not told a path that would let it
 *   fetch the picture.
 *
 * The rules about markers, and the prompt they produce, are shared with the
 * other four editors and live in `@genoffice/agent-core/redact-core`. What
 * stays here is what is specific to a document: the mark's name, and the two
 * consequences above, which is what `middle` in the spec below is for.
 */

import {
  placeholderInstruction as buildInstruction,
  type PlaceholderSpec,
} from '@genoffice/agent-core/redact-core'

export const REDACT_MARK = 'redaction'

/** What a document's markers need that a sheet's or a deck's do not. */
const DOC_SPEC: PlaceholderSpec = {
  subject: 'document',
  opening:
    'Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
  boundary: 'a line break',
  middle: [
    'Write the prose around them as if each stood for the words it replaces, so "call {{客户电话}}" reads as a natural instruction to phone someone.',
    'If a request needs what a placeholder hides, write around it rather than guessing.',
    'A placeholder may stand in for a picture: treat it as an image that was withheld and write around it.',
  ],
}

export function placeholderInstruction(labels: readonly string[]): string {
  return buildInstruction(labels, DOC_SPEC)
}

export {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
  type PlaceholderIssue,
} from '@genoffice/agent-core/redact-core'

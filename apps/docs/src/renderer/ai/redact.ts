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
 *   before it is allowed into the document.
 * - an image can be withheld too. Images carry no text, so there is nothing to
 *   redact — the point is that the model is not told a path that would let it
 *   fetch the picture.
 */

/** The form the model sees. Braces make a marker recognisable in a reply; a
 *  label containing them would be ambiguous, so they are stripped. */
const OPEN = '{{'
const CLOSE = '}}'

/** Long enough to name what a span stands for, short enough to stay readable. */
export const MAX_LABEL_LENGTH = 40

export const REDACT_MARK = 'redaction'

/** Clean a label for storage and for the model prompt. */
export function sanitizeLabel(raw: string): string {
  return raw
    .replace(/[{}<>="']/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_LABEL_LENGTH)
    .trim()
}

/** The literal text the model is shown for a span. */
export function placeholderSource(label: string): string {
  return `${OPEN}${sanitizeLabel(label) || 'private'}${CLOSE}`
}

export function isWholePlaceholder(text: string): boolean {
  const t = text.trim()
  return t.startsWith(OPEN) && t.endsWith(CLOSE) && t.length > OPEN.length + CLOSE.length
}

export function readPlaceholderLabel(text: string): string | null {
  if (!isWholePlaceholder(text)) return null
  return text.trim().slice(OPEN.length, -CLOSE.length)
}

export interface PlaceholderIssue {
  found: string
  expected: string
  reason: 'split' | 'missing' | 'unknown'
}

function scan(text: string): string[] {
  return [...text.matchAll(/\{\{[^{}]*\}\}/g)].map((m) => m[0])
}

/**
 * Braces a malformed marker left behind, with well-formed markers masked out
 * first. `{{a}` yields `{{`; a sentence containing `{{x}}` yields nothing.
 */
function leftoverBraces(text: string): string[] {
  const masked = text.replace(/\{\{[^{}]*\}\}/g, (m) => ' '.repeat(m.length))
  const out: string[] = []
  for (const m of masked.matchAll(/\{\{|\}\}/g)) out.push(m[0])
  for (const m of masked.matchAll(/(?<!\{)\{[^{}]*\}(?!\})/g)) out.push(m[0])
  return out
}

/**
 * Compare the markers in a model's answer against the ones it was given.
 *
 * Order is ignored — a model may rewrite the sentence — but the count and the
 * exact spelling are not. Accepting anything else would write a mangled marker
 * into a document that is about to be saved, and the file would look fine.
 */
export function checkPlaceholders(before: string, after: string): PlaceholderIssue[] {
  const issues: PlaceholderIssue[] = []
  for (const stray of leftoverBraces(after)) {
    issues.push({ found: stray, expected: '', reason: 'split' })
  }

  const gotCounts = new Map<string, number>()
  for (const m of scan(after)) gotCounts.set(m, (gotCounts.get(m) ?? 0) + 1)
  const wantCounts = new Map<string, number>()
  for (const m of scan(before)) wantCounts.set(m, (wantCounts.get(m) ?? 0) + 1)

  for (const [marker, want] of wantCounts) {
    const have = gotCounts.get(marker) ?? 0
    if (have < want) {
      issues.push({ found: have ? marker : '', expected: marker, reason: 'missing' })
    }
  }
  for (const [marker, have] of gotCounts) {
    if (!wantCounts.has(marker)) {
      issues.push({ found: marker, expected: '', reason: 'unknown' })
      continue
    }
    if (have > (wantCounts.get(marker) ?? 0)) {
      issues.push({ found: marker, expected: marker, reason: 'unknown' })
    }
  }
  return issues
}

export function collectPlaceholders(text: string): string[] {
  return scan(text)
}

/** The instruction the model gets, listing the markers this document contains. */
export function placeholderInstruction(labels: readonly string[]): string {
  const list = [...new Set(labels)].map((l) => `- {{${sanitizeLabel(l) || 'private'}}}`).join('\n')
  return [
    '## Private placeholders',
    'This document contains {{...}} placeholders. Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
    '',
    'Treat every placeholder as one indivisible object:',
    '- Copy it character for character — same letters, same order, same spacing.',
    '- Never split it across a line break or put a space inside it.',
    '- Never merge two into one, never split one into several, never reorder them.',
    '- Never rename, translate, re-case, expand or shorten it.',
    '- Never drop one, and never add a placeholder that was not already there.',
    '',
    'Write the prose around them as if each stood for the words it replaces, so "call {{客户电话}}" reads as a natural instruction to phone someone.',
    'If a request needs what a placeholder hides, write around it rather than guessing.',
    'A placeholder may stand in for a picture: treat it as an image that was withheld and write around it.',
    '',
    'The placeholders in this document:',
    list,
  ].join('\n')
}

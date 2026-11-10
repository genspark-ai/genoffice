/**
 * Withholding: the rules every editor shares.
 *
 * ## What this is for
 *
 * A reader can keep a span in their document and stop a model from reading it.
 * What the model sees instead is a `{{label}}` marker, and the marker is the
 * only thing standing between the reader's data and the answer — so the rules
 * about how a model may treat it are the feature, not documentation of it.
 *
 * Five editors need those rules and each grew its own copy. The copies were
 * already drifting: three spellings of `checkPlaceholders` that meant the same
 * thing, and five prompt texts that had quietly diverged. A reader who learns
 * the rules in one editor should not meet a different set in the next.
 *
 * ## What stays app-shaped
 *
 * What a placeholder can stand for, and what follows from that, differ per
 * editor — a withheld picture in a deck is not a withheld spreadsheet value.
 * So the prompt is assembled here from a spec the editor supplies, rather than
 * forked per editor. The shared rules are written once; the differences are
 * visible at the call site instead of hiding in a near-identical function.
 *
 * ## What the model is told, and what it is not
 *
 * The rules here govern markers, not data. An editor that withholds something
 * with no marker in it — a cell excluded from a statistic, a picture the
 * renderer only tints — is saying nothing to the model, and only that editor
 * can close the gap. Nothing in this file makes a promise on their behalf.
 */

/** The form the model sees. Braces make a marker recognisable in a reply; a
 *  label containing them would be ambiguous, so they are stripped. */
const OPEN = '{{'
const CLOSE = '}}'

/** Long enough to name what a span stands for, short enough to stay readable. */
export const MAX_LABEL_LENGTH = 40

/**
 * Characters no label may contain.
 *
 * `*` and `/` are not illegal in a document — they are illegal in a comment,
 * which is where a sheet's label ends up. An editor whose carrier needs them
 * stripped says so rather than every editor paying for the widest set.
 */
const LABEL_STRIP = /[{}<>="']/g
const LABEL_STRIP_COMMENT = /[{}<>="'*/]/g

export interface SanitizeLabelOptions {
  /** strip the characters an XML comment cannot carry, for a comment carrier */
  forComment?: boolean
}

/** Clean a label for storage and for the model prompt. */
export function sanitizeLabel(raw: string, options: SanitizeLabelOptions = {}): string {
  return raw
    .replace(options.forComment ? LABEL_STRIP_COMMENT : LABEL_STRIP, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, MAX_LABEL_LENGTH)
    .trim()
}

/** What the model reads in place of the withheld words. */
export function placeholderSource(label: string): string {
  return `${OPEN}${sanitizeLabel(label) || 'private'}${CLOSE}`
}

/** True when the text is nothing but one whole marker. */
export function isWholePlaceholder(text: string): boolean {
  const t = text.trim()
  return t.startsWith(OPEN) && t.endsWith(CLOSE) && t.length > OPEN.length + CLOSE.length
}

/** The label a whole marker stands for, or null when the text is not one. */
export function readPlaceholderLabel(text: string): string | null {
  const t = text.trim()
  if (!isWholePlaceholder(t)) return null
  return sanitizeLabel(t.slice(OPEN.length, t.length - CLOSE.length)) || null
}

/** Every whole marker in the text, in order. */
export function collectPlaceholders(text: string): string[] {
  return [...text.matchAll(/\{\{[^{}]*\}\}/g)].map((m) => m[0])
}

export interface PlaceholderIssue {
  /** what the model's answer actually contains at that spot; '' when absent */
  found: string
  /** the marker the model was given, '' when it invented one */
  expected: string
  reason: 'missing' | 'unknown' | 'split'
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
 * Every one of the original copies meant this; three of them spelled it
 * differently, which is the kind of drift that hides a real difference later.
 */
export function checkPlaceholders(before: string, after: string): PlaceholderIssue[] {
  const issues: PlaceholderIssue[] = []
  for (const stray of leftoverBraces(after)) {
    issues.push({ found: stray, expected: '', reason: 'split' })
  }

  const gotCounts = new Map<string, number>()
  for (const m of collectPlaceholders(after)) gotCounts.set(m, (gotCounts.get(m) ?? 0) + 1)
  const wantCounts = new Map<string, number>()
  for (const m of collectPlaceholders(before)) wantCounts.set(m, (wantCounts.get(m) ?? 0) + 1)

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

/** What the reader's document is called in this editor, and what differs. */
export interface PlaceholderSpec {
  /** "document", "workbook", "presentation", "page" */
  subject: string
  /** what a placeholder stands in for, after "This {subject} contains {{...}}" */
  opening: string
  /** what a marker may not be split across: "a line break", "a cell boundary" */
  boundary: string
  /**
   * The editor's own consequences, in order: what a placeholder can stand in
   * for here, and what the model must not infer from its absence.
   */
  middle: readonly string[]
}

/**
 * The prompt section every editor sends, assembled from the shared rules and
 * the editor's own spec.
 */
export function placeholderInstruction(labels: readonly string[], spec: PlaceholderSpec): string {
  const list = [...new Set(labels)].map((l) => `- ${placeholderSource(l)}`).join('\n')
  return [
    '## Private placeholders',
    `This ${spec.subject} contains {{...}} placeholders. ${spec.opening}`,
    '',
    'Treat every placeholder as one indivisible object:',
    '- Copy it character for character — same letters, same order, same spacing.',
    `- Never split it across ${spec.boundary} or put a space inside it.`,
    '- Never merge two into one, never split one into several, never reorder them.',
    '- Never rename, translate, re-case, expand or shorten it.',
    '- Never drop one, and never add a placeholder that was not already there.',
    '',
    ...spec.middle,
    '',
    `The placeholders in this ${spec.subject}:`,
    list,
  ].join('\n')
}

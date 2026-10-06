/**
 * Withheld spans: text the reader keeps but does not want a model to read.
 *
 * A withheld span is an ordinary inline **mark** over the real text, not a
 * replacement for it. The document and the file on disk still hold the actual
 * words — losing the reader's own data to hide it from a model would be a
 * worse failure than the leak. What changes is the *model's* view: when the
 * document is serialized for a request, each span is written out as
 * `{{label}}` instead of its contents, so the secret never crosses the wire.
 *
 * Two consequences follow from the mark being over real text rather than being
 * one:
 *
 * - the model can split it. `atom` guarded a single node; a span of five
 *   characters has five positions, and a model answering with
 *   `{{cli}}ent pho{{ne}}` is a plausible thing to write. So every model edit is
 *   checked before it is allowed in, the same way the node version was.
 * - a round trip is exact. The file contains what the reader typed, so opening
 *   it again restores the same document with the span still marked.
 */

import {
  placeholderInstruction as buildInstruction,
  placeholderSource,
  type PlaceholderSpec,
} from '@genoffice/agent-core/redact-core'

import {
  redactTextBetween as sharedRedactTextBetween,
  type WithheldNode,
} from '@genoffice/agent-core/redact-range'

/** The mark's node name, shared so the mark, the model view and the guard agree. */
export const REDACT_MARK = 'redaction'

/** What this editor's markers need that another's do not. */
const MARKDOWN_SPEC: PlaceholderSpec = {
  subject: 'document',
  opening:
    'Each one stands in for text the reader has deliberately withheld from you; you cannot see what is inside and that is the point.',
  boundary: 'a line break',
  middle: [
    'Write the prose around them as if each stood for the words it replaces, so "call {{客户电话}}" reads as a natural instruction to phone someone.',
    'If a request needs what a placeholder hides, write around it rather than guessing.',
  ],
}

export function placeholderInstruction(labels: readonly string[]): string {
  return buildInstruction(labels, MARKDOWN_SPEC)
}

/**
 * The model-facing view of a document: every withheld span becomes its marker
 * and the text it covers disappears.
 *
 * This runs on the JSON *before* it is handed to the markdown serializer, not
 * after. A mark hanging off a text node is not something the serializer routes
 * through the mark's own `renderMarkdown`, so replacing at that level would
 * silently do nothing; walking the JSON here is explicit and testable.
 *
 * `redactJson` is total: any node shape it does not recognise is passed
 * through untouched, so an unexpected document cannot make a secret leak by
 * throwing.
 */
export function redactJson<T>(json: T): T {
  return walk(json, false) as T
}

function walk(node: unknown, insideMarked: boolean): unknown {
  if (Array.isArray(node)) return node.map((child) => walk(child, insideMarked))
  if (!node || typeof node !== 'object') return node
  const obj = node as Record<string, unknown>

  const marks = obj.marks as Array<{ type?: string; attrs?: Record<string, unknown> }> | undefined
  const marked = marks?.some((m) => m?.type === REDACT_MARK)
  if (insideMarked) {
    // text under a withheld span never leaves the machine
    if (typeof obj.text === 'string') return { ...obj, text: '' }
  }
  if (marked && typeof obj.text === 'string') {
    const label = String(marks?.find((m) => m?.type === REDACT_MARK)?.attrs?.label ?? 'private')
    return { ...obj, text: placeholderSource(label), marks: withoutRedact(marks) }
  }
  const out: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(obj)) {
    out[key] = key === 'content' ? walk(value, insideMarked) : value
  }
  return out
}

/** the other marks on a span (bold, italic) are the reader's, not the model's */
function withoutRedact(marks: Array<{ type?: string }> | undefined) {
  if (!marks) return undefined
  const kept = marks.filter((m) => m?.type !== REDACT_MARK)
  return kept.length ? kept : undefined
}

/** The labels currently withheld in a document, for the model's instruction. */
export function redactLabels(node: unknown): string[] {
  const out: string[] = []
  const visit = (n: unknown) => {
    if (Array.isArray(n)) {
      n.forEach(visit)
      return
    }
    if (!n || typeof n !== 'object') return
    const obj = n as Record<string, unknown>
    const marks = obj.marks as Array<{ type?: string; attrs?: Record<string, unknown> }> | undefined
    const hit = marks?.find((m) => m?.type === REDACT_MARK)
    if (hit) out.push(String(hit.attrs?.label ?? 'private'))
    if (Array.isArray(obj.content)) obj.content.forEach(visit)
  }
  visit(node)
  return out
}

/**
 * The text of a document as the model may read it: every withheld span is
 * written out as its marker instead of its contents.
 *
 * ProseMirror's `textContent` walks straight through a mark, so reading the
 * document that way hands the model exactly what the reader withheld. This
 * walks the child list instead, which is the only way to see the marks.
 *
 * Every path that puts document text into a request goes through here — the
 * per-turn context, the selection line and the block serializer — because the
 * failure is silent and total: one call site left reading `textContent` is a
 * whole-document leak.
 */
export function modelTextOf(node: ProseNode): string {
  // a withheld span is replaced by its marker, not blanked: the model still
  // needs to know one is there, or it reads a gap as a missing word
  if (isWithheld(node)) return redactionLabelOf(node) ?? ''
  if (isLeafText(node)) return String(node.text ?? '')
  const children = node.content
  if (!Array.isArray(children)) return ''
  let out = ''
  for (const child of children as ProseNode[]) {
    out += modelTextOf(child)
    // a block boundary is a line break, but not a trailing one: the caller
    // joins this with other text and a stray newline reads as a blank line
    if (isBlock(child) && child !== children[children.length - 1]) out += '\n'
  }
  return out
}

/**
 * The text of a range, as the model may read it.
 *
 * A thin wrapper over the shared walker, so the options are stated once here
 * rather than at every call site. Docs has the same wrapper over the same code,
 * which is the point of it living in a package rather than in five editors.
 */
export function redactTextBetween(node: ProseNode, from: number, to: number): string {
  return sharedRedactTextBetween(node as unknown as WithheldNode, from, to, {
    markName: REDACT_MARK,
    marker: placeholderSource,
  })
}

/** The marker this node stands for, or null when it is not withheld. */
export function redactionLabelOf(node: ProseNode): string | null {
  if (!isWithheld(node)) return null
  const mark = (node.marks ?? []).find((m) => markName(m) === REDACT_MARK)
  const label =
    mark && 'attrs' in mark ? (mark.attrs as { label?: unknown } | undefined)?.label : undefined
  return placeholderSource(typeof label === 'string' ? label : 'private')
}

/** Every label withheld in a document, for the model's instruction. */
export function redactLabelsOf(node: ProseNode): string[] {
  const out: string[] = []
  const visit = (n: ProseNode) => {
    if (isWithheld(n)) {
      const mark = (n.marks ?? []).find((m) => markName(m) === REDACT_MARK)
      const label =
        mark && 'attrs' in mark ? (mark.attrs as { label?: unknown } | undefined)?.label : undefined
      out.push(typeof label === 'string' && label.trim() !== '' ? label : 'private')
    }
    const children = n.content
    if (Array.isArray(children)) (children as ProseNode[]).forEach(visit)
  }
  visit(node)
  return out
}

/**
 * TipTap serialises a mark's `type` as its schema object, not the registered
 * name, so both forms have to be accepted — a strict `=== 'redaction'` silently
 * matches nothing and the span is read back in full.
 */
interface ProseMark {
  type?: string | { name?: string }
  attrs?: Record<string, unknown>
}
export interface ProseNode {
  type?: string | { name: string }
  text?: string
  marks?: ProseMark[]
  content?: ProseNode[]
}

function markName(mark: ProseMark | undefined): string | undefined {
  const type = mark?.type
  if (typeof type === 'string') return type
  return type?.name
}

function isWithheld(node: ProseNode): boolean {
  return Boolean(node.marks?.some((m) => markName(m) === REDACT_MARK))
}
function isLeafText(node: ProseNode): boolean {
  return typeof node.text === 'string'
}
function isBlock(node: ProseNode): boolean {
  return Array.isArray(node.content)
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

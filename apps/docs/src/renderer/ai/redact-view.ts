import {
  redactTextBetween as redactRangeTextBetween,
  type WithheldMark,
  type WithheldNode,
} from '@genoffice/agent-core/redact-range'
import { REDACT_MARK, placeholderSource } from './redact'

/**
 * The model-facing view of a document: every withheld span becomes its marker,
 * and the text it covers disappears. A withheld picture becomes a marker too,
 * without a path, so the model is not handed a reference it could fetch.
 *
 * The walk takes a plain shape rather than a ProseMirror node so it can be
 * tested without a schema, and so it does not depend on which ProseMirror
 * version the editor is pinned to. It is total: anything it does not recognise
 * passes through, because a throw here would be a silent leak — the caller's
 * fallback would be the unredacted text.
 */
/** the shared shape, not a second one */
export type RedactableNode = WithheldNode
export type RedactableMark = WithheldMark

/**
 * TipTap serialises a mark's `type` as its schema object rather than the
 * registered name, so both forms have to be accepted — a strict
 * `=== 'redaction'` matches nothing and the span is read back in full.
 */
function markName(mark: RedactableMark | undefined): string {
  const type = mark?.type
  if (typeof type === 'string') return type
  return type?.name ?? ''
}

function nameOf(node: RedactableNode): string {
  const type = node.type
  if (typeof type === 'string') return type
  return type?.name ?? ''
}

function redactionMarkOf(node: RedactableNode): RedactableMark | undefined {
  return node.marks?.find((m) => markName(m) === REDACT_MARK)
}

function isWithheld(node: RedactableNode): boolean {
  return redactionMarkOf(node) !== undefined
}

function labelOf(node: RedactableNode): string {
  const label = redactionMarkOf(node)?.attrs?.label
  return typeof label === 'string' ? label : 'private'
}

/** A picture that is withheld still needs a marker: it has no text to redact. */
function isPicture(node: RedactableNode): boolean {
  const name = nameOf(node)
  return name === 'docInlineImage' || name === 'image'
}

/** The text of a document as the model may read it. */
export function modelTextOf(node: RedactableNode): string {
  if (isWithheld(node)) return placeholderSource(labelOf(node))
  if (typeof node.text === 'string') return node.text
  const children = node.content
  if (!Array.isArray(children)) return ''
  let out = ''
  for (const child of children) {
    out += modelTextOf(child)
    // a block boundary is a line break, but not a trailing one: the caller
    // joins this with other text and a stray newline reads as a blank line
    if (Array.isArray(child.content) && child !== children[children.length - 1]) out += '\n'
  }
  return out
}

/**
 * The text of a range, as the model may read it.
 *
 * A thin wrapper over the shared walker so the options are stated once for
 * docs — the mark name, the marker, and the picture rule all live here rather
 * than at every call site.
 */
export function redactTextBetween(
  node: RedactableNode,
  from: number,
  to: number,
  blockSeparator = '\n',
): string {
  return redactRangeTextBetween(node, from, to, {
    markName: REDACT_MARK,
    marker: placeholderSource,
    blockSeparator,
    isLeaf: (n) => isPicture(n as RedactableNode),
  })
}

/** The same document with spans replaced, for callers that need the nodes. */
export function redactNode<T extends RedactableNode>(node: T): T {
  if (isWithheld(node)) {
    if (isPicture(node)) {
      // a picture's `src` and inlined `dataUrl` are exactly what must not
      // travel: either one would let the model fetch the withheld image
      return { type: 'docInlineImage', attrs: { redact: labelOf(node) } } as unknown as T
    }
    return { type: 'text', text: placeholderSource(labelOf(node)) } as unknown as T
  }
  if (Array.isArray(node.content)) {
    return { ...node, content: node.content.map((c) => redactNode(c)) }
  }
  return node
}

/** Every label withheld in a document, for the model's instruction. */
export function redactLabelsOf(node: RedactableNode): string[] {
  const out: string[] = []
  const visit = (n: RedactableNode) => {
    if (isWithheld(n)) out.push(labelOf(n))
    if (Array.isArray(n.content)) n.content.forEach(visit)
  }
  visit(node)
  return out
}

/** How many spans a document withholds; the guard's "is there anything to check". */
export function redactionCount(node: RedactableNode): number {
  let n = 0
  const visit = (x: RedactableNode) => {
    if (isWithheld(x)) n++
    if (Array.isArray(x.content)) x.content.forEach(visit)
  }
  visit(node)
  return n
}

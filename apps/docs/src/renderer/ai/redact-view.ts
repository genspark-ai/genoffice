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
export interface RedactableMark {
  type?: string | { name?: string }
  attrs?: Record<string, unknown>
}
export interface RedactableNode {
  type?: string | { name?: string }
  text?: string
  /** a picture's src and inlined dataUrl live here — what must not travel */
  attrs?: Record<string, unknown>
  marks?: RedactableMark[]
  content?: RedactableNode[]
}

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
 * A node's size in ProseMirror's counting, which is what a selection offset
 * means: one token for the opening, the content, one for the closing.
 */
function nodeSize(n: RedactableNode): number {
  if (typeof n.text === 'string') return n.text.length
  if (isPicture(n)) return 1
  if (!Array.isArray(n.content)) return 1
  return 2 + n.content.reduce((sum, c) => sum + nodeSize(c), 0)
}

/**
 * The text of a document *range*, as the model may read it.
 *
 * The whole-document case above is the easy one; a selection is where a leak
 * hides, because the context sent for a partial selection quotes the span back
 * — and the span is exactly what the reader withheld. Offsets are
 * ProseMirror's, so a caller passes its own `textBetween` range straight in.
 *
 * A withheld span overlapping the range contributes its marker once, not the
 * slice that fell inside: the words are gone either way, and a fragment of a
 * marker would be worse than none — it reads as a typo the model should fix.
 */
export function redactTextBetween(
  node: RedactableNode,
  from: number,
  to: number,
  blockSeparator = '\n',
): string {
  if (to <= from) return ''
  const blocks = node.content
  if (!Array.isArray(blocks)) return ''
  const out: string[] = []
  // a top-level block's position: 0, then the size of the one before it
  let blockPos = 0
  // A separator only belongs between two blocks the range actually spans: a
  // range that starts at a block boundary gets no leading one, and one that
  // stops inside a block gets no trailing one. `textBetween` does the same, so
  // the preview reads identically to the text it replaces.
  let contributed = false
  blocks.forEach((block, i) => {
    if (i > 0 && contributed && blockPos < to) out.push(blockSeparator)
    const before = out.length
    // a block's inline content starts one past the block itself
    if (blockPos + 1 < to) walkInline(block, blockPos + 1, from, to, out)
    contributed = out.length > before
    blockPos += nodeSize(block)
  })
  return out.join('')
}

/** the inline children of one block, whose content starts at `contentStart` */
function walkInline(
  block: RedactableNode,
  contentStart: number,
  from: number,
  to: number,
  out: string[],
): void {
  if (isWithheld(block)) {
    out.push(placeholderSource(labelOf(block)))
    return
  }
  const children = block.content
  if (!Array.isArray(children)) {
    if (typeof block.text === 'string' && contentStart < to) {
      const start = contentStart
      out.push(block.text.slice(Math.max(0, from - start), Math.min(block.text.length, to - start)))
    }
    return
  }
  // a text child at `pos` occupies [pos, pos + len): its own position is the
  // first character, not one before it
  let pos = contentStart
  for (const child of children) {
    const start = pos
    if (isWithheld(child)) {
      // the span, not the words
      if (start < to && start + nodeSize(child) > from) out.push(placeholderSource(labelOf(child)))
    } else if (isPicture(child)) {
      if (start < to) out.push(' ')
    } else if (typeof child.text === 'string') {
      if (start + child.text.length > from && start < to) {
        out.push(
          child.text.slice(Math.max(0, from - start), Math.min(child.text.length, to - start)),
        )
      }
    } else {
      walkInline(child, start, from, to, out)
    }
    pos = start + nodeSize(child)
  }
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

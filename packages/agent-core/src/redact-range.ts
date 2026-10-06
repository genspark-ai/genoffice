/**
 * The model-facing text of a document *range*.
 *
 * ## Why this lives here and not in an app
 *
 * Two editors learned this lesson the same way: a selection is quoted back to
 * the model alongside the blocks around it, so a reader who marked a phone
 * number and then selected across it got the number in the request. Each app
 * fixed its own copy, and a fifth copy would have been a fifth place to forget.
 *
 * It is a subpath export rather than part of the barrel because the barrel
 * reaches `electron`, and this runs in a renderer.
 *
 * ## The shape
 *
 * Deliberately plain: a ProseMirror node as JSON, no schema, no editor. That
 * way it is testable without a schema and does not depend on which version an
 * app has pinned. Offsets are ProseMirror's, so a caller passes its own
 * `textBetween` range straight in.
 *
 * Total on purpose. Anything it does not recognise passes through, because a
 * throw here would be a silent leak — the caller's fallback would be the
 * unredacted text.
 */

export interface WithheldMark {
  /** TipTap serialises a mark's type as its schema object rather than the name */
  type?: string | { name?: string }
  attrs?: Record<string, unknown>
}

export interface WithheldNode {
  type?: string | { name?: string }
  text?: string
  /** a picture's `src` and inlined `dataUrl` live here — what must not travel */
  attrs?: Record<string, unknown>
  marks?: readonly WithheldMark[]
  content?: readonly WithheldNode[]
}

export interface RedactRangeOptions {
  /** the mark name that means "withheld from the model" */
  markName: string
  /** how a label becomes the text the model does see */
  marker: (label: string) => string
  /** what stands between two blocks */
  blockSeparator?: string
  /** what stands in for a node that has no text of its own */
  leafText?: string
  /** a node whose content must not travel: a picture's src, an atom's payload */
  isLeaf?: (node: WithheldNode) => boolean
}

const DEFAULT_BLOCK_SEPARATOR = '\n'
const DEFAULT_LEAF_TEXT = ' '

function nameOf(value: string | { name?: string } | undefined): string | undefined {
  return typeof value === 'string' ? value : value?.name
}

function markNameOf(mark: WithheldMark, wanted: string): boolean {
  return nameOf(mark.type) === wanted
}

function labelOf(node: WithheldNode, wanted: string): string {
  const mark = (node.marks ?? []).find((m: WithheldMark) => markNameOf(m, wanted))
  const label = mark?.attrs?.label
  return typeof label === 'string' ? label : 'private'
}

/** A node's size in ProseMirror's counting: one token each side, content between. */
export function withheldNodeSize(node: WithheldNode, isLeaf: (n: WithheldNode) => boolean): number {
  if (typeof node.text === 'string') return node.text.length
  if (isLeaf(node)) return 1
  if (!Array.isArray(node.content)) return 1
  return 2 + node.content.reduce((sum, child) => sum + withheldNodeSize(child, isLeaf), 0)
}

/**
 * The text of `from..to` as the model may read it.
 *
 * A withheld span overlapping the range contributes its whole marker rather
 * than the slice inside it: the words are gone either way, and a fragment of a
 * marker reads as a typo the model would try to repair.
 *
 * A separator belongs only between two blocks the range actually spans — a
 * range starting at a block boundary gets no leading one and one stopping
 * inside a block gets no trailing one, which is what `textBetween` does.
 */
export function redactTextBetween(
  node: WithheldNode,
  from: number,
  to: number,
  options: RedactRangeOptions,
): string {
  const {
    markName,
    marker,
    blockSeparator = DEFAULT_BLOCK_SEPARATOR,
    leafText = DEFAULT_LEAF_TEXT,
    isLeaf = () => false,
  } = options
  if (to <= from) return ''
  const blocks = node.content
  if (!Array.isArray(blocks)) return ''
  // resolved once, so the walk below cannot read an option the caller omitted
  const resolved: Required<RedactRangeOptions> = {
    markName,
    marker,
    blockSeparator,
    leafText,
    isLeaf,
  }
  const size = (n: WithheldNode) => withheldNodeSize(n, resolved.isLeaf)
  const out: string[] = []
  let contributed = false
  // a top-level block's position: 0, then the size of the one before it
  let blockPos = 0
  blocks.forEach((block, i) => {
    if (i > 0 && contributed && blockPos < to) out.push(blockSeparator)
    const before = out.length
    // a block's inline content starts one past the block itself
    if (blockPos + 1 < to) walkInline(block, blockPos + 1, from, to, out, resolved, size)
    contributed = out.length > before
    blockPos += size(block)
  })
  return out.join('')
}

/** the inline children of one block, whose content starts at `contentStart` */
function walkInline(
  block: WithheldNode,
  contentStart: number,
  from: number,
  to: number,
  out: string[],
  options: Required<RedactRangeOptions>,
  size: (n: WithheldNode) => number,
): void {
  const { markName, marker, leafText, isLeaf } = options
  if ((block.marks ?? []).some((m: WithheldMark) => markNameOf(m, markName))) {
    out.push(marker(labelOf(block, markName)))
    return
  }
  const children = block.content
  if (!Array.isArray(children)) {
    if (typeof block.text === 'string' && contentStart < to) {
      out.push(
        block.text.slice(
          Math.max(0, from - contentStart),
          Math.min(block.text.length, to - contentStart),
        ),
      )
    }
    return
  }
  // a text child at `pos` occupies [pos, pos + len): its own position is the
  // first character, not one before it
  let pos = contentStart
  for (const child of children) {
    const start = pos
    if ((child.marks ?? []).some((m: WithheldMark) => markNameOf(m, markName))) {
      if (start < to && start + size(child) > from) out.push(marker(labelOf(child, markName)))
    } else if (isLeaf(child)) {
      if (start < to) out.push(leafText)
    } else if (typeof child.text === 'string') {
      if (start + child.text.length > from && start < to) {
        out.push(
          child.text.slice(Math.max(0, from - start), Math.min(child.text.length, to - start)),
        )
      }
    } else {
      walkInline(child, start, from, to, out, options, size)
    }
    pos = start + size(child)
  }
}

/**
 * The namespace the "withheld from the model" label lives in, and the prefix
 * bound to it.
 *
 * Both file carriers write into it — a `go:redact` element in a docx run's rPr
 * and the same element inside a pptx `<a:ext>` — and an undeclared prefix is a
 * parse error rather than a warning: Word and PowerPoint both offer to repair
 * the file, and repair drops the run carrying the mark, so the span quietly
 * becomes readable by a model again on the next open.
 *
 * It lives here rather than in either engine because the two of them need it,
 * and a per-app stack has to build with only the core and one app present.
 */
export const REDACT_NS = 'https://genspark.ai/genoffice/redaction/2026'
export const REDACT_PREFIX = 'go'
export const REDACT_EL = `${REDACT_PREFIX}:redact`

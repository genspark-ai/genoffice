import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { TABLE_TRAILING_SKIP } from '../editor/extensions'
import { TRACK_IGNORE } from '../editor/revisions'
import { blockRangePositions, HTML_RULES, parseHtmlFragment, type NumIds } from './protocol'

/**
 * Long-form writing runs as its own tool-less request whose reply body IS the
 * restricted-HTML fragment. Text deltas stream straight into the document as
 * a draft, so the user watches the text arrive and no gateway sees a silent
 * connection (tool arguments are buffered server-side until the JSON is
 * complete; a document-sized argument exceeds gateway idle cutoffs).
 */
export const DOC_MAX_CHARS = 200_000
const CONTEXT_CAP = 8000
const DRAFT_TICK_MS = 300

export interface DocWriteSpec {
  plan: string
  title?: string | undefined
  context?: string | undefined
  /** the user's own request, verbatim */
  instruction: string
}

export interface DocWriteResult {
  ok: boolean
  /** the fragment to land (complete, or the adopted partial) */
  html?: string
  /** the stream ended early: transport drop, stop, or output cap */
  truncated?: boolean
  error?: string
}

/** Panel-owned writer: streams the fragment, reports progress, settles partial output with the user. */
export interface AiDocWriter {
  write(
    spec: DocWriteSpec,
    onProgress: (html: string) => void,
    signal?: AbortSignal,
  ): Promise<DocWriteResult>
}

const WRITER_SYSTEM = [
  'You are the document writer of GenOffice Docs, a Word-style editor. You receive a plan (title, sections, key points, tone, length) and reference material, and you write the complete content.',
  '',
  '## Writing rules',
  '- Write the requested content in full: every section in the plan, real prose, no placeholders, no "[to be added]" notes.',
  '- Real data only: use the facts, figures, names and quotes from the reference material and the plan; never invent facts or numbers beyond them.',
  '- Follow the length the plan asks for; when it gives none, write what the content genuinely needs.',
  '',
  '## Output',
  'Reply with the restricted HTML fragment only: it starts with the first block tag and ends with the last closing tag. No markdown code fences, no commentary before or after, no <html>/<body> wrapper.',
  HTML_RULES,
].join('\n')

export function buildDocWriterRequest(
  spec: DocWriteSpec,
  langDirective: string,
): { system: string; user: string } {
  const context = spec.context?.trim()
    ? `\n\n## Reference material (all names, figures and facts come from here)\n${spec.context.slice(0, CONTEXT_CAP)}`
    : ''
  const instruction = spec.instruction.trim()
    ? `\n\n## The user's request\n${spec.instruction.slice(0, 2000)}`
    : ''
  return {
    system: WRITER_SYSTEM + langDirective,
    user: `## Plan\n${spec.title ? `Title: ${spec.title}\n` : ''}${spec.plan.slice(0, 6000)}${context}${instruction}\n\nWrite the content now.`,
  }
}

/** Strip a fence or chatter before the first block tag and a closing fence at the end. */
export function extractFragment(raw: string): string {
  let text = raw.replace(/\r\n/g, '\n')
  const start = text.search(/<(?:h[1-6]|p|ul|ol|table|pre|blockquote|formula)[\s>]/i)
  if (start > 0) text = text.slice(start)
  else if (start < 0) text = text.replace(/^\s*```[a-z]*\n?/i, '')
  return text.replace(/\n?```\s*$/, '').trimEnd()
}

/** rough top-level block count of a fragment, for the progress chip */
export function countFragmentBlocks(html: string): number {
  return (html.match(/<(?:h[1-6]|p|ul|ol|table|pre|blockquote|formula)[\s>]/gi) ?? []).length
}

export type WritePosition = { kind: 'whole' } | { kind: 'after'; index: number }

/** meta on every draft transaction: not an edit, not history, not a tracked change */
export const DRAFT_META = 'aiDraft'

/**
 * Live draft of a streaming fragment inside the document. Each tick re-parses
 * the fragment received so far (the HTML parser closes open tags, so a
 * half-received block still renders) and swaps only the blocks that changed.
 * The draft is outside undo history and Track Changes; `finish()` removes it
 * so the caller can land the final fragment through the regular insert path,
 * giving one undo step and normal revision marks.
 */
export class DraftLanding {
  private from: number
  private to: number
  /** the text of the last tick that parsed and rendered */
  private rendered: string | null = null
  private pending: string | null = null
  private timer: ReturnType<typeof setTimeout> | null = null
  private closed = false
  private follow = true
  private readonly onTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (!transaction.docChanged || transaction.getMeta(DRAFT_META)) return
    this.from = transaction.mapping.map(this.from, 1)
    this.to = Math.max(this.from, transaction.mapping.map(this.to, -1))
  }

  constructor(
    private readonly editor: Editor,
    private readonly numIds: NumIds,
    position: WritePosition,
  ) {
    this.from =
      position.kind === 'after' && position.index >= 0
        ? blockRangePositions(editor, position.index, position.index).to
        : 0
    this.to = this.from
    editor.on('transaction', this.onTransaction)
  }

  /** latest cumulative fragment; rendered on the next tick */
  update(html: string): void {
    if (this.closed) return
    this.pending = html
    if (this.timer === null) this.timer = setTimeout(() => this.flush(), DRAFT_TICK_MS)
  }

  /** the top-level block index the draft sits after (-1 = document start), after any user edits */
  indexBefore(): number {
    const doc = this.editor.state.doc
    let index = -1
    let i = 0
    doc.forEach((node, offset) => {
      if (offset + node.nodeSize <= this.from) index = i
      i++
    })
    return index
  }

  /**
   * Remove the draft; the document is back to what it was around it. Returns the
   * text of the last tick that rendered, so a kept stream whose tail no longer
   * parses (a half-received formula or table) can still land what the user saw.
   */
  finish(): string | null {
    if (this.closed) return this.rendered
    if (this.timer !== null) clearTimeout(this.timer)
    this.timer = null
    this.flush()
    this.closed = true
    this.editor.off('transaction', this.onTransaction)
    if (this.to > this.from && !this.editor.isDestroyed) {
      this.editor.view.dispatch(this.draftTr().delete(this.from, this.to))
    }
    this.to = this.from
    return this.rendered
  }

  /** live top-level blocks inside the draft range, after any user edits */
  private liveNodes(): Array<{ node: ProseMirrorNode; end: number }> {
    const out: Array<{ node: ProseMirrorNode; end: number }> = []
    this.editor.state.doc.forEach((node, offset) => {
      if (offset >= this.from && offset + node.nodeSize <= this.to)
        out.push({ node, end: offset + node.nodeSize })
    })
    return out
  }

  private draftTr(): Transaction {
    return this.editor.state.tr
      .setMeta(DRAFT_META, true)
      .setMeta('addToHistory', false)
      .setMeta(TRACK_IGNORE, true)
      .setMeta(TABLE_TRAILING_SKIP, true)
  }

  private flush(): void {
    this.timer = null
    if (this.closed || this.pending === null || this.editor.isDestroyed) return
    const html = this.pending
    this.pending = null
    let next: ProseMirrorNode[]
    try {
      next = parseHtmlFragment(html, this.numIds).map((n) =>
        this.editor.schema.nodeFromJSON({ ...n, attrs: { ...n.attrs, aiChanged: false } }),
      )
    } catch {
      return // a half-received formula or table is retried on the next tick
    }
    this.rendered = html
    // compare against the live document: a user edit inside the draft changes sizes and drops that block from the kept prefix
    const live = this.liveNodes()
    let keep = 0
    while (keep < live.length && keep < next.length && live[keep]!.node.eq(next[keep]!)) keep++
    if (keep === live.length && keep === next.length) return
    const at = keep > 0 ? live[keep - 1]!.end : this.from
    const wasVisible = this.follow && this.endVisible()
    this.editor.view.dispatch(this.draftTr().replaceWith(at, this.to, next.slice(keep)))
    this.to = at + next.slice(keep).reduce((size, n) => size + n.nodeSize, 0)
    if (wasVisible) this.scrollToEnd()
    else this.follow = false
  }

  private endVisible(): boolean {
    if (this.to === this.from) return true
    try {
      const { top } = this.editor.view.coordsAtPos(this.to - 1)
      return top >= 0 && top <= window.innerHeight
    } catch {
      return true
    }
  }

  private scrollToEnd(): void {
    if (this.to === this.from) return
    try {
      const { node } = this.editor.view.domAtPos(this.to - 1)
      const el = node instanceof Element ? node : node.parentElement
      el?.scrollIntoView?.({ block: 'nearest' })
    } catch {
      /* layout not ready */
    }
  }
}

import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { ResolvedPos } from '@tiptap/pm/model'
import { wordUnitAt } from '../editor/word-range'

/** Format-painter (Home > Format Painter) state machine, extracted from Ribbon.tsx. */
export interface PainterFormat {
  marks: Array<{ type: string; attrs: Record<string, unknown> }>
  /** source paragraph's node type + formatting attrs (null when the caret is not in a paintable block) */
  block: { type: string; attrs: Record<string, unknown> } | null
}

export interface PainterState extends PainterFormat {
  /** double-click on the button: stays armed until Esc or another click */
  locked: boolean
}

/** Character-formatting marks the painter transfers; semantic marks (links,
 *  comments, revisions, fields) are neither picked up nor stripped from the target. */
export const PAINTER_MARK_TYPES = ['bold', 'italic', 'underline', 'strike', 'docTextStyle']

/** Paragraph-formatting attrs the painter transfers. Identity/anchor attrs
 *  (docxIndex, bookmarks, comment ranges, revisions, sdtShell…) stay with the target. */
export const PAINTER_PARA_KEYS = [
  'styleId',
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'snapToGrid',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'bidi',
  'autoSpace',
  'shadingFill',
  'emptyRunSize',
  'borders',
  'borderLines',
  'tabStops',
]

/** Per-block-type attrs that define the block's identity as formatting (heading level, list numbering) */
export const PAINTER_BLOCK_EXTRA: Record<string, string[]> = {
  docParagraph: [],
  docHeading: ['level'],
  docListItem: ['kind', 'numId', 'ilvl'],
}

/** the word under a painter click: Word brushes the word and its trailing spaces */
export function painterWordRangeAt($pos: ResolvedPos): { from: number; to: number } | null {
  const para = $pos.parent
  if (!para.isTextblock) return null
  const text = para.textBetween(0, para.content.size, '\0', '\0')
  const base = $pos.start()
  const span = wordUnitAt(text, $pos.parentOffset)
  // no word under the pointer: nothing to mark, the paragraph format still applies
  if (!span) return { from: base + $pos.parentOffset, to: base + $pos.parentOffset }
  return { from: base + span.start, to: base + span.end }
}

export function applyPainterFormat(
  editor: Editor,
  fmt: PainterFormat,
  from: number,
  to: number,
  caretAfter: number | null,
): void {
  // Word: every brush stroke is its own undo step, even quick strokes on
  // neighbouring paragraphs that history would otherwise merge into one
  let c = editor
    .chain()
    .focus()
    .command(({ tr }) => {
      closeHistory(tr)
      return true
    })
    .setTextSelection({ from, to })
  if (to > from) {
    // strip only formatting marks, then re-add the picked-up ones: semantic
    // marks on the target (links, comments, revisions) survive the brush
    for (const t of PAINTER_MARK_TYPES) c = c.unsetMark(t)
    for (const m of fmt.marks) c = c.setMark(m.type, m.attrs)
  }
  c = c.command(({ tr }) => {
    const block = fmt.block
    if (!block) return true
    const type = editor.schema.nodes[block.type]
    if (!type) return true
    const sel = tr.selection
    const jobs: Array<{ pos: number; attrs: Record<string, unknown> }> = []
    tr.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
      if (!(node.type.name in PAINTER_BLOCK_EXTRA)) return true
      // keep the target's identity attrs, overwrite every formatting attr
      // (explicit nulls in block.attrs reset what the source didn't set)
      jobs.push({ pos, attrs: { ...node.attrs, ...block.attrs } })
      return false
    })
    for (const job of jobs) tr.setNodeMarkup(job.pos, type, job.attrs)
    return true
  })
  if (caretAfter != null) c = c.setTextSelection(caretAfter)
  c.run()
}

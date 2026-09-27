import type { Editor } from '@tiptap/core'
import type { Node as PMNode, ResolvedPos } from '@tiptap/pm/model'
import { AllSelection, TextSelection, type Selection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { CellSelection } from '@tiptap/pm/tables'
import { getActiveSubEditor } from './active-editor'

/**
 * Word's selection bar: the left page margin beside the text column. A click
 * selects the visual line next to the pointer, a drag extends line by line,
 * double-click takes the paragraph, triple-click the document, Shift+click
 * extends the current selection to that line.
 */

export const SELECTION_BAR_HOVER_CLASS = 'doc-selbar-hover'

interface Range {
  from: number
  to: number
}

/** page geometry in client px: the root padding is the page margin (scaled by the CSS zoom in effect) */
interface Band {
  left: number
  textLeft: number
  textRight: number
  top: number
  bottom: number
}

type Unit = 'line' | 'paragraph'

function bandOf(dom: HTMLElement): Band | null {
  if (!dom.offsetWidth) return null
  const rect = dom.getBoundingClientRect()
  const f = rect.width / dom.offsetWidth
  const cs = getComputedStyle(dom)
  return {
    left: rect.left,
    textLeft: rect.left + parseFloat(cs.paddingLeft) * f,
    textRight: rect.right - parseFloat(cs.paddingRight) * f,
    top: rect.top + parseFloat(cs.paddingTop) * f,
    bottom: rect.bottom - parseFloat(cs.paddingBottom) * f,
  }
}

/**
 * The band of the row at client y. Blocks of translated pages (mirrored even
 * pages, mixed-column layouts) carry the shift as --col-dx; the text column
 * beside the pointer moves with the block on that row.
 */
function bandAt(view: EditorView, band: Band, y: number): Band {
  const pos = view.posAtCoords({ left: band.textLeft + 1, top: y })?.pos
  if (pos == null) return band
  const $pos = view.state.doc.resolve(pos)
  if ($pos.depth < 1) return band
  const el = view.nodeDOM($pos.before(1))
  const dx = el instanceof HTMLElement ? parseFloat(el.style.getPropertyValue('--col-dx')) : 0
  if (!dx) return band
  const dom = view.dom as HTMLElement
  const f = dom.offsetWidth ? dom.getBoundingClientRect().width / dom.offsetWidth : 1
  return { ...band, textLeft: band.textLeft + dx * f, textRight: band.textRight + dx * f }
}

/** the pointer is on the root's own left padding beside the text column (page gaps, strips and floats are children and miss this) */
function onBar(dom: HTMLElement, band: Band, e: MouseEvent): boolean {
  return (
    e.target === dom &&
    e.clientX >= band.left &&
    e.clientX < band.textLeft &&
    e.clientY >= band.top &&
    e.clientY <= band.bottom
  )
}

/**
 * Document position at the text-column edge on the pointer's row. A probe that
 * lands outside the nearest block's box (indent, paragraph spacing) resolves to
 * the gap before/after that block; re-probe inside its box to reach the line.
 */
function edgePos(view: EditorView, x: number, y: number, side: 'start' | 'end'): number | null {
  const probe = (px: number, py: number) => view.posAtCoords({ left: px, top: py })?.pos ?? null
  const pos = probe(x, y)
  if (pos == null) return null
  const $pos = view.state.doc.resolve(pos)
  if ($pos.parent.inlineContent) return pos
  const node = side === 'start' ? $pos.nodeAfter : $pos.nodeBefore
  if (!node) return pos
  const el = view.nodeDOM(side === 'start' ? pos : pos - node.nodeSize)
  if (!(el instanceof HTMLElement)) return pos
  const r = el.getBoundingClientRect()
  const cx = Math.min(Math.max(x, r.left + 1), r.right - 1)
  const cy = Math.min(Math.max(y, r.top + 1), r.bottom - 1)
  return probe(cx, cy) ?? pos
}

function rowAround($pos: ResolvedPos): { node: PMNode; pos: number } | null {
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d)
    if (node.type.spec.tableRole === 'row') return { node, pos: $pos.before(d) }
  }
  return null
}

function lineRangeAt(view: EditorView, band: Band, y: number): Range | null {
  const a = edgePos(view, band.textLeft + 1, y, 'start')
  if (a == null) return null
  const $a = view.state.doc.resolve(a)
  // the line ends at this block's own right edge (columns, side-by-side cells)
  let right = band.textRight - 1
  const blockEl = $a.depth > 0 ? view.nodeDOM($a.before()) : null
  if (blockEl instanceof HTMLElement)
    right = Math.min(right, blockEl.getBoundingClientRect().right - 1)
  const b = edgePos(view, Math.max(right, band.textLeft + 1), y, 'end') ?? a
  return { from: Math.min(a, b), to: Math.max(a, b) }
}

function paragraphRange(view: EditorView, line: Range): Range {
  const $from = view.state.doc.resolve(line.from)
  if (!$from.parent.isTextblock) return line
  return { from: $from.start(), to: $from.end() }
}

function unitRangeAt(view: EditorView, band: Band, y: number, unit: Unit): Range | null {
  const line = lineRangeAt(view, band, y)
  if (!line) return null
  return unit === 'paragraph' ? paragraphRange(view, line) : line
}

/** a line beside a table row selects the row (Word) */
function rowSelectionFor(view: EditorView, range: Range): Selection | null {
  const { doc } = view.state
  const row = rowAround(doc.resolve(range.from))
  if (!row || row.node.childCount === 0) return null
  const first = row.pos + 1
  let last = first
  row.node.forEach((cell, offset, i) => {
    if (i === row.node.childCount - 1) last = first + offset
  })
  return CellSelection.create(doc, first, last)
}

function between(view: EditorView, anchor: number, head: number): Selection {
  const { doc } = view.state
  return TextSelection.between(doc.resolve(anchor), doc.resolve(head))
}

function select(view: EditorView, sel: Selection): void {
  if (!sel.eq(view.state.selection)) view.dispatch(view.state.tr.setSelection(sel))
  view.focus()
}

export function installSelectionBar(editor: Editor): () => void {
  const view = editor.view
  const dom = view.dom as HTMLElement
  let drag: { band: Band; unit: Unit; anchor: Range } | null = null

  const enabled = () => !getActiveSubEditor() && !dom.closest('.page-wrap.hf-editing')

  const onMove = (e: MouseEvent) => {
    if (drag) return
    const root = e.target === dom && enabled() ? bandOf(dom) : null
    const band = root && bandAt(view, root, e.clientY)
    dom.classList.toggle(SELECTION_BAR_HOVER_CLASS, !!band && onBar(dom, band, e))
  }
  const onLeave = () => dom.classList.remove(SELECTION_BAR_HOVER_CLASS)

  const onDown = (e: MouseEvent) => {
    if (e.button !== 0 || !enabled()) return
    const root = bandOf(dom)
    const band = root && bandAt(view, root, e.clientY)
    if (!band || !onBar(dom, band, e)) return
    // keep the native caret placement and drag-selection out of the way
    e.preventDefault()
    e.stopImmediatePropagation()
    const clicks = e.detail || 1
    if (clicks >= 3) {
      select(view, new AllSelection(view.state.doc))
      return
    }
    const unit: Unit = clicks === 2 ? 'paragraph' : 'line'
    const range = unitRangeAt(view, band, e.clientY, unit)
    if (!range) {
      view.focus()
      return
    }
    if (e.shiftKey && clicks === 1) {
      const anchor = view.state.selection.anchor
      select(view, between(view, anchor, range.from >= anchor ? range.to : range.from))
      return
    }
    select(
      view,
      (unit === 'line' && rowSelectionFor(view, range)) || between(view, range.from, range.to),
    )
    drag = { band, unit, anchor: range }
    window.addEventListener('mousemove', onDragMove, true)
    window.addEventListener('mouseup', onDragEnd, true)
  }

  const onDragMove = (e: MouseEvent) => {
    if (!drag) return
    const root = bandOf(dom) ?? drag.band
    const y = Math.min(Math.max(e.clientY, root.top), root.bottom)
    const cur = unitRangeAt(view, bandAt(view, root, y), y, drag.unit)
    if (!cur) return
    const { anchor } = drag
    const sel =
      cur.from >= anchor.from
        ? between(view, anchor.from, Math.max(cur.to, anchor.to))
        : between(view, anchor.to, cur.from)
    if (!sel.eq(view.state.selection)) view.dispatch(view.state.tr.setSelection(sel))
  }
  const onDragEnd = () => {
    drag = null
    window.removeEventListener('mousemove', onDragMove, true)
    window.removeEventListener('mouseup', onDragEnd, true)
  }

  dom.addEventListener('mousedown', onDown, true)
  dom.addEventListener('mousemove', onMove)
  dom.addEventListener('mouseleave', onLeave)
  return () => {
    onDragEnd()
    onLeave()
    dom.removeEventListener('mousedown', onDown, true)
    dom.removeEventListener('mousemove', onMove)
    dom.removeEventListener('mouseleave', onLeave)
  }
}

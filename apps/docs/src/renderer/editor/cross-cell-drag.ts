/**
 * Word-style continuous text selection across table cells (genspark-ai/genoffice#2329).
 *
 * prosemirror-tables hijacks every drag that starts inside a cell: its
 * document-level mousemove listener dispatches a `CellSelection` the moment
 * the pointer enters another cell, and even a drag aimed *out* of the table
 * collapses into a one-cell selection. The text selection therefore "stops at
 * the cell boundary" and jumps to whole cells. Word instead extends the text
 * selection continuously — partial content of each crossed cell, structure
 * untouched.
 *
 * Chromium extends a native text selection across cell boundaries on its own;
 * the only thing in the way is that hijack. This extension tracks drags that
 * start inside a cell's text and, once the pointer leaves that cell,
 * `stopImmediatePropagation()`s the document-level mousemove in the capture
 * phase, before it reaches prosemirror-tables' bubble listener, with a
 * `posAtCoords` dispatch as the clamp-proof fallback. Drags that stay inside
 * the anchor cell pass through untouched, as do shift-clicks (prosemirror-
 * tables' extend-to-cell flow) and ctrl/⌘ clicks (multi-select).
 *
 * A cross-cell selection then behaves like Word's: deleting or typing over it
 * clears the covered text per cell and keeps every cell it crossed, and an
 * IME composition over it starts on a collapsed caret.
 */
import { Extension } from '@tiptap/core'
import type { Node as PmNode, ResolvedPos } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { columnResizingPluginKey } from '@tiptap/pm/tables'

/** the nearest table cell (or header cell) a resolved position sits in */
function cellAncestor($pos: ResolvedPos): PmNode | null {
  for (let depth = $pos.depth; depth > 0; depth--) {
    const node = $pos.node(depth)
    const role = node.type.spec.tableRole
    if (role === 'cell' || role === 'header_cell') return node
  }
  return null
}

/**
 * Whether the `from..to` range behaves like one Word selection crossing cell
 * boundaries: its ends sit in different cells, or one end sits in a cell and
 * the other outside any table. Both get the structure-preserving deletion.
 */
export function selectionCrossesCellBoundary(doc: PmNode, from: number, to: number): boolean {
  if (from === to) return false
  const a = cellAncestor(doc.resolve(from))
  const b = cellAncestor(doc.resolve(to))
  // null === null: both ends outside any table, nothing to preserve
  return (a === null) !== (b === null) || (a !== null && b !== null && a !== b)
}

/**
 * The text spans `from..to` covers, one per textblock, in document order.
 * Deletion walks these instead of replacing the whole range, so the cells
 * between the ends keep their paragraphs (emptied where fully covered) and
 * the table structure survives — Word's own behaviour.
 */
export function cellDeletionSpans(
  doc: PmNode,
  from: number,
  to: number,
): { from: number; to: number }[] {
  const spans: { from: number; to: number }[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isTextblock) {
      const lo = Math.max(from, pos + 1)
      const hi = Math.min(to, pos + node.nodeSize - 1)
      if (hi > lo) spans.push({ from: lo, to: hi })
      return false
    }
    return true
  })
  return spans
}

/// One transaction that removes every text span `from..to` covers and inserts
/// `text` where the selection began (empty text = plain deletion). Deleting in
/// reverse keeps the earlier spans' positions valid; the first span starts at
/// the smallest position and survives them all. The caret parks at the end of
/// the inserted text.
function clearSpansTr(state: EditorState, from: number, to: number, text: string): Transaction {
  const tr = state.tr
  const spans = cellDeletionSpans(state.doc, from, to)
  for (const span of [...spans].reverse()) tr.delete(span.from, span.to)
  const caret = spans.length ? spans[0].from : from
  if (text) tr.insertText(text, caret)
  tr.setSelection(TextSelection.create(tr.doc, caret + text.length))
  tr.scrollIntoView()
  return tr
}

/**
 * Delete the covered text of a cross-cell selection cell by cell, keep the
 * structure, and park the caret where the selection began. Returns false for
 * selections that are not cross-cell, leaving them to the default handling.
 */
export function deleteCrossCellSelection(view: EditorView): boolean {
  const { state } = view
  if (!(state.selection instanceof TextSelection) || state.selection.empty) return false
  const { from, to } = state.selection
  if (!selectionCrossesCellBoundary(state.doc, from, to)) return false
  view.dispatch(clearSpansTr(state, from, to, ''))
  return true
}

const escapeHtml = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

const dragKey = new PluginKey('crossCellDrag')

export const CrossCellDrag = Extension.create({
  name: 'crossCellDrag',

  addProseMirrorPlugins() {
    /// the drag this extension owns, armed by a plain mousedown inside a
    /// cell; null whenever no drag is in flight
    let drag: {
      anchorCell: Element
      anchorPos: number
      onMove: (event: MouseEvent) => void
      onUp: () => void
      root: Document
    } | null = null

    const cleanup = (): void => {
      if (!drag) return
      drag.root.removeEventListener('mousemove', drag.onMove, true)
      drag.root.removeEventListener('mouseup', drag.onUp, true)
      drag = null
    }

    const cellFromTarget = (view: EditorView, target: EventTarget | null): Element | null => {
      if (!(target instanceof Element)) return null
      if (!view.dom.contains(target)) return null
      return target.closest('td,th')
    }

    return [
      new Plugin({
        key: dragKey,
        props: {
          handleDOMEvents: {
            mousedown: (view, event) => {
              const mouse = event as MouseEvent
              // ctrl/⌘ is multi-select, shift-click is prosemirror-tables'
              // extend-to-cell — both keep their existing flows
              if (mouse.button !== 0 || mouse.ctrlKey || mouse.metaKey || mouse.shiftKey) {
                return false
              }
              const anchorCell = cellFromTarget(view, mouse.target)
              if (!anchorCell) return false
              // the column-resize handle owns drags starting on the cell edge
              const resize = columnResizingPluginKey.getState(view.state) as
                { activeHandle?: number } | undefined
              if (resize && (resize.activeHandle ?? -1) !== -1) return false
              const coords = view.posAtCoords({ left: mouse.clientX, top: mouse.clientY })
              if (!coords) return false
              cleanup()
              const anchorPos = coords.pos
              const root = view.dom.ownerDocument ?? document
              const onMove = (move: MouseEvent): void => {
                if (!(move.buttons & 1)) return cleanup()
                // pass the event through while the pointer stays in the cell
                // the drag started in; once it crosses into another cell (or
                // the gutter, or outside the table) stop it before it reaches
                // prosemirror-tables' bubble listener, which would otherwise
                // swap the text selection for a cell-selection rectangle
                if (cellFromTarget(view, move.target) === anchorCell) return
                move.stopImmediatePropagation()
                // keep the selection crossing even where the native drag would
                // clamp: the head follows the pointer, the anchor stays put
                const coords = view.posAtCoords({ left: move.clientX, top: move.clientY })
                if (!coords || coords.pos === anchorPos) return
                if (anchorPos > view.state.doc.content.size) return
                const { anchor, head } =
                  coords.pos >= anchorPos
                    ? { anchor: anchorPos, head: coords.pos }
                    : { anchor: coords.pos, head: anchorPos }
                view.dispatch(
                  view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)),
                )
              }
              const onUp = (): void => cleanup()
              drag = { anchorCell, anchorPos, onMove, onUp, root }
              root.addEventListener('mousemove', onMove, true)
              root.addEventListener('mouseup', onUp, true)
              return false
            },
            compositionstart: (view) => {
              // an IME replacing a cross-cell selection reaches PM as one
              // structural DOM replace; clear the covered text per cell first
              // so the composition starts on a collapsed caret like any other
              deleteCrossCellSelection(view)
              return false
            },
            // the default paste replaces the whole range as one slice, which
            // takes the cells between the ends with it; collapse per cell
            // first and let the paste land at the caret. Only when there is
            // something to paste: an empty clipboard must not clear the cells.
            paste: (view, event) => {
              if (!view.editable) return false
              const data = event.clipboardData
              if (!data) return false
              const hasContent =
                data.files.length > 0 || !!(data.getData('text/html') || data.getData('text/plain'))
              if (!hasContent) return false
              deleteCrossCellSelection(view)
              return false
            },
            // DOM handlers run before every plugin's handleKeyDown, so the
            // keymap's Enter splits a collapsed caret, not a range across cells
            keydown: (view, event) => {
              if (event.key === 'Enter' && !event.isComposing) deleteCrossCellSelection(view)
              return false
            },
            // Word puts the covered text on the clipboard, one paragraph per
            // cell span, not the rows and cells the slice would carry; a paste
            // of it then restores text, not a nested table
            cut: (view, event) => {
              const { state } = view
              if (!view.editable) return false
              if (!(state.selection instanceof TextSelection) || state.selection.empty) return false
              const { from, to } = state.selection
              if (!selectionCrossesCellBoundary(state.doc, from, to)) return false
              const data = event.clipboardData
              if (!data) return false
              // a hard break counts as a line break, so Shift+Enter inside a
              // spanned cell survives the trip through the clipboard
              const lines = cellDeletionSpans(state.doc, from, to).map((span) =>
                state.doc.textBetween(span.from, span.to, '\n', (leaf) =>
                  leaf.type.name === 'hardBreak' ? '\n' : '',
                ),
              )
              event.preventDefault()
              data.clearData()
              data.setData('text/plain', lines.join('\n'))
              data.setData(
                'text/html',
                lines.map((line) => `<p>${escapeHtml(line).split('\n').join('<br>')}</p>`).join(''),
              )
              view.dispatch(clearSpansTr(state, from, to, '').setMeta('uiEvent', 'cut'))
              return true
            },
          },
          handleKeyDown: (view, event) => {
            if (event.key !== 'Backspace' && event.key !== 'Delete') return false
            return deleteCrossCellSelection(view)
          },
          handleTextInput: (view, from, to, text) => {
            const { state } = view
            if (from === to || !selectionCrossesCellBoundary(state.doc, from, to)) return false
            view.dispatch(clearSpansTr(state, from, to, text))
            return true
          },
        },
      }),
    ]
  },
})

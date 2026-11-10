import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { columnResizingPluginKey } from '@tiptap/pm/tables'
import type { EditorView } from '@tiptap/pm/view'

export interface RowHeightPreview {
  pos: number
  heightPx: number
}

export const tableRowDragKey = new PluginKey<RowHeightPreview[] | null>('docTableRowDrag')

/**
 * Paint rows at the given heights while a drag is live (null ends the preview).
 * A node decoration, not an inline style: a style written straight onto the
 * <tr> is a DOM mutation ProseMirror redraws away.
 */
export function previewRowHeights(view: EditorView, rows: RowHeightPreview[] | null): void {
  view.dispatch(view.state.tr.setMeta(tableRowDragKey, rows))
}

/** pointer distance from a row's bottom border that arms the drag (screen px) */
const EDGE_PX = 4
const MIN_ROW_PX = 4
const CURSOR_CLASS = 'row-resize-cursor'

interface RowTarget {
  tr: HTMLTableRowElement
  /** absolute pos of the docTableRow node */
  pos: number
}

function columnHandleLive(view: EditorView): boolean {
  const state = columnResizingPluginKey.getState(view.state) as
    { activeHandle: number; dragging: unknown } | undefined
  return !!state && (state.activeHandle >= 0 || !!state.dragging)
}

/** the row whose bottom border lies under the pointer (page-gap rows are not rows) */
export function rowBorderAt(view: EditorView, event: MouseEvent): RowTarget | null {
  const el = event.target instanceof Element ? event.target : null
  const cell = el?.closest('td,th')
  const table = cell?.closest('table.doc-table')
  if (!cell || !(table instanceof HTMLTableElement) || !view.dom.contains(table)) return null
  let hit: HTMLTableRowElement | null = null
  for (const tr of Array.from(table.rows)) {
    if (tr.closest('table') !== table || tr.classList.contains('page-gap-table')) continue
    if (Math.abs(tr.getBoundingClientRect().bottom - event.clientY) <= EDGE_PX) {
      hit = tr
      break
    }
  }
  if (!hit) return null
  let pos: number
  try {
    pos = view.posAtDOM(hit, 0)
  } catch {
    return null
  }
  const $pos = view.state.doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    if ($pos.node(depth).type.name !== 'docTableRow') continue
    const rowPos = $pos.before(depth)
    // a nested-table row or a paginated repeat-header clone resolves to the
    // enclosing model row, whose DOM is a different element
    return view.nodeDOM(rowPos) === hit ? { tr: hit, pos: rowPos } : null
  }
  return null
}

/** the new w:trHeight for a row drawn `startPx` tall after a pointer delta of `deltaPx` */
export function draggedRowHeightTwips(startPx: number, deltaPx: number): number {
  return Math.round(Math.max(MIN_ROW_PX, startPx + deltaPx) * 15)
}

/**
 * Height the row's rule adds on top of w:trHeight (cell margins, gridline
 * terms). Measured, not computed: the terms are CSS variables, and the row's
 * used height would also count cell content.
 */
export function rowChromePx(tr: HTMLElement, heightTwips: number): number {
  const cell = tr.querySelector('td,th')
  const own = `${((heightTwips / 1440) * 96).toFixed(1)}px`
  // the authored attribute, not CSSOM: that reserializes 40.0px as 40px and folds calc()
  const expr = /(?:^|;)\s*height:([^;]+)/.exec(tr.getAttribute('style') ?? '')?.[1]?.trim()
  if (!(cell instanceof HTMLElement) || !expr?.includes(own)) return 0
  const probe = tr.ownerDocument.createElement('div')
  probe.style.cssText = `position:absolute;visibility:hidden;height:${expr.replace(own, '0px')}`
  cell.appendChild(probe)
  const scale = tr.offsetHeight > 0 ? tr.getBoundingClientRect().height / tr.offsetHeight : 1
  const chrome = probe.getBoundingClientRect().height / scale
  probe.remove()
  return Number.isFinite(chrome) ? chrome : 0
}

function startDrag(view: EditorView, event: MouseEvent, target: RowTarget): boolean {
  const { tr, pos } = target
  const win = view.dom.ownerDocument.defaultView ?? window
  const startCss = tr.offsetHeight
  const screenPerCss = startCss > 0 ? tr.getBoundingClientRect().height / startCss : 1
  const startY = event.clientY
  const cssAt = (clientY: number) =>
    Math.max(MIN_ROW_PX, startCss + (clientY - startY) / screenPerCss)
  const preview = (heightPx: number | null) =>
    previewRowHeights(view, heightPx === null ? null : [{ pos, heightPx }])

  const cleanup = () => {
    win.removeEventListener('mousemove', move)
    win.removeEventListener('mouseup', finish)
    win.removeEventListener('keydown', cancel, true)
  }
  function move(moveEvent: MouseEvent) {
    if (moveEvent.buttons === 0) return finish(moveEvent)
    preview(cssAt(moveEvent.clientY))
  }
  function finish(upEvent: MouseEvent) {
    cleanup()
    preview(null)
    const deltaPx = (upEvent.clientY - startY) / screenPerCss
    if (Math.abs(deltaPx) < 0.5) return
    const row = view.state.doc.nodeAt(pos)
    if (!row || row.type.name !== 'docTableRow') return
    const heightRule = row.attrs.heightRule === 'exact' ? 'exact' : 'atLeast'
    const commit = (heightTwips: number, correction: boolean) =>
      view.dispatch(
        view.state.tr
          .setNodeMarkup(pos, undefined, { ...row.attrs, heightTwips, heightRule })
          .setMeta('addToHistory', !correction),
      )
    const twips = draggedRowHeightTwips(startCss, deltaPx)
    commit(twips, false)
    // w:trHeight is a content floor: cell margins and gridlines render on top of
    // it, so take that chrome back out and the border stays under the pointer
    // (the correction shares the undo step)
    const painted = view.nodeDOM(pos)
    if (!(painted instanceof HTMLElement)) return
    const chromePx = rowChromePx(painted, twips)
    if (chromePx > 0.5) commit(Math.max(15, twips - Math.round(chromePx * 15)), true)
  }
  function cancel(keyEvent: KeyboardEvent) {
    if (keyEvent.key !== 'Escape') return
    keyEvent.preventDefault()
    keyEvent.stopPropagation()
    cleanup()
    preview(null)
  }
  win.addEventListener('mousemove', move)
  win.addEventListener('mouseup', finish)
  win.addEventListener('keydown', cancel, true)
  event.preventDefault()
  return true
}

/** Word's horizontal border drag: pulling a row's bottom border sets that row's height */
export function tableRowDrag(): Plugin {
  return new Plugin<RowHeightPreview[] | null>({
    key: tableRowDragKey,
    state: {
      init: () => null,
      apply(tr, prev) {
        const meta = tr.getMeta(tableRowDragKey) as RowHeightPreview[] | null | undefined
        if (meta !== undefined) return meta
        if (!prev || !tr.docChanged) return prev
        const kept: RowHeightPreview[] = []
        for (const row of prev) {
          const pos = tr.mapping.mapResult(row.pos)
          if (!pos.deleted) kept.push({ ...row, pos: pos.pos })
        }
        return kept.length ? kept : null
      },
    },
    view(view) {
      const onMouseMove = (event: MouseEvent) => {
        const armed = view.editable && !columnHandleLive(view) && rowBorderAt(view, event) !== null
        view.dom.classList.toggle(CURSOR_CLASS, armed)
      }
      const disarm = () => view.dom.classList.remove(CURSOR_CLASS)
      view.dom.addEventListener('mousemove', onMouseMove)
      view.dom.addEventListener('mouseleave', disarm)
      return {
        destroy() {
          view.dom.removeEventListener('mousemove', onMouseMove)
          view.dom.removeEventListener('mouseleave', disarm)
          disarm()
        },
      }
    },
    props: {
      decorations(state) {
        const drag = tableRowDragKey.getState(state)
        if (!drag) return DecorationSet.empty
        const decorations: Decoration[] = []
        for (const { pos, heightPx } of drag) {
          const row = state.doc.nodeAt(pos)
          if (!row || row.type.name !== 'docTableRow') continue
          decorations.push(
            Decoration.node(pos, pos + row.nodeSize, { style: `height:${heightPx.toFixed(1)}px` }),
          )
        }
        return DecorationSet.create(state.doc, decorations)
      },
      handleDOMEvents: {
        mousedown: (view, event) => {
          if (!view.editable || event.button !== 0 || columnHandleLive(view)) return false
          const target = rowBorderAt(view, event)
          return target ? startDrag(view, event, target) : false
        },
      },
    },
  })
}

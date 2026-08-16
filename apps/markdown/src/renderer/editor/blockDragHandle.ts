import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { installPopoverDismiss } from '@genoffice/ui'
import { t } from '../i18n/locale'

/**
 * Notion-style block gutter: a `+` (insert below, opens the slash menu) and a
 * grip that follows the hovered top-level block. Dragging the grip hands the
 * block to ProseMirror's native drag machinery; clicking it opens a small
 * block menu (duplicate / move / delete — backed by the BlockKeymap commands).
 */
export const BlockDragHandle = Extension.create({
  name: 'blockDragHandle',

  addProseMirrorPlugins() {
    return [dragHandlePlugin(this.editor)]
  },
})

function topLevelPosAt(view: EditorView, coords: { left: number; top: number }): number | null {
  const found = view.posAtCoords(coords)
  if (!found) return null
  const $pos = view.state.doc.resolve(found.inside >= 0 ? found.inside : found.pos)
  if ($pos.depth === 0 && found.inside >= 0) return found.inside
  return $pos.depth > 0 ? $pos.before(1) : null
}

const GRIP_SVG =
  '<svg width="10" height="16" viewBox="0 0 10 16" aria-hidden="true">' +
  '<circle cx="2.5" cy="3" r="1.4" fill="currentColor"/><circle cx="7.5" cy="3" r="1.4" fill="currentColor"/>' +
  '<circle cx="2.5" cy="8" r="1.4" fill="currentColor"/><circle cx="7.5" cy="8" r="1.4" fill="currentColor"/>' +
  '<circle cx="2.5" cy="13" r="1.4" fill="currentColor"/><circle cx="7.5" cy="13" r="1.4" fill="currentColor"/>' +
  '</svg>'

const PLUS_SVG =
  '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" ' +
  'stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M8 3v10M3 8h10"/></svg>'

function dragHandlePlugin(editor: Editor): Plugin {
  return new Plugin({
    key: new PluginKey('blockDragHandle'),
    view(view) {
      const container = view.dom.parentElement

      const handle = document.createElement('div')
      handle.className = 'md-block-gutter'
      handle.style.display = 'none'
      handle.setAttribute('contenteditable', 'false')

      const plus = document.createElement('button')
      plus.type = 'button'
      plus.className = 'md-gutter-btn md-gutter-plus'
      plus.innerHTML = PLUS_SVG
      plus.dataset.tip = t('blockAddBelow')
      plus.setAttribute('aria-label', t('blockAddBelow'))

      const grip = document.createElement('button')
      grip.type = 'button'
      grip.className = 'md-gutter-btn md-gutter-grip'
      grip.draggable = true
      grip.innerHTML = GRIP_SVG
      grip.dataset.tip = t('blockGripHint')
      grip.setAttribute('aria-label', t('blockGripHint'))

      handle.append(plus, grip)

      const menu = document.createElement('div')
      menu.className = 'md-block-menu'
      menu.style.display = 'none'

      if (container) {
        container.style.position ||= 'relative'
        container.append(handle, menu)
      }

      let hoverPos: number | null = null
      let menuPos: number | null = null

      const hideHandle = () => {
        handle.style.display = 'none'
        hoverPos = null
      }

      // Unified dismissal (outside press / window blur / shell chrome press);
      // installed on open, torn down here so it runs exactly once per open —
      // every close path (item click, scroll, doc change, destroy, the
      // installer's own close) funnels through hideMenu.
      let offDismiss: (() => void) | null = null

      const hideMenu = () => {
        menu.style.display = 'none'
        menuPos = null
        offDismiss?.()
        offDismiss = null
      }

      const onMouseMove = (event: MouseEvent) => {
        if (!view.editable || !container) return
        const pos = topLevelPosAt(view, { left: event.clientX, top: event.clientY })
        if (pos === null) return
        const node = view.state.doc.nodeAt(pos)
        if (!node) return hideHandle()
        const dom = view.nodeDOM(pos)
        if (!(dom instanceof HTMLElement)) return hideHandle()
        const blockRect = dom.getBoundingClientRect()
        const containerRect = container.getBoundingClientRect()
        // DOMRect values are viewport pixels, while absolute offsets inside a
        // CSS-zoomed page are layout pixels. Convert the deltas back so the
        // gutter is scaled exactly once with the rest of the page.
        const scale = container.offsetWidth ? containerRect.width / container.offsetWidth : 1
        hoverPos = pos
        handle.style.display = 'flex'
        handle.style.top = `${(blockRect.top - containerRect.top) / scale + container.scrollTop + 2}px`
        handle.style.left = `${Math.max(0, (blockRect.left - containerRect.left) / scale - 52)}px`
      }

      // hoverPos was computed on a past mousemove — validate before selecting
      const selectHovered = (): NodeSelection | null => {
        const pos = menuPos ?? hoverPos
        if (pos === null || pos >= view.state.doc.content.size) return null
        if (!view.state.doc.nodeAt(pos)) return null
        try {
          return NodeSelection.create(view.state.doc, pos)
        } catch {
          return null
        }
      }

      const onDragStart = (event: DragEvent) => {
        hideMenu()
        if (!event.dataTransfer) return
        const selection = selectHovered()
        if (!selection) return event.preventDefault()
        view.dispatch(view.state.tr.setSelection(selection))
        view.dragging = { slice: selection.content(), move: true }
        event.dataTransfer.effectAllowed = 'move'
        event.dataTransfer.setData('text/plain', ' ')
        const dom = view.nodeDOM(selection.from)
        if (dom instanceof HTMLElement) event.dataTransfer.setDragImage(dom, 0, 0)
      }

      // insert an empty paragraph below the block, type "/" so the slash menu opens
      const onPlusClick = () => {
        hideMenu()
        const selection = selectHovered()
        if (!selection) return
        const after = selection.to
        const paragraph = editor.schema.nodes.paragraph!.create()
        let tr = view.state.tr.insert(after, paragraph)
        tr = tr.setSelection(TextSelection.near(tr.doc.resolve(after + 1)))
        view.dispatch(tr)
        view.focus()
        editor.commands.insertContent('/')
      }

      const menuAction = (run: () => void) => () => {
        const selection = selectHovered()
        hideMenu()
        if (!selection) return
        view.dispatch(view.state.tr.setSelection(selection))
        run()
        view.focus()
      }

      const MENU_ITEMS: Array<{
        labelKey: Parameters<typeof t>[0]
        danger?: boolean
        run: () => void
      }> = [
        { labelKey: 'blockAddBelow', run: () => onPlusClick() },
        { labelKey: 'blockDuplicate', run: () => void editor.commands.duplicateBlock() },
        { labelKey: 'blockMoveUp', run: () => void editor.commands.moveBlockUp() },
        { labelKey: 'blockMoveDown', run: () => void editor.commands.moveBlockDown() },
        { labelKey: 'blockDelete', danger: true, run: () => void editor.commands.deleteBlock() },
      ]

      const openMenu = () => {
        if (hoverPos === null || !container) return
        menuPos = hoverPos
        menu.replaceChildren(
          ...MENU_ITEMS.map((item) => {
            const btn = document.createElement('button')
            btn.type = 'button'
            btn.className = `md-block-menu-item${item.danger ? ' danger' : ''}`
            btn.textContent = t(item.labelKey)
            btn.addEventListener(
              'click',
              item.labelKey === 'blockAddBelow' ? item.run : menuAction(item.run),
            )
            return btn
          }),
        )
        menu.style.display = 'block'
        menu.style.top = `${parseFloat(handle.style.top) + 26}px`
        menu.style.left = handle.style.left
        // presses inside the menu or the gutter are the popover's own; anything
        // else closes it. Hiding the gutter too matches the old outside-press
        // behavior — the editor's mousemove re-syncs it on the next hover.
        offDismiss ??= installPopoverDismiss(
          () => {
            hideMenu()
            hideHandle()
          },
          { inside: () => [menu, handle] },
        )
      }

      const onGripClick = () => {
        if (menu.style.display === 'block') hideMenu()
        else openMenu()
      }

      const onScrollOrLeave = () => {
        hideHandle()
        hideMenu()
      }

      // The gutter sits 52px left of the block with a page-background gap in
      // between, so reaching it always raises mouseleave on the editor first.
      // Hide on a grace period instead of instantly; entering the gutter (or
      // coming back into the editor) cancels the pending hide.
      let hideTimer: number | null = null
      const cancelHide = () => {
        if (hideTimer !== null) window.clearTimeout(hideTimer)
        hideTimer = null
      }
      const scheduleHide = () => {
        cancelHide()
        hideTimer = window.setTimeout(() => {
          hideTimer = null
          if (menu.style.display !== 'block') hideHandle()
        }, 250)
      }

      view.dom.addEventListener('mousemove', onMouseMove)
      view.dom.addEventListener('mouseenter', cancelHide)
      view.dom.addEventListener('mouseleave', scheduleHide)
      handle.addEventListener('mousemove', (e) => e.stopPropagation())
      handle.addEventListener('mouseenter', cancelHide)
      handle.addEventListener('mouseleave', scheduleHide)
      grip.addEventListener('dragstart', onDragStart)
      grip.addEventListener('click', onGripClick)
      plus.addEventListener('click', onPlusClick)
      document.addEventListener('scroll', onScrollOrLeave, true)

      return {
        update(_view, prevState) {
          // any doc change invalidates the menu's captured position
          if (!prevState.doc.eq(view.state.doc)) hideMenu()
        },
        destroy() {
          cancelHide()
          hideMenu() // also tears down the popover-dismiss listeners
          view.dom.removeEventListener('mousemove', onMouseMove)
          view.dom.removeEventListener('mouseenter', cancelHide)
          view.dom.removeEventListener('mouseleave', scheduleHide)
          document.removeEventListener('scroll', onScrollOrLeave, true)
          handle.remove()
          menu.remove()
        },
      }
    },
  })
}

import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { TextEditOverlay } from '../src/renderer/TextEditOverlay'
import { ContextMenu } from '../src/renderer/components/ContextMenu'

const node = {
  type: 'text',
  sourceId: 's1',
  box: { x: 0, y: 0, w: 200, h: 40 },
} as unknown as ShapeRenderNode

const roots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: React.ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  roots.push({ root, container })
  return container
}

function selectChars(editor: HTMLElement, start: number, end: number) {
  editor.textContent = 'hello world'
  const range = document.createRange()
  range.setStart(editor.firstChild!, start)
  range.setEnd(editor.firstChild!, end)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
}

afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
})

describe('right-click inside the text edit overlay', () => {
  it('opens the text menu at the pointer without committing or moving the selection', () => {
    const onCommit = vi.fn()
    const onCancel = vi.fn()
    const onContextMenu = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, { node, scale: 1, onCommit, onCancel, onContextMenu }),
    )
    const editor = container.querySelector('.slide-text-editor') as HTMLElement
    selectChars(editor, 0, 5)

    const ev = new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: 30,
      clientY: 20,
    })
    act(() => void editor.dispatchEvent(ev))

    expect(ev.defaultPrevented).toBe(true)
    expect(onContextMenu).toHaveBeenCalledWith(30, 20, false)
    expect(onCommit).not.toHaveBeenCalled()
    expect(onCancel).not.toHaveBeenCalled()
    const r = window.getSelection()!.getRangeAt(0)
    expect([r.startOffset, r.endOffset]).toEqual([0, 5])
  })

  it('reports a caret-only selection as collapsed', () => {
    const onContextMenu = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node,
        scale: 1,
        onCommit: vi.fn(),
        onCancel: vi.fn(),
        onContextMenu,
      }),
    )
    const editor = container.querySelector('.slide-text-editor') as HTMLElement
    selectChars(editor, 3, 3)
    act(() => {
      editor.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }))
    })
    expect(onContextMenu).toHaveBeenLastCalledWith(0, 0, true)
  })

  it('Escape with the menu open only closes the menu; the edit keeps going', () => {
    const onCommit = vi.fn()
    const onClose = vi.fn()
    const container = mount(
      createElement(TextEditOverlay, {
        node,
        scale: 1,
        onCommit,
        onCancel: onCommit,
        onContextMenu: vi.fn(),
      }),
    )
    mount(createElement(ContextMenu, { x: 0, y: 0, items: [], onClose, keepEdit: true }))
    const editor = container.querySelector('.slide-text-editor') as HTMLElement
    selectChars(editor, 0, 5)
    act(() => {
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(onCommit).not.toHaveBeenCalled()
    expect(
      [...document.querySelectorAll('.ctx-menu')].every((m) => m.hasAttribute('data-keep-edit')),
    ).toBe(true)
  })
})

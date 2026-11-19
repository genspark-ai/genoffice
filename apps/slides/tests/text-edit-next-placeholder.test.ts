import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { TextEditOverlay } from '../src/renderer/TextEditOverlay'

const node = {
  type: 'text',
  sourceId: 's1',
  placeholder: 'title',
  box: { x: 0, y: 0, w: 200, h: 40 },
} as unknown as ShapeRenderNode

const roots: Array<{ root: Root; container: HTMLElement }> = []

function mount(props: Partial<Parameters<typeof TextEditOverlay>[0]>): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(TextEditOverlay, {
        node,
        scale: 1,
        onCommit: vi.fn(),
        onCancel: vi.fn(),
        ...props,
      }),
    ),
  )
  roots.push({ root, container })
  return container.querySelector('.slide-text-editor') as HTMLElement
}

function press(editor: HTMLElement, init: KeyboardEventInit): KeyboardEvent {
  const ev = new KeyboardEvent('keydown', {
    key: 'Enter',
    bubbles: true,
    cancelable: true,
    ...init,
  })
  act(() => void editor.dispatchEvent(ev))
  return ev
}

afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
})

describe('Ctrl+Enter in the text edit overlay', () => {
  it('hands unchanged text to the next-placeholder handler as null', () => {
    const onNextPlaceholder = vi.fn()
    const onCancel = vi.fn()
    const editor = mount({ onNextPlaceholder, onCancel })

    const ev = press(editor, { ctrlKey: true })

    expect(ev.defaultPrevented).toBe(true)
    expect(onNextPlaceholder).toHaveBeenCalledWith(null)
    expect(onCancel).not.toHaveBeenCalled()
  })

  it('hands changed text to the handler instead of committing', () => {
    const onNextPlaceholder = vi.fn()
    const onCommit = vi.fn()
    const editor = mount({ onNextPlaceholder, onCommit })
    editor.textContent = 'hello'

    press(editor, { ctrlKey: true })

    expect(onCommit).not.toHaveBeenCalled()
    const paras = onNextPlaceholder.mock.calls[0][0]
    expect(paras[0].runs.map((r: { text: string }) => r.text).join('')).toBe('hello')
  })

  it('keeps Cmd+Enter as a plain commit even with the handler wired', () => {
    const onNextPlaceholder = vi.fn()
    const onCommit = vi.fn()
    const editor = mount({ onNextPlaceholder, onCommit })
    editor.textContent = 'hello'

    press(editor, { metaKey: true })

    expect(onNextPlaceholder).not.toHaveBeenCalled()
    expect(onCommit).toHaveBeenCalledTimes(1)
  })

  it('falls back to commit when no handler is provided (mac)', () => {
    const onCommit = vi.fn()
    const editor = mount({ onCommit })
    editor.textContent = 'hello'

    const ev = press(editor, { ctrlKey: true })

    expect(ev.defaultPrevented).toBe(true)
    expect(onCommit).toHaveBeenCalledTimes(1)
  })
})

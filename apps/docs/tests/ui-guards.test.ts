import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { FindPanel, findMatches, foldCase } from '../src/renderer/components/FindPanel'
import { SCAN_THRESHOLD } from '../src/renderer/editor/find'
import {
  clampPictureTwips,
  PICTURE_TWIPS_MAX,
  PICTURE_TWIPS_MIN,
} from '../src/renderer/components/Ribbon'

function createEditor(text: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text }],
        },
      ],
    },
  })
}

function render(element: React.ReactElement): {
  container: HTMLElement
  rerender: (next: React.ReactElement) => void
  unmount: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    rerender: (next) => act(() => root.render(next)),
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

beforeEach(() => {
  Element.prototype.scrollIntoView = vi.fn()
})

describe('foldCase', () => {
  it('lowercases without changing string length', () => {
    expect(foldCase('ABC def')).toBe('abc def')
    // 'İ'.toLowerCase() is 2 chars ('i' + combining dot) and would shift offsets
    expect('İ'.toLowerCase().length).toBe(2)
    expect(foldCase('İstanbul').length).toBe('İstanbul'.length)
    expect(foldCase('AİB')).toBe('aİb')
  })
})

describe('findMatches', () => {
  it('keeps offsets aligned after a length-changing lowercase char', () => {
    const editor = createEditor('İİİ test')
    const [m] = findMatches(editor, 'TEST', { matchCase: false, wholeWord: false })
    expect(m).toBeDefined()
    expect(editor.state.doc.textBetween(m.from, m.to)).toBe('test')
    editor.destroy()
  })

  it('respects matchCase and wholeWord', () => {
    const editor = createEditor('Cat cats CAT')
    expect(findMatches(editor, 'cat', { matchCase: false, wholeWord: false })).toHaveLength(3)
    expect(findMatches(editor, 'cat', { matchCase: false, wholeWord: true })).toHaveLength(2)
    expect(findMatches(editor, 'CAT', { matchCase: true, wholeWord: false })).toHaveLength(1)
    editor.destroy()
  })
})

describe('clampPictureTwips', () => {
  it('clamps to the Word picture size range', () => {
    expect(clampPictureTwips(1)).toBe(PICTURE_TWIPS_MIN)
    expect(clampPictureTwips(999999)).toBe(PICTURE_TWIPS_MAX)
    expect(clampPictureTwips(5670)).toBe(5670)
  })
})

describe('FindPanel', () => {
  afterEach(() => {
    vi.useRealTimers()
  })

  it('hides the Replace tab when the editor is read-only', () => {
    const editor = createEditor('hello world')
    editor.setEditable(false)
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    expect(container.querySelectorAll('.find-tab')).toHaveLength(2)
    expect(container.querySelector('.find-action')).toBeNull()
    unmount()
    editor.destroy()
  })

  it('shows the replace row on the Replace tab of an editable document', () => {
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    expect(container.querySelectorAll('.find-tab')).toHaveLength(3)
    expect(container.querySelectorAll('.find-row')).toHaveLength(2)
    act(() => {
      container.querySelectorAll<HTMLButtonElement>('.find-tab')[1].click()
    })
    expect(container.querySelectorAll('.find-row')).toHaveLength(3)
    expect(container.querySelectorAll('.find-action')).toHaveLength(2)
    unmount()
    editor.destroy()
  })

  it('debounces the scan while typing', () => {
    vi.useFakeTimers()
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    const input = container.querySelector<HTMLInputElement>('.find-input')!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(input, 'hello')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const count = () => container.querySelector('.find-count')!.textContent
    expect(count()).not.toContain('1/1')
    act(() => {
      vi.advanceTimersByTime(200)
    })
    expect(count()).toBe('1/1')
    unmount()
    editor.destroy()
  })

  it('holds Next and Replace while a time-sliced scan is in flight', async () => {
    vi.useFakeTimers()
    const saved = SCAN_THRESHOLD.nodeSize
    SCAN_THRESHOLD.nodeSize = 0
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    try {
      act(() => {
        container.querySelectorAll<HTMLButtonElement>('.find-tab')[1].click()
      })
      const input = container.querySelector<HTMLInputElement>('.find-input')!
      const setValue = (el: HTMLInputElement, v: string) => {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
        setter.call(el, v)
        el.dispatchEvent(new Event('input', { bubbles: true }))
      }
      act(() => setValue(input, 'hello'))
      act(() => {
        vi.advanceTimersByTime(200)
      })
      // the debounce fired and the scan is pending: hits are not known yet
      expect(container.querySelector('.find-count')!.textContent).toBe('\u2026')
      const replaceAll = container.querySelectorAll<HTMLButtonElement>('.find-action')[1]
      expect(replaceAll.disabled).toBe(true)
      act(() => setValue(container.querySelectorAll<HTMLInputElement>('.find-input')[1], 'X'))
      act(() => replaceAll.click())
      act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })))
      expect(editor.state.doc.textContent).toBe('hello world')
      await act(async () => {
        await vi.runAllTimersAsync()
      })
      expect(container.querySelector('.find-count')!.textContent).toBe('1/1')
      expect(replaceAll.disabled).toBe(false)
    } finally {
      SCAN_THRESHOLD.nodeSize = saved
      unmount()
      editor.destroy()
    }
  })

  it('focuses the find input on mount and again when focusFindNonce bumps', () => {
    const editor = createEditor('hello world')
    const { container, rerender, unmount } = render(
      createElement(FindPanel, { editor, onClose: () => {}, focusFindNonce: 0 }),
    )
    const input = container.querySelector<HTMLInputElement>('.find-input')!
    expect(document.activeElement).toBe(input)
    // user clicks back into the document, then hits Ctrl+F with the panel open
    act(() => {
      input.blur()
    })
    expect(document.activeElement).not.toBe(input)
    rerender(createElement(FindPanel, { editor, onClose: () => {}, focusFindNonce: 1 }))
    expect(document.activeElement).toBe(input)
    unmount()
    editor.destroy()
  })

  it('labels both inputs and announces the match count', () => {
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    act(() => {
      container.querySelectorAll<HTMLButtonElement>('.find-tab')[1].click()
    })
    const [find, replace] = Array.from(container.querySelectorAll<HTMLInputElement>('.find-input'))
    expect(find!.getAttribute('aria-label')).toBeTruthy()
    expect(find!.getAttribute('aria-label')).toBe(find!.placeholder)
    expect(replace!.getAttribute('aria-label')).toBeTruthy()
    expect(replace!.getAttribute('aria-label')).toBe(replace!.placeholder)
    expect(container.querySelector('.find-count')!.getAttribute('aria-live')).toBe('polite')
    unmount()
    editor.destroy()
  })
})

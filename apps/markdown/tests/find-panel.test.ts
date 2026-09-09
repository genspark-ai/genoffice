import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { FindPanel, findInText, foldCase, type FindPanelStrings } from '@genoffice/ui'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { findMatches, tiptapFindTarget } from '../src/renderer/editor/findTarget'
import { searchPluginKey } from '../src/renderer/editor/searchHighlight'

const STRINGS: FindPanelStrings = {
  findPlaceholder: 'Find',
  replacePlaceholder: 'Replace with',
  matchCase: 'Match case',
  wholeWord: 'Whole words only',
  noResults: 'No results',
  prevMatch: 'Previous',
  nextMatch: 'Next',
  closeEsc: 'Close',
  replace: 'Replace',
  replaceAll: 'Replace All',
}

function createEditor(markdown: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: markdown,
    contentType: 'markdown',
  })
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function type(input: HTMLInputElement, value: string): void {
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

beforeEach(() => {
  Element.prototype.scrollIntoView = vi.fn()
})

afterEach(() => {
  vi.useRealTimers()
})

describe('find helpers', () => {
  it('foldCase never changes string length', () => {
    expect(foldCase('ABC def')).toBe('abc def')
    expect(foldCase('İstanbul').length).toBe('İstanbul'.length)
  })

  it('findInText honors matchCase and wholeWord', () => {
    expect(findInText('Cat cats CAT', 'cat', { matchCase: false, wholeWord: false })).toEqual([
      0, 4, 9,
    ])
    expect(findInText('Cat cats CAT', 'cat', { matchCase: false, wholeWord: true })).toEqual([0, 9])
    expect(findInText('Cat cats CAT', 'CAT', { matchCase: true, wholeWord: false })).toEqual([9])
    expect(findInText('aaaa', 'aa', { matchCase: true, wholeWord: false })).toEqual([0, 2])
  })
})

describe('findMatches', () => {
  it('finds text across marks and inside list items', () => {
    const editor = createEditor('# Title\n\nSome **bold title** here\n\n- title in list\n')
    const ranges = findMatches(editor, 'title', { matchCase: false, wholeWord: false })
    expect(ranges).toHaveLength(3)
    for (const r of ranges) {
      expect(editor.state.doc.textBetween(r.from, r.to).toLowerCase()).toBe('title')
    }
    editor.destroy()
  })

  it('keeps offsets aligned after a length-changing lowercase char', () => {
    const editor = createEditor('İİİ test')
    const [m] = findMatches(editor, 'TEST', { matchCase: false, wholeWord: false })
    expect(m).toBeDefined()
    expect(editor.state.doc.textBetween(m!.from, m!.to)).toBe('test')
    editor.destroy()
  })
})

describe('tiptapFindTarget', () => {
  it('replaces one match and then all remaining ones', () => {
    const editor = createEditor('foo bar foo\n\nfoo')
    const target = tiptapFindTarget(editor)
    expect(target.search('foo', { matchCase: false, wholeWord: false }, 0)).toBe(3)
    target.replaceOne(1, 'baz')
    expect(editor.getText()).toContain('foo bar baz')
    expect(target.search('foo', { matchCase: false, wholeWord: false }, 0)).toBe(2)
    target.replaceAll('qux')
    expect(target.search('foo', { matchCase: false, wholeWord: false }, 0)).toBe(0)
    expect(editor.getText()).toBe('qux bar baz\n\nqux')
    editor.destroy()
  })

  it('reports document changes and skips highlight-only transactions', () => {
    const editor = createEditor('hello')
    const target = tiptapFindTarget(editor)
    const listener = vi.fn()
    const off = target.onDocChanged(listener)
    target.search('hello', { matchCase: false, wholeWord: false }, 0)
    expect(listener).not.toHaveBeenCalled()
    editor.commands.insertContentAt(editor.state.doc.content.size, ' world')
    expect(listener).toHaveBeenCalledTimes(1)
    off()
    editor.commands.insertContentAt(editor.state.doc.content.size, '!')
    expect(listener).toHaveBeenCalledTimes(1)
    editor.destroy()
  })
})

describe('FindPanel', () => {
  it('debounces the scan, navigates with Enter and replaces all', () => {
    vi.useFakeTimers()
    const editor = createEditor('one two one')
    const target = tiptapFindTarget(editor)
    const onClose = vi.fn()
    const { container, unmount } = render(
      createElement(FindPanel, { target, strings: STRINGS, onClose }),
    )
    expect(container.querySelectorAll('.find-row')).toHaveLength(2)
    const [find, replace] = Array.from(container.querySelectorAll<HTMLInputElement>('.find-input'))
    type(find!, 'one')
    const count = () => container.querySelector('.find-count')!.textContent
    expect(count()).not.toBe('1/2')
    act(() => {
      vi.advanceTimersByTime(200)
    })
    expect(count()).toBe('1/2')
    act(() => {
      find!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })
    expect(count()).toBe('2/2')

    type(replace!, 'uno')
    act(() => {
      container
        .querySelectorAll<HTMLButtonElement>('.find-action')[1]!
        .dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(editor.getText()).toBe('uno two uno')
    expect(count()).toBe('No results')

    act(() => {
      find!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
    editor.destroy()
  })

  it('clears the painted hits when unmounted by the host', () => {
    vi.useFakeTimers()
    const editor = createEditor('one two one')
    const target = tiptapFindTarget(editor)
    const { container, unmount } = render(
      createElement(FindPanel, { target, strings: STRINGS, onClose() {} }),
    )
    type(container.querySelector<HTMLInputElement>('.find-input')!, 'one')
    act(() => {
      vi.advanceTimersByTime(200)
    })
    const hits = () => searchPluginKey.getState(editor.state)!.find().length
    expect(hits()).toBe(2)
    unmount()
    expect(hits()).toBe(0)
    editor.destroy()
  })

  it('refocuses the requested field when the host bumps the focus request', () => {
    const editor = createEditor('hello')
    const target = tiptapFindTarget(editor)
    const container = document.createElement('div')
    document.body.appendChild(container)
    const root = createRoot(container)
    const show = (nonce: number, field: 'find' | 'replace') =>
      act(() => {
        root.render(
          createElement(FindPanel, {
            target,
            strings: STRINGS,
            onClose() {},
            focusRequest: { field, nonce },
          }),
        )
      })
    show(0, 'find')
    const [find, replace] = Array.from(container.querySelectorAll<HTMLInputElement>('.find-input'))
    expect(document.activeElement).toBe(find)
    find!.blur()
    expect(document.activeElement).not.toBe(find)
    show(1, 'find')
    expect(document.activeElement).toBe(find)
    show(2, 'replace')
    expect(document.activeElement).toBe(replace)
    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('hides the replace row on a read-only editor', () => {
    const editor = createEditor('hello')
    editor.setEditable(false)
    const { container, unmount } = render(
      createElement(FindPanel, {
        target: tiptapFindTarget(editor),
        strings: STRINGS,
        onClose() {},
      }),
    )
    expect(container.querySelectorAll('.find-row')).toHaveLength(1)
    unmount()
    editor.destroy()
  })
})

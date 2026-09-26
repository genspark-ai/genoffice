import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { ribbonProps } from './helpers/ribbon-props'

const setNative = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!

function editorWith(runs: Array<Record<string, unknown>>) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: runs.map((attrs, i) => ({
            type: 'text',
            text: `run${i} `,
            marks: [{ type: 'docTextStyle', attrs }],
          })),
        },
      ],
    },
  })
  editor.commands.selectAll()
  return editor
}

describe('font size box on a mixed selection', () => {
  it('reports mixed sizes only when the runs disagree', () => {
    const same = editorWith([{ sizeHalfPoints: 24 }, { sizeHalfPoints: 24 }])
    expect(computeFormatState(same).fontSizeMixed).toBe(false)
    same.destroy()
    const mixed = editorWith([{ sizeHalfPoints: 24 }, { sizeHalfPoints: 28 }])
    const state = computeFormatState(mixed)
    expect(state.fontSizeMixed).toBe(true)
    expect(state.fontSizePt).toBe(12)
    mixed.commands.setTextSelection(2)
    expect(computeFormatState(mixed).fontSizeMixed).toBe(false)
    mixed.destroy()
  })
})

describe('ribbon font and size combos', () => {
  let editor: Editor
  let host: HTMLDivElement
  let root: Root
  const render = () =>
    act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
  const fontBox = () => host.querySelector<HTMLInputElement>('input.rb-font-family')!
  const sizeBox = () => host.querySelector<HTMLInputElement>('input.rb-font-size')!
  const type = (box: HTMLInputElement, text: string) =>
    act(() => {
      setNative.call(box, text)
      box.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }))
    })
  const key = (box: HTMLInputElement, k: string) =>
    act(() => {
      box.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true }))
    })
  const menu = () => host.querySelector('.rb-font-family-menu')

  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    vi.stubGlobal(
      'ResizeObserver',
      class {
        observe() {}
        disconnect() {}
        unobserve() {}
      },
    )
    editor = editorWith([
      { fontAscii: 'Calibri', sizeHalfPoints: 24 },
      { fontAscii: 'Calibri', sizeHalfPoints: 28 },
    ])
    host = document.createElement('div')
    document.body.append(host)
    root = createRoot(host)
    render()
  })
  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    editor.destroy()
    vi.unstubAllGlobals()
  })

  it('blanks the size box on mixed sizes and Enter on a typed value applies it', () => {
    expect(sizeBox().value).toBe('')
    act(() => sizeBox().focus())
    type(sizeBox(), '14')
    key(sizeBox(), 'Enter')
    expect(editor.getAttributes('docTextStyle').sizeHalfPoints).toBe(28)
    render()
    // the test UI language is zh: 14pt reads as its CJK size name
    expect(sizeBox().value).toBe('\u56db\u53f7')
    act(() => sizeBox().focus())
    type(sizeBox(), '\u5c0f\u56db')
    key(sizeBox(), 'Enter')
    expect(editor.getAttributes('docTextStyle').sizeHalfPoints).toBe(24)
  })

  it('Esc and blur restore the previous value without applying', () => {
    act(() => sizeBox().focus())
    type(sizeBox(), '30')
    expect(sizeBox().value).toBe('30')
    key(sizeBox(), 'Escape')
    expect(sizeBox().value).toBe('')
    act(() => fontBox().focus())
    type(fontBox(), 'Georgia')
    act(() => fontBox().blur())
    expect(fontBox().value).toBe('Calibri')
    expect(editor.getAttributes('docTextStyle').fontAscii).toBe('Calibri')
    expect(editor.getAttributes('docTextStyle').sizeHalfPoints).toBe(24)
  })

  it('type-ahead opens the list, highlights the first match and Enter commits it', () => {
    act(() => fontBox().focus())
    type(fontBox(), 'ge')
    expect(menu()).not.toBeNull()
    expect(fontBox().value).toBe('Georgia')
    expect(fontBox().selectionStart).toBe(2)
    expect(fontBox().selectionEnd).toBe('Georgia'.length)
    const lit = menu()!.querySelector<HTMLButtonElement>('button.kbd-focus')
    expect(lit?.textContent).toBe('Georgia')
    key(fontBox(), 'ArrowDown')
    expect(menu()!.querySelector('button.kbd-focus')?.textContent).toBe('Verdana')
    expect(fontBox().value).toBe('Verdana')
    key(fontBox(), 'ArrowUp')
    expect(fontBox().value).toBe('Georgia')
    key(fontBox(), 'Escape')
    expect(menu()).toBeNull()
    expect(fontBox().value).toBe('Calibri')
    expect(editor.getAttributes('docTextStyle').fontAscii).toBe('Calibri')

    act(() => fontBox().focus())
    type(fontBox(), 'ta')
    key(fontBox(), 'Enter')
    expect(editor.getAttributes('docTextStyle').fontAscii).toBe('Tahoma')
  })

  it('opens on the common fonts section and commits a list click', () => {
    act(() => host.querySelector<HTMLButtonElement>('.rb-combo-caret')!.click())
    const labels = [...menu()!.querySelectorAll('.rb-menu-group-label')].map((el) => el.textContent)
    expect(labels[0]).toBe('\u5e38\u7528\u5b57\u4f53')
    const impact = [...menu()!.querySelectorAll<HTMLButtonElement>('button')].find(
      (el) => el.textContent === 'Impact',
    )!
    act(() => impact.click())
    expect(editor.getAttributes('docTextStyle').fontAscii).toBe('Impact')
    expect(menu()).toBeNull()
  })
})

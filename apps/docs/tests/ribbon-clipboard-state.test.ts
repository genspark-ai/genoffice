/**
 * The ribbon's Cut and Copy are greyed out by what is selected, by the rule
 * their menu items follow: nothing selected, nothing to cut or copy.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { clipboardOff } from '../src/renderer/commands/menu-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

const body = { hasSelection: false, editable: true, field: false }

describe('clipboard command state', () => {
  it('Cut needs a selection it may remove; Copy only needs a selection', () => {
    const selected = { ...body, hasSelection: true }
    expect(clipboardOff('copy', { hasDoc: true, target: body })).toBe(true)
    expect(clipboardOff('cut', { hasDoc: true, target: body })).toBe(true)
    expect(clipboardOff('copy', { hasDoc: true, target: selected })).toBe(false)
    expect(clipboardOff('cut', { hasDoc: true, target: selected })).toBe(false)
    const readOnly = { ...selected, editable: false }
    expect(clipboardOff('copy', { hasDoc: true, target: readOnly })).toBe(false)
    expect(clipboardOff('cut', { hasDoc: true, target: readOnly })).toBe(true)
  })

  it('with no document, only a focused text field has anything to copy', () => {
    const selected = { ...body, hasSelection: true }
    expect(clipboardOff('copy', { hasDoc: false, target: selected })).toBe(true)
    expect(clipboardOff('copy', { hasDoc: false, target: { ...selected, field: true } })).toBe(
      false,
    )
  })
})

describe('ribbon Cut and Copy', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'some words' }] }],
      },
    })
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  const buttons = () => {
    act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
    const byLabel = (key: 'ribbonCutTip' | 'ribbonCopyTip') =>
      container.querySelector<HTMLButtonElement>(`button[aria-label="${t(key)}"]`)!
    return { cut: byLabel('ribbonCutTip'), copy: byLabel('ribbonCopyTip') }
  }

  it('are greyed out with only a caret and live with a selection', () => {
    editor.commands.setTextSelection(1)
    let { cut, copy } = buttons()
    expect([cut.disabled, copy.disabled]).toEqual([true, true])
    editor.commands.selectAll()
    ;({ cut, copy } = buttons())
    expect([cut.disabled, copy.disabled]).toEqual([false, false])
  })
})

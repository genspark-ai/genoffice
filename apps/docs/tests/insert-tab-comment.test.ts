// Insert → Comment must start a new comment (not just toggle the pane), and it
// follows the Review-tab gate: live under the comments-only restriction, dead
// when a protection forbids commenting or there is nothing to anchor on.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

function makeEditor(text = 'hello world'): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', ...(text ? { content: [{ type: 'text', text }] } : {}) }],
    },
  })
}

describe('Insert → Comment', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onCommand = vi.fn()

  beforeEach(() => {
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
    onCommand.mockClear()
  })

  function renderInsertTab(overrides: Record<string, unknown> = {}) {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          onCommand,
          ...overrides,
        }),
      ),
    )
    const insertTab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (b) => b.textContent === t('ribbonTabInsert'),
    )!
    act(() => insertTab.click())
    return [...container.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonComment'),
    )!
  }

  it('starts a new comment when an anchor is available', () => {
    const btn = renderInsertTab()
    expect(btn.disabled).toBe(false)
    act(() => btn.click())
    expect(onCommand).toHaveBeenCalledTimes(1)
    expect(onCommand).toHaveBeenCalledWith('insert-comment')
  })

  it('is disabled with nothing to anchor on', () => {
    editor.destroy()
    editor = makeEditor('... ...')
    const btn = renderInsertTab()
    expect(btn.disabled).toBe(true)
  })

  it('is live in an empty paragraph, which holds the comment itself', () => {
    editor.destroy()
    editor = makeEditor('')
    const btn = renderInsertTab()
    expect(btn.disabled).toBe(false)
  })

  it('stays live under the comments-only restriction', () => {
    const btn = renderInsertTab({ isProtected: true, commentsAllowed: true })
    expect(btn.disabled).toBe(false)
  })

  it('is disabled when a protection forbids commenting', () => {
    const btn = renderInsertTab({ isProtected: true, commentsAllowed: false })
    expect(btn.disabled).toBe(true)
  })
})

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

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello world' }] }],
    },
  })
}

describe('Insert → Comment', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onNewComment = vi.fn()

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
    onNewComment.mockClear()
  })

  function renderInsertTab(overrides: Record<string, unknown>) {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          onNewComment,
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
    const btn = renderInsertTab({ canComment: true })
    expect(btn.disabled).toBe(false)
    act(() => btn.click())
    expect(onNewComment).toHaveBeenCalledTimes(1)
  })

  it('is disabled with nothing to anchor on', () => {
    const btn = renderInsertTab({ canComment: false })
    expect(btn.disabled).toBe(true)
  })

  it('stays live under the comments-only restriction', () => {
    const btn = renderInsertTab({ canComment: true, isProtected: true, commentsAllowed: true })
    expect(btn.disabled).toBe(false)
  })

  it('is disabled when a protection forbids commenting', () => {
    const btn = renderInsertTab({ canComment: true, isProtected: true, commentsAllowed: false })
    expect(btn.disabled).toBe(true)
  })
})

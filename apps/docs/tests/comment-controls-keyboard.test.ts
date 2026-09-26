import { describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { CommentsPanel } from '../src/renderer/components/CommentsPanel'

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

const noop = () => {}

function mount(onDelete = noop): { root: Root; container: HTMLElement } {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] }],
    },
  })
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const panel = createElement(CommentsPanel, {
    comments: [
      { id: 'c1', author: 'A', text: 'thread', done: false },
      { id: 'r1', author: 'B', text: 'reply', parentId: 'c1' },
    ] as CommentInfo[],
    docNode: editor.state.doc,
    composing: false,
    onSubmitNew: noop,
    onReply: noop,
    onEdit: noop,
    onResolve: noop,
    onCancelNew: noop,
    onDelete,
    onClose: noop,
  })
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: panel })))
  return { root, container }
}

const press = (el: HTMLElement, key: string) =>
  act(() => {
    el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  })

describe('comment controls keyboard activation', () => {
  it('Enter on the reply pencil opens the editor', () => {
    const { root, container } = mount()
    const pencil = container.querySelector<HTMLElement>('.comment-reply .comment-card-edit')!
    press(pencil, 'Enter')
    expect(container.querySelector('.comment-reply textarea')).not.toBeNull()
    act(() => root.unmount())
  })

  it('Space on the reply trash deletes the reply', () => {
    const onDelete = vi.fn()
    const { root, container } = mount(onDelete)
    press(container.querySelector<HTMLElement>('.comment-reply .comment-card-del')!, ' ')
    expect(onDelete).toHaveBeenCalledWith('r1')
    act(() => root.unmount())
  })

  it('Space on the thread trash deletes the thread', () => {
    const onDelete = vi.fn()
    const { root, container } = mount(onDelete)
    press(container.querySelector<HTMLElement>('.comment-card .comment-card-del')!, ' ')
    expect(onDelete).toHaveBeenCalledWith('c1')
    act(() => root.unmount())
  })
})

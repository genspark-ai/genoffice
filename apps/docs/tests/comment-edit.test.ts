import { describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { CommentsPanel } from '../src/renderer/components/CommentsPanel'
import { editComment, type ReviewContext } from '../src/renderer/review-actions'

// ---- editComment (review action) ----

function makeCtx(comments: CommentInfo[]): {
  ctx: ReviewContext
  state: { comments: CommentInfo[]; commentsDirty: boolean; dirty: boolean; status: string }
} {
  const state = { comments, commentsDirty: false, dirty: false, status: '' }
  const ctx = {
    editor: null,
    doc: null,
    dirtyRef: {
      get current() {
        return state.dirty
      },
      set current(v: boolean) {
        state.dirty = v
      },
    },
    setStatus: (s: string) => {
      state.status = s
    },
    setComments: (updater: unknown) => {
      state.comments = (updater as (prev: CommentInfo[]) => CommentInfo[])(state.comments)
    },
    setCommentsDirty: (v: boolean) => {
      state.commentsDirty = v
    },
  } as unknown as ReviewContext
  return { ctx, state }
}

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

describe('editComment', () => {
  const comments = (): CommentInfo[] => [
    { id: '1', author: 'Alice', date: '2026-01-01T00:00:00Z', text: 'first' },
    { id: '2', author: 'Bob', date: '2026-01-02T00:00:00Z', text: 'reply', parentId: '1' },
  ]

  it('replaces only the text, keeping author, date and thread structure', () => {
    const { ctx, state } = makeCtx(comments())
    editComment(ctx, '1', 'rewritten')
    expect(state.comments[0]).toEqual({
      id: '1',
      author: 'Alice',
      date: '2026-01-01T00:00:00Z',
      text: 'rewritten',
    })
    expect(state.comments[1]).toEqual(comments()[1])
    expect(state.commentsDirty).toBe(true)
    expect(state.dirty).toBe(true)
    expect(state.status).not.toBe('')
  })

  it('edits a reply without touching its parent', () => {
    const { ctx, state } = makeCtx(comments())
    editComment(ctx, '2', 'better reply')
    expect(state.comments[0]).toEqual(comments()[0])
    expect(state.comments[1]).toMatchObject({ id: '2', parentId: '1', text: 'better reply' })
  })
})

// ---- CommentsPanel edit flow ----

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] }],
    },
  })
}

function mount(element: React.ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return { root, container }
}

const noop = () => {}

describe('CommentsPanel editing', () => {
  it('opens a prefilled composer from the pencil and submits the new text', () => {
    const editor = makeEditor()
    const onEdit = vi.fn()
    const { root, container } = mount(
      createElement(CommentsPanel, {
        comments: [{ id: 'c1', author: 'A', text: 'old text', done: false }] as CommentInfo[],
        docNode: editor.state.doc,
        composing: false,
        onSubmitNew: noop,
        onReply: noop,
        onEdit,
        onResolve: noop,
        onCancelNew: noop,
        onDelete: noop,
        onClose: noop,
      }),
    )

    const pencil = container.querySelector<HTMLElement>('.comment-card-edit')!
    act(() => pencil.click())

    const textarea = container.querySelector<HTMLTextAreaElement>(
      '.comment-reply-compose textarea',
    )!
    expect(textarea.value).toBe('old text')

    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
      setter.call(textarea, 'new text')
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const save = [...container.querySelectorAll<HTMLButtonElement>('button.primary')].find(
      (b) => !b.disabled,
    )!
    act(() => save.click())

    expect(onEdit).toHaveBeenCalledWith('c1', 'new text')
    // the composer closes back to the plain card
    expect(container.querySelector('.comment-reply-compose')).toBeNull()

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  const mountOne = (handlers: {
    onEdit?: (id: string, text: string) => void
    onResolve?: (id: string, done: boolean) => void
  }) => {
    const editor = makeEditor()
    const mounted = mount(
      createElement(CommentsPanel, {
        comments: [{ id: 'c1', author: 'A', text: 'old text', done: false }] as CommentInfo[],
        docNode: editor.state.doc,
        composing: false,
        onSubmitNew: noop,
        onReply: noop,
        onEdit: handlers.onEdit ?? noop,
        onResolve: handlers.onResolve ?? noop,
        onCancelNew: noop,
        onDelete: noop,
        onClose: noop,
      }),
    )
    return { editor, ...mounted }
  }

  const typeInto = (textarea: HTMLTextAreaElement, value: string) => {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(textarea, value)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  }

  it('second pencil click cancels the edit instead of resetting the draft', () => {
    const onEdit = vi.fn()
    const { editor, root, container } = mountOne({ onEdit })
    const pencil = () => container.querySelector<HTMLElement>('.comment-card-edit')!
    act(() => pencil().click())
    const textarea = container.querySelector<HTMLTextAreaElement>(
      '.comment-reply-compose textarea',
    )!
    act(() => typeInto(textarea, 'half-typed'))

    act(() => pencil().click())
    // the composer is gone, nothing was submitted
    expect(container.querySelector('.comment-reply-compose')).toBeNull()
    expect(onEdit).not.toHaveBeenCalled()

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('resolving a thread mid-edit commits the draft instead of dropping it', () => {
    const onEdit = vi.fn()
    const onResolve = vi.fn()
    const { editor, root, container } = mountOne({ onEdit, onResolve })
    act(() => container.querySelector<HTMLElement>('.comment-card-edit')!.click())
    const textarea = container.querySelector<HTMLTextAreaElement>(
      '.comment-reply-compose textarea',
    )!
    act(() => typeInto(textarea, 'typed but not saved'))

    const resolveBtn = [...container.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === 'Resolve',
    )!
    act(() => resolveBtn.click())

    expect(onEdit).toHaveBeenCalledWith('c1', 'typed but not saved')
    expect(onResolve).toHaveBeenCalledWith('c1', true)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
})

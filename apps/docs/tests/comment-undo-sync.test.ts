/**
 * Comment metadata undo-sync: deleting a comment removes the mark (PM doc)
 * AND the metadata card (React state outside the document). Undo reverts the
 * mark — without the sync the metadata stayed deleted and the comment looked
 * lost. Pins the whole cycle: delete → undo restores card + text → redo
 * removes it again → undo restores once more.
 */
import { describe, expect, it } from 'vitest'
import { EditorState } from '@tiptap/pm/state'
import { Schema, Node as PmNode } from '@tiptap/pm/model'
import type { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import {
  attachCommentUndoSync,
  resetCommentUndoStore,
} from '../src/renderer/editor/comment-undo-sync'
import { resetCrossDocEditState, type CrossDocEditStateSink } from '../src/renderer/doc-dirty'

/** minimal sink satisfying the cross-doc reset (only the required setters) */
function cleanSink(): CrossDocEditStateSink {
  return {
    setSectionsDirty: () => {},
    setTrailingStartType: () => {},
    setSectionHfEdits: () => {},
    setHfLinks: () => {},
    setPgNumEdit: () => {},
    setPgNumDirtySections: () => {},
    setPendingNumbering: () => {},
    setStyleUpserts: () => {},
  }
}

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: { group: 'block', content: 'text*' },
    text: { group: 'inline' },
  },
  marks: {
    comment: { attrs: { ids: { default: '' } } },
  },
})

function docWithMarkedText(text: string, id: string): PmNode {
  const marked = schema.text(text, [schema.marks.comment.create({ ids: id })])
  return schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, marked))
}

function docPlain(text: string): PmNode {
  return schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text(text)))
}

/** minimal editor stub: real PM state, dispatch applies + fires listeners */
function stubEditor(initial: PmNode): {
  editor: Editor
  fire: () => void
  setState: (doc: PmNode) => void
  removeMark: () => void
  restoreMark: (text: string, id: string) => void
} {
  let state = EditorState.create({ doc: initial })
  const listeners: Array<(info: { transaction?: { docChanged: boolean } }) => void> = []
  const editor = {
    get state() {
      return state
    },
    set state(next) {
      state = next
    },
    on: (_ev: string, fn: (info: { transaction?: { docChanged: boolean } }) => void) => {
      listeners.push(fn)
    },
    off: (_ev: string, fn: (info: { transaction?: { docChanged: boolean } }) => void) => {
      const i = listeners.indexOf(fn)
      if (i >= 0) listeners.splice(i, 1)
    },
    view: {
      dispatch(tr: import('@tiptap/pm/state').Transaction) {
        state = state.apply(tr)
        for (const fn of [...listeners]) fn({ transaction: tr })
      },
    },
  } as unknown as Editor
  return {
    editor,
    fire: () => {
      for (const fn of [...listeners]) fn({ transaction: { docChanged: true } })
    },
    setState: (doc: PmNode) => {
      state = EditorState.create({ doc })
    },
    removeMark: () => {
      const tr = state.tr
      tr.removeMark(0, state.doc.content.size, schema.marks.comment)
      editor.view.dispatch(tr)
    },
    restoreMark: (text: string, id: string) => {
      const tr = state.tr
      tr.addMark(1, 1 + text.length, schema.marks.comment.create({ ids: id }))
      editor.view.dispatch(tr)
    },
  }
}

const entryOf = (id: string, text: string): CommentInfo => ({
  id,
  author: 'Reviewer',
  date: '2026-11-16T10:00:00Z',
  text,
})

describe('comment metadata undo-sync', () => {
  it('delete → undo restores the metadata card; redo removes it again', async () => {
    resetCommentUndoStore()
    const entry = entryOf('2', 'please reword this')
    const hook = stubEditor(docWithMarkedText('hello world', '2'))
    let store: CommentInfo[] = [entry]
    const applyMetadata = (entries: CommentInfo[]) => {
      store = entries
    }
    attachCommentUndoSync(hook.editor, () => store, applyMetadata)

    // delete: mark leaves the doc — metadata is parked and leaves the store
    hook.removeMark()
    expect(store).toEqual([])

    // undo: mark returns — the card must come back with author and text
    hook.setState(docWithMarkedText('hello world', '2'))
    hook.fire()
    await new Promise<void>((resolve) => queueMicrotask(() => resolve()))
    expect(
      store.some((c) => c.id === '2' && c.text === 'please reword this' && c.author === 'Reviewer'),
    ).toBe(true)

    // redo: mark leaves again — the card leaves with it
    hook.removeMark()
    expect(store).toEqual([])
  })

  it('undo of an add removes the card; undo of that undo brings it back', async () => {
    resetCommentUndoStore()
    const entry = entryOf('3', 'fresh note')
    const hook = stubEditor(docWithMarkedText('hello world', '3'))
    let store: CommentInfo[] = [entry]
    const applyMetadata = (entries: CommentInfo[]) => {
      store = entries
    }
    attachCommentUndoSync(hook.editor, () => store, applyMetadata)

    // undo the add: the mark disappears — the card must go too
    hook.setState(docPlain('hello world'))
    hook.fire()
    await new Promise<void>((resolve) => queueMicrotask(() => resolve()))
    expect(store).toEqual([])

    // undo the undo: the mark returns — the card returns
    hook.setState(docWithMarkedText('hello world', '3'))
    hook.fire()
    await new Promise<void>((resolve) => queueMicrotask(() => resolve()))
    expect(store.some((c) => c.id === '3' && c.text === 'fresh note')).toBe(true)
  })

  it('a forward re-add of a recycled id keeps the new card (no stale graveyard revive)', async () => {
    resetCommentUndoStore()
    const old = entryOf('2', 'old text')
    const hook = stubEditor(docWithMarkedText('hello world', '2'))
    let store: CommentInfo[] = [old]
    const applyMetadata = (entries: CommentInfo[]) => {
      store = entries
    }
    attachCommentUndoSync(hook.editor, () => store, applyMetadata)

    // delete, then re-add the same numeric id with different text — the real
    // flow (addComment) writes the new card in the same tick as the marks
    hook.removeMark()
    expect(store).toEqual([])
    hook.restoreMark('hello world', '2')
    store = [entryOf('2', 'hello world')]
    // the deferred revive must not resurrect the old card beside the new one
    await new Promise<void>((resolve) => queueMicrotask(() => resolve()))
    expect(store).toEqual([expect.objectContaining({ id: '2', text: 'hello world' })])
  })

  it('a document swap clears the graveyard (ids never cross documents)', () => {
    resetCommentUndoStore()
    const entry = entryOf('9', 'belongs to doc A')
    const hook = stubEditor(docWithMarkedText('doc A body', '9'))
    let store: CommentInfo[] = [entry]
    const applyMetadata = (entries: CommentInfo[]) => {
      store = entries
    }
    attachCommentUndoSync(hook.editor, () => store, applyMetadata)

    // doc A's comment is deleted, then doc B opens (the cross-doc reset runs)
    hook.removeMark()
    resetCommentUndoStore()
    hook.setState(docWithMarkedText('doc B body', '9'))
    hook.fire()
    // doc B's own id-9 comment must not be mistaken for doc A's deleted one
    expect(store.some((c) => c.text === 'belongs to doc A')).toBe(false)
  })

  it('loading a second document after a delete clears the graveyard through the cross-doc reset', async () => {
    resetCommentUndoStore()
    const entry = entryOf('4', 'doc A card')
    const hook = stubEditor(docWithMarkedText('doc A body', '4'))
    let store: CommentInfo[] = [entry]
    const applyMetadata = (entries: CommentInfo[]) => {
      store = entries
    }
    attachCommentUndoSync(hook.editor, () => store, applyMetadata)

    // doc A: the comment is deleted (metadata parked)
    hook.removeMark()
    expect(store).toEqual([])

    // doc B loads in the same window: loadFile / the save reload run the
    // cross-doc reset, which now also clears the graveyard — a doc B comment
    // reusing anchor id 4 must not revive doc A's parked card
    resetCrossDocEditState({ ...cleanSink(), resetCommentUndoStore })
    hook.setState(docWithMarkedText('doc B body', '4'))
    hook.fire()
    await new Promise<void>((resolve) => queueMicrotask(() => resolve()))
    expect(store).toEqual([])

    // doc A's card never comes back
    expect(store.some((c) => c.text === 'doc A card')).toBe(false)
  })
})

import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const paraWithBookmarks = (text: string, bookmarks: string[] | null) => ({
  type: 'docParagraph' as const,
  attrs: { docxIndex: null, bookmarks },
  content: text ? [{ type: 'text' as const, text }] : [],
})

function open(paragraphs: unknown[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: paragraphs } as never,
  })
  editors.push(editor)
  return editor
}

describe('BookmarkTargets widgets', () => {
  it('renders one zero-size id anchor per bookmark name at its block', () => {
    const editor = open([
      paraWithBookmarks('第一章', ['_Toc1', '_Ref2']),
      paraWithBookmarks('第二章', ['_Toc3']),
      paraWithBookmarks('无书签', null),
    ])
    const dom = editor.view.dom
    const anchors = [...dom.querySelectorAll('a.doc-bookmark-target')]
    expect(anchors.map((a) => a.id)).toEqual(['_Toc1', '_Ref2', '_Toc3'])
    // the anchors carry no content: copying, caret geometry and the save
    // patch never see them
    expect(anchors.every((a) => a.textContent === '')).toBe(true)
    expect(anchors[0].closest('p')).not.toBeNull()
  })

  it('keeps the anchors in step with edits that move the bookmarked block', () => {
    const editor = open([paraWithBookmarks('目标段', ['_Toc9']), paraWithBookmarks('占位', null)])
    // typing into the block below must not disturb the anchor above
    editor.commands.insertContentAt(editor.state.doc.child(1).nodeSize + 1, 'x')
    const ids = [...editor.view.dom.querySelectorAll('a.doc-bookmark-target')].map((a) => a.id)
    expect(ids).toEqual(['_Toc9'])
    // deleting the bookmarked block removes its anchor (the editor keeps an
    // empty paragraph alive, which carries no bookmark names)
    editor.commands.deleteRange({ from: 0, to: editor.state.doc.child(0).nodeSize })
    expect([...editor.view.dom.querySelectorAll('a.doc-bookmark-target')]).toEqual([])
  })

  it('does not touch the document model (save/copy stay bookmark-free)', () => {
    const editor = open([paraWithBookmarks('正文', ['_Toc1'])])
    expect(editor.state.doc.textContent).toBe('正文')
    const paragraph = editor.state.doc.child(0)
    expect(JSON.stringify(paragraph.toJSON())).not.toContain('doc-bookmark-target')
  })
})

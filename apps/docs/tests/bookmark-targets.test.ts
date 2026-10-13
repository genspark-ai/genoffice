import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string, attrs: Record<string, unknown>) => ({
  type: 'docParagraph' as const,
  attrs: { docxIndex: null, bookmarks: null, hiddenBookmarks: null, ...attrs },
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

const anchorIds = (editor: Editor): string[] =>
  [...editor.view.dom.querySelectorAll<HTMLAnchorElement>('a.doc-bookmark-target')].map((a) => a.id)

const bookmarkNames = (editor: Editor): string[] =>
  [...editor.view.dom.querySelectorAll<HTMLAnchorElement>('a.doc-bookmark-target')].map(
    (a) => a.dataset.bookmark ?? '',
  )

describe('BookmarkTargets widgets', () => {
  it('anchors user bookmarks', () => {
    const editor = open([
      para('Chapter one', { bookmarks: ['chap01'] }),
      para('Chapter two', { bookmarks: ['chap02'] }),
      para('Plain paragraph', {}),
    ])
    expect(anchorIds(editor)).toEqual(['chap01', 'chap02'])
  })

  it('anchors Word internal names too, or a TOC document gets no anchors at all', () => {
    // the parser splits _Toc…/_Ref… into hiddenBookmarks so the bookmark
    // manager ignores them; a generated TOC links to exactly those
    const editor = open([
      para('Contents', { hiddenBookmarks: ['_Toc271234567'] }),
      para('Heading one', { hiddenBookmarks: ['_Toc900000001', '_Ref100000001'] }),
      para('Heading two', { bookmarks: ['chap02'], hiddenBookmarks: ['_Toc900000002'] }),
    ])
    expect(anchorIds(editor)).toEqual([
      '_Toc271234567',
      '_Toc900000001',
      '_Ref100000001',
      'chap02',
      '_Toc900000002',
    ])
  })

  it('keeps the anchors in step with edits that move the bookmarked block', () => {
    const editor = open([para('Target block', { bookmarks: ['chap01'] }), para('Filler', {})])
    // typing into the block below must not disturb the anchor above
    editor.commands.insertContentAt(editor.state.doc.child(1).nodeSize + 1, 'x')
    expect(anchorIds(editor)).toEqual(['chap01'])
    // deleting the bookmarked block removes its anchor (the editor keeps an
    // empty paragraph alive, which carries no bookmark names)
    editor.commands.deleteRange({ from: 0, to: editor.state.doc.child(0).nodeSize })
    expect([...editor.view.dom.querySelectorAll('a.doc-bookmark-target')]).toEqual([])
  })

  it('does not touch the document model (save/copy stay bookmark-free)', () => {
    const editor = open([para('Body text', { bookmarks: ['chap01'], hiddenBookmarks: ['_Toc1'] })])
    expect(editor.state.doc.textContent).toBe('Body text')
    expect(JSON.stringify(editor.state.doc.child(0).toJSON())).not.toContain('doc-bookmark-target')
  })

  it('carries the name in data-bookmark so export can move the id onto one page clone', () => {
    const editor = open([para('Heading', { hiddenBookmarks: ['_Toc1'] })])
    expect(bookmarkNames(editor)).toEqual(['_Toc1'])
  })

  it('renders anchors with no content, so no line box gains ink', () => {
    const editor = open([para('Heading', { bookmarks: ['chap01'] })])
    const anchors = [...editor.view.dom.querySelectorAll('a.doc-bookmark-target')]
    expect(anchors.every((a) => a.textContent === '')).toBe(true)
    expect(anchors[0].closest('p')).not.toBeNull()
  })
})

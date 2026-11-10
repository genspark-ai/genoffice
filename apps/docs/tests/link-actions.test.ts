/**
 * Word's hyperlink context menu acts on the link run under the pointer:
 * the run is found from either side of its boundary, Remove Hyperlink keeps
 * the text and the other marks, and Open resolves http(s) vs #bookmark.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { linkRangeAt, linkTarget, removeLink } from '../src/renderer/editor/link-actions'

const live: Editor[] = []
afterEach(() => {
  for (const e of live.splice(0)) e.destroy()
})

function makeEditor() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'visit ' },
            {
              type: 'text',
              text: 'our ',
              marks: [{ type: 'link', attrs: { href: 'https://example.com' } }],
            },
            {
              type: 'text',
              text: 'site',
              marks: [{ type: 'bold' }, { type: 'link', attrs: { href: 'https://example.com' } }],
            },
            { type: 'text', text: ' today' },
          ],
        },
      ],
    },
  })
  live.push(editor)
  return editor
}

describe('linkRangeAt', () => {
  it('spans every run of one link, from inside and from both edges', () => {
    const editor = makeEditor()
    // "visit " = 1..7, link "our site" = 7..15
    for (const pos of [7, 9, 12, 15]) {
      const range = linkRangeAt(editor.state, pos)
      expect(range, `pos ${pos}`).toMatchObject({ from: 7, to: 15, href: 'https://example.com' })
    }
    expect(linkRangeAt(editor.state, 3)).toBeNull()
    expect(linkRangeAt(editor.state, 17)).toBeNull()
  })
})

describe('removeLink', () => {
  it('keeps the text and the other marks, leaves the selection alone', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(3)
    const range = linkRangeAt(editor.state, 9)!
    expect(removeLink(editor, range)).toBe(true)
    expect(editor.state.doc.textContent).toBe('visit our site today')
    expect(linkRangeAt(editor.state, 9)).toBeNull()
    const site = editor.state.doc.nodeAt(11)!
    expect(site.marks.map((m) => m.type.name)).toEqual(['bold'])
    expect(editor.state.selection.from).toBe(3)
  })
})

describe('linkTarget', () => {
  it('classifies browser links, bookmarks and the rest', () => {
    expect(linkTarget('https://a.b/c')).toEqual({ kind: 'external', url: 'https://a.b/c' })
    expect(linkTarget(' HTTP://a.b ')).toEqual({ kind: 'external', url: 'HTTP://a.b' })
    expect(linkTarget('www.example.com/x')).toEqual({
      kind: 'external',
      url: 'https://www.example.com/x',
    })
    expect(linkTarget('#_Toc123')).toEqual({ kind: 'anchor', name: '_Toc123' })
    expect(linkTarget('#')).toBeNull()
    expect(linkTarget('javascript:alert(1)')).toBeNull()
    expect(linkTarget('file:///etc/passwd')).toBeNull()
    expect(linkTarget('')).toBeNull()
  })
})

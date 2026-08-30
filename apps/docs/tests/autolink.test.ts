/**
 * Word AutoFormat parity (alpha ledger r151): a URL followed by space or
 * Enter becomes a hyperlink; trailing punctuation stays outside; existing
 * links and non-URLs are untouched.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const makeEditor = (text: string) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
    },
  })

const pressAtEnd = (editor: Editor, key: string) => {
  editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
  editor.view.dom.dispatchEvent(event)
}

const linkOf = (editor: Editor) => {
  const found: { href: string; from: number; to: number }[] = []
  editor.state.doc.descendants((node, pos) => {
    const mark = node.marks.find((m) => m.type.name === 'link')
    if (mark) found.push({ href: mark.attrs.href as string, from: pos, to: pos + node.nodeSize })
  })
  return found.at(-1) ?? null
}

describe('auto-link on delimiter (r151)', () => {
  it('linkifies a pasted/typed https URL on Enter', () => {
    const editor = makeEditor('see https://example.com/page')
    pressAtEnd(editor, 'Enter')
    expect(linkOf(editor)?.href).toBe('https://example.com/page')
    editor.destroy()
  })

  it('linkifies on space and keeps trailing punctuation outside', () => {
    const editor = makeEditor('see https://example.com/page.')
    pressAtEnd(editor, ' ')
    const link = linkOf(editor)
    expect(link?.href).toBe('https://example.com/page')
    editor.destroy()
  })

  it('prefixes www. links with http://', () => {
    const editor = makeEditor('go www.example.com')
    pressAtEnd(editor, 'Enter')
    expect(linkOf(editor)?.href).toBe('http://www.example.com')
    editor.destroy()
  })

  it('does not swallow an inline atom (hard break) into the URL', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'see https://example.com' },
              { type: 'hardBreak' },
              { type: 'text', text: 'tail' },
            ],
          },
        ],
      } as never,
    })
    pressAtEnd(editor, ' ')
    expect(linkOf(editor)).toBeNull()
    editor.destroy()
  })

  it('leaves plain text and existing links alone', () => {
    const plain = makeEditor('nothing to see here')
    pressAtEnd(plain, 'Enter')
    expect(linkOf(plain)).toBeNull()
    plain.destroy()

    const linked = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'https://kept.example.com',
                marks: [{ type: 'link', attrs: { href: 'https://original.example.com' } }],
              },
            ],
          },
        ],
      } as never,
    })
    pressAtEnd(linked, ' ')
    expect(linkOf(linked)?.href).toBe('https://original.example.com')
    linked.destroy()
  })
})

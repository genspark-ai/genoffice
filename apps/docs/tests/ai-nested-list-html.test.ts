import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { serializeRangeToHtml } from '../src/renderer/ai/protocol'

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

const item = (text: string, ilvl: number, kind = 'bullet') => ({
  type: 'docListItem',
  attrs: { kind, ilvl },
  content: [{ type: 'text', text }],
})

const html = (content: ReturnType<typeof item>[]): string => {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return serializeRangeToHtml(editor, 0, content.length - 1)
}

describe('serializeRangeToHtml lists', () => {
  it('nests a list that follows one of another kind opened at a deeper level', () => {
    expect(html([item('b', 1), item('x', 0, 'ordered'), item('x1', 1, 'ordered')])).toBe(
      '<ul><li>b</li></ul>\n<ol><li>x<ol><li>x1</li></ol></li></ol>',
    )
  })

  it('nests under a shallower item when the range starts mid-list', () => {
    expect(html([item('b1', 1), item('c', 0), item('c1', 1)])).toBe(
      '<ul><li>b1</li><li>c<ul><li>c1</li></ul></li></ul>',
    )
  })

  it('nests list items by their Word level', () => {
    const content = [
      item('a', 0),
      item('b', 0),
      item('b1', 1),
      item('b1x', 2, 'ordered'),
      item('b2', 1),
      item('c', 0),
    ]
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content },
    })
    editors.add(editor)

    expect(serializeRangeToHtml(editor, 0, content.length - 1)).toBe(
      '<ul><li>a</li><li>b<ul><li>b1<ol><li>b1x</li></ol></li><li>b2</li></ul></li><li>c</li></ul>',
    )
  })
})

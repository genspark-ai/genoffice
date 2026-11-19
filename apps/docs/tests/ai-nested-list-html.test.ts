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

describe('serializeRangeToHtml lists', () => {
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

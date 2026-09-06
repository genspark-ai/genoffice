import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

function paraStyle(attrs: Record<string, unknown>): CSSStyleDeclaration {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null, ...attrs },
          content: [{ type: 'text', text: 'x' }],
        },
      ],
    },
  })
  const style = editor.view.dom.querySelector('p')!.style
  editor.destroy()
  return style
}

describe('direct w:pBdr none over a style border', () => {
  it('emits an inline border-none for the reset sides only', () => {
    const style = paraStyle({ styleId: 'HDR', borderReset: 'b' })
    expect(style.borderBottomStyle).toBe('none')
    expect(style.borderTopStyle).toBe('')
    expect(style.padding).toBe('')
  })

  it('keeps drawn sides next to reset ones', () => {
    const style = paraStyle({ borders: 't', borderReset: 'b' })
    expect(style.borderTopStyle).toBe('solid')
    expect(style.borderTopWidth).toBe('1px')
    expect(style.borderBottomStyle).toBe('none')
    expect(style.padding).toBe('1px 4px')
  })
})

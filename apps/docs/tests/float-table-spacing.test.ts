import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { floatTableSpacingDecorations } from '../src/renderer/editor/float-table-spacing'

const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
const cell = { type: 'docTableCell', content: [{ type: 'docParagraph' }] }
const table = (attrs: Record<string, unknown>) => ({
  type: 'docTable',
  attrs,
  content: [{ type: 'docTableRow', content: [cell] }],
})
const para = (attrs: Record<string, unknown>, text = 'x') => ({
  type: 'docParagraph',
  attrs,
  content: [{ type: 'text', text }],
})

function decorate(content: unknown[]) {
  const doc = editor.schema.nodeFromJSON({ type: 'doc', content })
  return floatTableSpacingDecorations(doc)
    .find()
    .map((d) => (d as unknown as { type: { attrs: Record<string, string> } }).type.attrs)
}

describe('paragraph spacing across a floating table', () => {
  it('the paragraph after a floated table carries its before plus the previous after', () => {
    const decos = decorate([
      para({ spaceAfter: 180 }),
      table({ tblFloat: 'left', tblFloatSource: 'left' }),
      para({ spaceBefore: 180 }),
    ])
    expect(decos).toEqual([{ class: 'doc-after-tblp', style: '--tblp-sp-before:18.0pt' }])
  })

  it('auto spacing counts as 14pt on either side', () => {
    const decos = decorate([
      para({ spaceAfterAuto: true }),
      table({ tblFloat: 'right', tblFloatSource: 'right' }),
      para({ spaceBefore: 120 }),
    ])
    expect(decos[0]?.style).toBe('--tblp-sp-before:20.0pt')
  })

  it('nothing to add without a direct previous space-after, an inline table, or a suppressed float', () => {
    expect(
      decorate([
        para({}),
        table({ tblFloat: 'left', tblFloatSource: 'left' }),
        para({ spaceBefore: 180 }),
      ]),
    ).toEqual([])
    expect(decorate([para({ spaceAfter: 180 }), table({}), para({ spaceBefore: 180 })])).toEqual([])
    expect(
      decorate([
        para({ spaceAfter: 180 }),
        table({ tblFloat: null, tblFloatSource: 'left', tblFloatSuppressed: true }),
        para({ spaceBefore: 180 }),
      ]),
    ).toEqual([])
  })

  it('the editor renders the class and the sum variable on the anchor paragraph', () => {
    editor.commands.setContent({
      type: 'doc',
      content: [
        para({ spaceAfter: 180 }),
        table({ tblFloat: 'left', tblFloatSource: 'left' }),
        para({ spaceBefore: 180 }, 'anchor'),
      ],
    })
    const p = editor.view.dom.querySelector('.doc-after-tblp') as HTMLElement | null
    expect(p?.textContent).toBe('anchor')
    expect(p?.style.getPropertyValue('--tblp-sp-before')).toBe('18.0pt')
    expect(editor.view.dom.querySelector('.doc-table-float-left + .doc-after-tblp')).toBe(p)
  })
})

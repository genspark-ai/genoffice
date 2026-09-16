import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { mergePPrFormat } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { paraBorderPadding } from '../src/renderer/editor/hf-dom'

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

function listItemStyle(attrs: Record<string, unknown>): CSSStyleDeclaration {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docListItem',
          attrs: { docxIndex: null, kind: 'bullet', ilvl: 0, ...attrs },
          content: [{ type: 'text', text: 'x' }],
        },
      ],
    },
  })
  const style = editor.view.dom.querySelector<HTMLElement>('.doc-li')!.style
  editor.destroy()
  return style
}

describe('direct w:pBdr none over a style border', () => {
  it('emits an inline border-none for the reset sides only', () => {
    const style = paraStyle({ styleId: 'HDR', borderReset: 'b' })
    expect(style.borderBottomStyle).toBe('none')
    expect(style.borderTopStyle).toBe('')
    expect(style.paddingBottom).toBe('0px')
    expect(style.paddingTop).toBe('')
  })

  it('keeps drawn sides next to reset ones', () => {
    const style = paraStyle({ borders: 't', borderReset: 'b' })
    expect(style.borderTopStyle).toBe('solid')
    expect(style.borderTopWidth).toBe('1px')
    expect(style.borderBottomStyle).toBe('none')
    expect(style.paddingTop).toBe('0px')
    // the reset side also drops the style-level w:space gap
    expect(style.paddingBottom).toBe('0px')
    expect(style.paddingLeft).toBe('')
  })

  it('a reset left/right side drops the style inset on a paragraph', () => {
    const style = paraStyle({ borderReset: 'lr' })
    expect(style.borderLeftStyle).toBe('none')
    expect(style.paddingLeft).toBe('0px')
    expect(style.paddingRight).toBe('0px')
  })

  it('a list item keeps its indent padding when a left/right side is reset', () => {
    const style = listItemStyle({ borderReset: 'lr' })
    expect(style.borderLeftStyle).toBe('none')
    expect(style.paddingLeft).toBe('')
    expect(style.paddingRight).toBe('')
  })

  it('pads drawn top/bottom sides by their w:space, left/right by the text inset', () => {
    const style = paraStyle({
      borders: 'bl',
      borderLines: JSON.stringify({ b: { color: '4F81BD', szPt: 1, spacePt: 4 } }),
    })
    expect(style.paddingTop).toBe('')
    expect(style.paddingBottom).toBe('4pt')
    expect(style.paddingLeft).toBe('4px')
    expect(style.paddingRight).toBe('')
  })

  it('a side with no declared w:space displays the same gap the save writes', () => {
    const xml = mergePPrFormat('<w:pPr></w:pPr>', { borders: 'tb' })
    const saved = [...xml.matchAll(/w:space="(\d+)"/g)].map((m) => Number(m[1]))
    expect(saved).toEqual([0, 0])
    const padding = paraBorderPadding('tb')
    expect(parseFloat(padding.paddingTop!)).toBe(saved[0])
    expect(parseFloat(padding.paddingBottom!)).toBe(saved[1])
    expect(paraStyle({ borders: 'tb' }).paddingBottom).toBe('0px')
  })
})

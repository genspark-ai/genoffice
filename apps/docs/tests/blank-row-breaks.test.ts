import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blankRowBreaks } from '../src/renderer/text-layout/blank-row-breaks'

const face = {
  type: 'docTextStyle',
  attrs: { font: 'Cambria', fontAscii: 'Cambria', sizeHalfPoints: 24 },
}
const text = (t: string) => ({ type: 'text', text: t, marks: [face] })
const br = { type: 'hardBreak' }

describe('blankRowBreaks', () => {
  it('lists the soft breaks that open a text-less row by how their run sizes itself', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { ...br, marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 22 } }] },
              text('abc'),
              br,
              br,
              br,
              text('def'),
              { ...br, marks: [{ type: 'docTextStyle', attrs: { font: 'Arial' } }] },
              { ...br, marks: [{ type: 'docTextStyle', attrs: { font: 'Arial' } }] },
            ],
          },
        ],
      } as never,
    })
    const para = editor.state.doc.firstChild!
    // offsets: sized break 0; 'abc' 1..4, breaks at 4 (closes abc), 5, 6; 'def' 7..10, Arial breaks 10, 11
    expect(blankRowBreaks(para)).toEqual({ sized: [0], inherited: [5, 6], leadingEnd: 1 })
    const rows = editor.view.dom.querySelectorAll('.doc-br-line')
    expect(rows.length).toBe(3)
    expect(editor.view.dom.querySelectorAll('.doc-br-line.doc-br-inherit').length).toBe(2)
    editor.destroy()
  })

  it('keeps the paragraph strut on blank rows when the runs inherit the face too', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'abc' }, br, br, { type: 'text', text: 'def' }],
          },
        ],
      } as never,
    })
    expect(blankRowBreaks(editor.state.doc.firstChild!)).toEqual({
      sized: [],
      inherited: [4],
      leadingEnd: 0,
    })
    expect(editor.view.dom.querySelectorAll('.doc-br-line').length).toBe(0)
    editor.destroy()
  })

  it('keeps mid-paragraph rows inline when the first line is indented', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { indentFirstLine: 720 },
            content: [
              { ...br, marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 22 } }] },
              text('abc'),
              br,
              br,
              text('def'),
            ],
          },
        ],
      } as never,
    })
    expect(editor.view.dom.querySelectorAll('.doc-br-line').length).toBe(1)
    expect(editor.view.dom.querySelectorAll('.doc-br-inherit').length).toBe(0)
    editor.destroy()
  })

  it('a page or column break ends its row: the soft break after it opens a blank row', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              text('abc'),
              { type: 'hardBreak', attrs: { pageBreak: true } },
              br,
              text('def'),
              { type: 'hardBreak', attrs: { colBreak: true } },
              { ...br, marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 22 } }] },
            ],
          },
        ],
      } as never,
    })
    const para = editor.state.doc.firstChild!
    // 'abc' 0..3, page break 3, soft 4; 'def' 5..8, column break 8, sized soft 9
    expect(blankRowBreaks(para)).toEqual({ sized: [9], inherited: [4], leadingEnd: 0 })
    expect(editor.view.dom.querySelectorAll('.doc-br-line').length).toBe(2)
    editor.destroy()
  })

  it('a leading page break ends the leading breaks: the blank row after it stays inline under an indent', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { indentFirstLine: 720 },
            content: [{ type: 'hardBreak', attrs: { pageBreak: true } }, br, text('abc')],
          },
        ],
      } as never,
    })
    const para = editor.state.doc.firstChild!
    expect(blankRowBreaks(para)).toEqual({ sized: [], inherited: [1], leadingEnd: 0 })
    expect(editor.view.dom.querySelectorAll('.doc-br-line').length).toBe(0)
    editor.destroy()
  })

  it('Normal / docDefaults first-line indent reaches body paragraphs only', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [text('abc'), br, br, text('def')] },
          { type: 'docHeading', attrs: { level: 1 }, content: [text('abc'), br, br, text('def')] },
          {
            type: 'docListItem',
            attrs: { indentLeft: 720 },
            content: [text('abc'), br, br, text('def')],
          },
        ],
      } as never,
    })
    editor.storage.listNumbering.docDefaults = { indentFirstLineTwips: 420 }
    // the style table is read when a block is (re)decorated
    let tr = editor.state.tr
    editor.state.doc.forEach((_n, pos) => {
      tr = tr.insertText('x', tr.mapping.map(pos + 1))
    })
    editor.view.dispatch(tr)
    const rows = editor.view.dom.querySelectorAll('.doc-br-inherit')
    expect(rows.length).toBe(2)
    expect(rows[0].closest('h1')).not.toBeNull()
    expect(rows[1].closest('.doc-li')).not.toBeNull()
    editor.destroy()
  })

  it('a list item hanging indent is no first-line indent: its blank rows keep the wrappers', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { indentLeft: 720, indentFirstLine: -360 },
            content: [text('abc'), br, br, text('def')],
          },
          {
            type: 'docListItem',
            attrs: { indentLeft: 720, indentFirstLine: 360 },
            content: [text('abc'), br, br, text('def')],
          },
        ],
      } as never,
    })
    expect(editor.view.dom.querySelectorAll('.doc-br-inherit').length).toBe(1)
    editor.destroy()
  })

  it('an explicit zero first-line indent keeps the wrappers', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { indentFirstLine: 0 },
            content: [text('abc'), br, br, text('def')],
          },
        ],
      } as never,
    })
    expect(editor.view.dom.querySelectorAll('.doc-br-inherit').length).toBe(1)
    editor.destroy()
  })
})

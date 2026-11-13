import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectStructuralHeadings } from '../src/renderer/editor/structural-headings'
import type { HeadingStyles } from '../src/renderer/editor/headings'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (value: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: value,
  ...(marks ? { marks } : {}),
})

const bold = (value: string): JsonNode => text(value, [{ type: 'bold' }])

const caps = (value: string): JsonNode =>
  text(value, [{ type: 'docTextStyle', attrs: { caps: 'all' } }])

const para = (content: JsonNode[] = []): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

const heading = (value: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(value)],
})

const listItem = (value: string): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, kind: 'bullet', numId: '1', ilvl: 0 },
  content: [text(value)],
})

const table = (cellText: string): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: null },
  content: [
    {
      type: 'docTableRow',
      content: [
        { type: 'docTableCell', content: [para([bold(cellText)])] },
        { type: 'docTableCell', content: [para([text('value')])] },
      ],
    },
  ],
})

const editors: Editor[] = []

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

const infer = (content: JsonNode[], styles?: HeadingStyles): Array<Record<string, unknown>> => {
  const editor = createEditor(content)
  return collectStructuralHeadings(editor.state.doc, styles).map(({ text, level, basis }) => ({
    text,
    level,
    basis,
  }))
}

describe('collectStructuralHeadings: positive shapes', () => {
  it('reads a short all-bold line followed by body text as a heading', () => {
    expect(
      infer([para([bold('RENEWAL TERMS')]), para([text('The parties agree as follows.')])]),
    ).toEqual([{ text: 'RENEWAL TERMS', level: 1, basis: 'bold' }])
  })

  it('reads a short all-capitals line followed by body text as a heading', () => {
    expect(
      infer([para([text('BACKGROUND')]), para([text('The parties agree as follows.')])]),
    ).toEqual([{ text: 'BACKGROUND', level: 1, basis: 'caps' }])
  })

  it("reads a run formatted with Word's all-caps setting as capitals", () => {
    expect(
      infer([para([caps('Renewal Terms')]), para([text('The parties agree as follows.')])]),
    ).toEqual([{ text: 'Renewal Terms', level: 1, basis: 'caps' }])
  })

  it('skips blank paragraphs between the heading and the body', () => {
    expect(
      infer([para([bold('RENEWAL TERMS')]), para(), para([text('The parties agree as follows.')])]),
    ).toEqual([{ text: 'RENEWAL TERMS', level: 1, basis: 'bold' }])
  })

  it('puts Roman numerals above Arabic numbers and letters', () => {
    expect(
      infer([
        para([text('I. Background')]),
        para([text('First body paragraph.')]),
        para([text('1. Scope')]),
        para([text('Second body paragraph.')]),
        para([text('A. Definitions')]),
        para([text('Third body paragraph.')]),
      ]),
    ).toEqual([
      { text: 'I. Background', level: 1, basis: 'number' },
      { text: '1. Scope', level: 2, basis: 'number' },
      { text: 'A. Definitions', level: 3, basis: 'number' },
    ])
  })

  it('drops a numbering kind the accepted headings never use', () => {
    expect(
      infer([
        para([text('1. Scope')]),
        para([text('First body paragraph.')]),
        para([text('A. Definitions')]),
        para([text('Second body paragraph.')]),
      ]),
    ).toEqual([
      { text: '1. Scope', level: 1, basis: 'number' },
      { text: 'A. Definitions', level: 2, basis: 'number' },
    ])
  })
})

describe('collectStructuralHeadings: negative shapes', () => {
  it('leaves an unmarked short line alone', () => {
    expect(
      infer([para([text('Renewal terms')]), para([text('The parties agree as follows.')])]),
    ).toEqual([])
  })

  it('leaves a long bold paragraph alone', () => {
    const long =
      'The parties agree that this agreement, together with every schedule attached to it, ' +
      'is governed by the laws of the place where the company is registered.'
    expect(infer([para([bold(long)]), para([text('The parties agree as follows.')])])).toEqual([])
  })

  it('leaves a bold sentence that ends with a full stop alone', () => {
    expect(
      infer([para([bold('The parties agree as follows.')]), para([text('Second paragraph.')])]),
    ).toEqual([])
  })

  it('leaves an all-caps line that ends with a question mark alone', () => {
    expect(
      infer([para([text('WHAT IS THE RENEWAL TERM?')]), para([text('Second paragraph.')])]),
    ).toEqual([])
  })

  it('does not treat a numbered sentence as a heading', () => {
    expect(
      infer([
        para([text('1. The tenant shall pay the rent on the first day of each month.')]),
        para([text('Second paragraph.')]),
      ]),
    ).toEqual([])
  })

  it('does not treat a bold line followed by a table as a heading', () => {
    expect(infer([para([bold('SCHEDULE A')]), table('Amount')])).toEqual([])
  })

  it('does not treat a bold line followed by another bold line as a heading', () => {
    expect(
      infer([para([bold('SCHEDULE A')]), para([bold('SCHEDULE B')]), table('Amount')]),
    ).toEqual([])
  })

  it('does not reach into tables', () => {
    expect(infer([table('HEADING')])).toEqual([])
  })

  it('does not treat bold list items as headings', () => {
    expect(infer([listItem('TERM'), para([text('The parties agree as follows.')])])).toEqual([])
  })

  it('does not treat a trailing bold line without a following paragraph as a heading', () => {
    expect(infer([para([bold('SCHEDULE A')])])).toEqual([])
  })

  it('prefers real heading styles when the document has them', () => {
    expect(
      infer([
        heading('Background'),
        para([bold('RENEWAL TERMS')]),
        para([text('The parties agree as follows.')]),
      ]),
    ).toEqual([])
  })

  it('prefers style-driven headings when the document has them', () => {
    const styles: HeadingStyles = new Map([['CustomHeading', { headingLevel: 2 }]])
    expect(
      infer(
        [
          {
            type: 'docParagraph',
            attrs: { docxIndex: null, styleId: 'CustomHeading' },
            content: [text('Background')],
          },
          para([bold('RENEWAL TERMS')]),
          para([text('The parties agree as follows.')]),
        ],
        styles,
      ),
    ).toEqual([])
  })
})

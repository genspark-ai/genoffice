import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectHeadings, type HeadingStyles } from '../src/renderer/editor/headings'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })

const heading = (t: string, level: number): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(t)],
})

/** a numbered heading: the paragraph is a list item whose style carries the outline level */
const numberedHeading = (t: string, styleId: string): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, styleId, kind: 'ordered', numId: '1', ilvl: 0 },
  content: [text(t)],
})

const listItem = (t: string, styleId?: string): JsonNode => ({
  type: 'docListItem',
  attrs: { docxIndex: null, ...(styleId ? { styleId } : {}), kind: 'bullet', numId: '2', ilvl: 0 },
  content: [text(t)],
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

function styleWith(fields: Partial<StyleInfo>): StyleInfo {
  return { styleId: 'x', name: 'x', type: 'paragraph', ...fields } as StyleInfo
}

const STYLES: HeadingStyles = new Map<string, StyleInfo>([
  ['2', styleWith({ styleId: '2', name: 'heading 1', headingLevel: 1 })],
  ['3', styleWith({ styleId: '3', name: 'heading 2', headingLevel: 2 })],
  ['40', styleWith({ styleId: '40', name: 'List Paragraph' })],
  ['54', styleWith({ styleId: '54', name: 'body', headingOutlineOff: true, headingLevel: 1 })],
])

describe('collectHeadings', () => {
  it('collects docHeading nodes without a style lookup', () => {
    const editor = createEditor([
      heading('\u65b9\u6848\u4fee\u8ba2\u53f2', 1),
      heading('\u9644\u5f551', 2),
    ])
    const refs = collectHeadings(editor.state.doc)
    expect(refs.map(({ level, text }) => ({ level, text }))).toEqual([
      { level: 1, text: '\u65b9\u6848\u4fee\u8ba2\u53f2' },
      { level: 2, text: '\u9644\u5f551' },
    ])
    expect(refs[0]!.pos).toBe(0)
    expect(refs[1]!.pos).toBeGreaterThan(refs[0]!.pos)
    editor.destroy()
  })

  it('collects numbered headings (style-driven outline level) in document order', () => {
    const editor = createEditor([
      heading('\u76ee\u5f55', 1),
      numberedHeading('\u80cc\u666f\u4ecb\u7ecd', '2'),
      numberedHeading('\u7814\u7a76\u7406\u8bba\u4f9d\u636e', '3'),
      listItem('\u80ba\u529f\u80fd\u6d4b\u5b9a', '40'),
    ])
    const refs = collectHeadings(editor.state.doc, STYLES)
    expect(refs.map(({ level, text }) => ({ level, text }))).toEqual([
      { level: 1, text: '\u76ee\u5f55' },
      { level: 1, text: '\u80cc\u666f\u4ecb\u7ecd' },
      { level: 2, text: '\u7814\u7a76\u7406\u8bba\u4f9d\u636e' },
    ])
    // document order, no sorting
    expect(refs.map((r) => r.pos)).toEqual([...refs.map((r) => r.pos)].sort((a, b) => a - b))
    editor.destroy()
  })

  it('ignores nodes without a style lookup, without a heading style, or with outlineLvl 9', () => {
    const editor = createEditor([
      numberedHeading('\u80cc\u666f\u4ecb\u7ecd', '2'),
      listItem('\u80ba\u529f\u80fd\u6d4b\u5b9a'),
      listItem('\u666e\u901a\u5217\u8868', '40'),
      numberedHeading('\u6b63\u6587\u4f2a\u88c5', '54'),
    ])
    expect(collectHeadings(editor.state.doc)).toEqual([])
    expect(
      collectHeadings(editor.state.doc, STYLES).map(({ level, text }) => ({ level, text })),
    ).toEqual([{ level: 1, text: '\u80cc\u666f\u4ecb\u7ecd' }])
    editor.destroy()
  })

  it('skips empty headings', () => {
    const editor = createEditor([heading('', 1), numberedHeading('  ', '2')])
    expect(collectHeadings(editor.state.doc, STYLES)).toEqual([])
    editor.destroy()
  })
})

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
    const editor = createEditor([heading('方案修订史', 1), heading('附录1', 2)])
    const refs = collectHeadings(editor.state.doc)
    expect(refs.map(({ level, text }) => ({ level, text }))).toEqual([
      { level: 1, text: '方案修订史' },
      { level: 2, text: '附录1' },
    ])
    expect(refs[0]!.pos).toBe(0)
    expect(refs[1]!.pos).toBeGreaterThan(refs[0]!.pos)
    editor.destroy()
  })

  it('collects numbered headings (style-driven outline level) in document order', () => {
    const editor = createEditor([
      heading('目录', 1),
      numberedHeading('背景介绍', '2'),
      numberedHeading('研究理论依据', '3'),
      listItem('肺功能测定', '40'),
    ])
    const refs = collectHeadings(editor.state.doc, STYLES)
    expect(refs.map(({ level, text }) => ({ level, text }))).toEqual([
      { level: 1, text: '目录' },
      { level: 1, text: '背景介绍' },
      { level: 2, text: '研究理论依据' },
    ])
    // document order, no sorting
    expect(refs.map((r) => r.pos)).toEqual([...refs.map((r) => r.pos)].sort((a, b) => a - b))
    editor.destroy()
  })

  it('ignores nodes without a style lookup, without a heading style, or with outlineLvl 9', () => {
    const editor = createEditor([
      numberedHeading('背景介绍', '2'),
      listItem('肺功能测定'),
      listItem('普通列表', '40'),
      numberedHeading('正文伪装', '54'),
    ])
    expect(collectHeadings(editor.state.doc)).toEqual([])
    expect(
      collectHeadings(editor.state.doc, STYLES).map(({ level, text }) => ({ level, text })),
    ).toEqual([{ level: 1, text: '背景介绍' }])
    editor.destroy()
  })

  it('skips empty headings', () => {
    const editor = createEditor([heading('', 1), numberedHeading('  ', '2')])
    expect(collectHeadings(editor.state.doc, STYLES)).toEqual([])
    editor.destroy()
  })
})

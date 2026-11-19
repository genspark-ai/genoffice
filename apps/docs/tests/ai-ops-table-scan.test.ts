import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const cell = (label: string): JsonNode => ({ type: 'docTableCell', content: [para([text(label)])] })
const table = (rows: string[][]): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: null },
  content: rows.map((r) => ({ type: 'docTableRow', content: r.map(cell) })),
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

function cellTexts(editor: Editor): Array<{ text: string; bold: boolean }> {
  const out: Array<{ text: string; bold: boolean }> = []
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docTableCell') return true
    node.descendants((n) => {
      if (n.isText)
        out.push({ text: n.text ?? '', bold: n.marks.some((m) => m.type.name === 'bold') })
      return true
    })
    return false
  })
  return out
}

setModuleLang('en')

describe('character-level ops reach table cells', () => {
  it('setMatchedFont styles matches inside a table, with and without a target', () => {
    const editor = createEditor([
      para([text('Status: TODO')]),
      table([
        ['Item', 'Status'],
        ['Roof', 'TODO'],
      ]),
    ])
    const outcome = executeOps(editor, [{ op: 'setMatchedFont', text: 'TODO', bold: true }])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 2, detail: '2' })
    expect(cellTexts(editor)).toEqual([
      { text: 'Item', bold: false },
      { text: 'Status', bold: false },
      { text: 'Roof', bold: false },
      { text: 'TODO', bold: true },
    ])

    const scoped = executeOps(editor, [
      { op: 'setMatchedFont', text: 'Item', target: { nodeType: 'table' }, bold: true },
    ])
    expect(scoped.ok, scoped.error).toBe(true)
    expect(scoped.results[0]).toMatchObject({ matched: 1 })
    expect(cellTexts(editor)[0]).toEqual({ text: 'Item', bold: true })
  })

  it('findReplace rewrites text inside a table cell and reports the block', () => {
    const editor = createEditor([table([['Q3 revenue', 'Q3 cost']]), para([text('Q3 summary')])])
    const outcome = executeOps(editor, [{ op: 'findReplace', find: 'Q3', replace: 'Q4' }])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 2 })
    expect(cellTexts(editor).map((c) => c.text)).toEqual(['Q4 revenue', 'Q4 cost'])
    expect(editor.state.doc.lastChild?.textContent).toBe('Q4 summary')
  })

  it('a range target overlapping a table only touches cells inside the range', () => {
    const editor = createEditor([table([['TODO', 'TODO']])])
    // table(1) + row(1) + cell(1) + paragraph(1): first cell text starts at 4
    const outcome = executeOps(editor, [
      { op: 'setMatchedFont', text: 'TODO', target: { range: { from: 4, to: 8 } }, bold: true },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(cellTexts(editor)).toEqual([
      { text: 'TODO', bold: true },
      { text: 'TODO', bold: false },
    ])
    const replaced = executeOps(editor, [
      { op: 'findReplace', find: 'TODO', replace: 'DONE', target: { range: { from: 4, to: 8 } } },
    ])
    expect(replaced.ok, replaced.error).toBe(true)
    expect(cellTexts(editor).map((c) => c.text)).toEqual(['DONE', 'TODO'])
  })

  it('an empty target object means the whole document, not a validation error', () => {
    const editor = createEditor([para([text('alpha beta')]), para([text('beta gamma')])])
    const outcome = executeOps(editor, [
      { op: 'setMatchedFont', text: 'beta', target: {}, bold: true },
    ])
    expect(outcome.ok, outcome.error).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 2 })
    const required = executeOps(editor, [{ op: 'setFont', target: {}, bold: true }])
    expect(required.ok).toBe(false)
    expect(required.error).toContain('at least one condition')
  })
})

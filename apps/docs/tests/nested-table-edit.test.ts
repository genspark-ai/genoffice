import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type TableModel } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildDocx, NESTED_TABLE_XML } from '../../../packages/docx-engine/tests/helpers/build-docx'

async function openNestedDoc() {
  const bytes = await buildDocx({ bodyXml: NESTED_TABLE_XML })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, bytes }
}

/** equivalent of the NodeView commit: mutate the docNestedTable model attribute */
function editNestedCell(editor: Editor, newText: string) {
  let done = false
  editor.state.doc.descendants((node, pos) => {
    if (done || node.type.name !== 'docNestedTable') return
    const model = node.attrs.model as TableModel
    const rows = model.rows.map((row, r) =>
      row.map((cell, c) => (r === 0 && c === 0 ? { ...cell, paras: [newText] } : cell)),
    )
    editor.view.dispatch(
      editor.view.state.tr.setNodeMarkup(pos, undefined, { model: { ...model, rows } }),
    )
    done = true
  })
  expect(done).toBe(true)
}

describe('nested table editing end to end', () => {
  it('untouched round-trip is byte-identical', async () => {
    const { editor, parsed, bytes } = await openNestedDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(bytes)
    editor.destroy()
  })

  it('nested-cell text edits use the surgical patch, outer bytes untouched', async () => {
    const { editor, parsed } = await openNestedDoc()
    editNestedCell(editor, 'EditedInnerA')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const outer = reparsed.blocks[0].table!
    expect(outer.rows[0][0].paras[0]).toBe('Outer')
    expect(outer.rows[0][1].paras[0]).toBe('RightCell')
    expect(outer.rows[0][0].nestedTables![0].rows[0].map((c) => c.paras[0])).toEqual([
      'EditedInnerA',
      'InnerB',
    ])
    editor.destroy()
  })

  it('outer structural change (formatting) no longer drops nested tables on regeneration', async () => {
    const { editor, parsed } = await openNestedDoc()
    // changing the right cell's shading changes the formatting signature, taking the whole-table regeneration path
    let done = false
    editor.state.doc.descendants((node, pos) => {
      if (done || node.type.name !== 'docTableCell') return true
      if (node.textContent !== 'RightCell') return true
      editor.view.dispatch(
        editor.view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, fill: 'FFFF00' }),
      )
      done = true
      return false
    })
    expect(done).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const outer = reparsed.blocks[0].table!
    expect(outer.rows[0][1].fill).toBe('FFFF00')
    expect(outer.rows[0][0].nestedTables![0].rows[0].map((c) => c.paras[0])).toEqual([
      'InnerA',
      'InnerB',
    ])
    editor.destroy()
  })
})

// the DOM read-back is not a faithful text encoding of the model (a trailing
// empty paragraph renders as a stripped newline): only cells the user typed in
// may be committed, or every save rewrote the untouched nested cell lossily
const NESTED_TRAILING_EMPTY_XML = NESTED_TABLE_XML.replace(
  '<w:t>InnerA</w:t></w:r></w:p>',
  '<w:t>InnerA</w:t></w:r></w:p><w:p/>',
)

function nestedModel(editor: Editor): TableModel {
  let model: TableModel | null = null
  editor.state.doc.descendants((node) => {
    if (!model && node.type.name === 'docNestedTable') model = node.attrs.model as TableModel
  })
  return model!
}

describe('nested cell pre-save commit', () => {
  it('leaves untyped cells alone even when the DOM text differs from the model', async () => {
    const bytes = await buildDocx({ bodyXml: NESTED_TRAILING_EMPTY_XML })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(nestedModel(editor).rows[0][0].paras).toEqual(['InnerA', ''])
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    expect(nestedModel(editor).rows[0][0].paras).toEqual(['InnerA', ''])
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(bytes)
    editor.destroy()
  })

  it('commits the cell the user typed in', async () => {
    const { editor } = await openNestedDoc()
    const td = editor.view.dom.querySelector('.doc-nested-table td') as HTMLElement
    td.textContent = 'Typed'
    td.dispatchEvent(new Event('input', { bubbles: true }))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    expect(nestedModel(editor).rows[0].map((c) => c.paras[0])).toEqual(['Typed', 'InnerB'])
    editor.destroy()
  })
})

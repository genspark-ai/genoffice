import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { BLANK_BULLET_NUM_ID, buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openBlank() {
  const source = await buildBlankDocx()
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** caret into the first paragraph of the first table cell */
function selectFirstCellParagraph(editor: Editor) {
  let target = -1
  editor.state.doc.descendants((node, pos) => {
    if (target >= 0) return false
    if (node.type.name === 'docTableCell') {
      target = pos + 2
      return false
    }
    return true
  })
  expect(target).toBeGreaterThan(0)
  editor.commands.setTextSelection(target)
}

async function saveAndUnzip(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  return { bytes, xml }
}

describe('list toggle inside a table cell', () => {
  it('setNode docListItem succeeds in a cell and saves w:numPr into the tc', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 2, 2)
    selectFirstCellParagraph(editor)
    editor.commands.insertContent('清单项')
    selectFirstCellParagraph(editor)
    const applied = editor
      .chain()
      .setNode('docListItem', { kind: 'bullet', numId: BLANK_BULLET_NUM_ID, ilvl: 0 })
      .run()
    expect(applied).toBe(true)
    expect(editor.isActive('docListItem')).toBe(true)

    const { bytes, xml } = await saveAndUnzip(editor, parsed)
    const tc = /<w:tc>[\s\S]*?<\/w:tc>/.exec(xml)?.[0] ?? ''
    expect(tc).toContain('<w:numPr>')
    expect(tc).toContain(`<w:numId w:val="${BLANK_BULLET_NUM_ID}"/>`)

    const reparsed = await parseDocx(bytes)
    const table = reparsed.blocks.find((b) => b.type === 'table')
    expect(table?.table?.rows[0][0].richParas?.[0].list).toEqual({
      kind: 'bullet',
      numId: BLANK_BULLET_NUM_ID,
      ilvl: 0,
    })
    // reloaded editor shows the list item inside the cell again
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    let cellListItems = 0
    editor2.state.doc.descendants((node) => {
      if (node.type.name === 'docListItem') cellListItems++
      return true
    })
    expect(cellListItems).toBe(1)
    editor2.destroy()
    editor.destroy()
  })
})

describe('insert table inside a table cell', () => {
  it('nests the new table in the cell instead of splitting the outer table', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 2, 2)
    selectFirstCellParagraph(editor)
    insertTableAt(editor, 3, 3)

    const topTables = (editor.getJSON().content ?? []).filter(
      (n: { type?: string }) => n.type === 'docTable',
    )
    expect(topTables).toHaveLength(1)
    let nested = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docNestedTable') nested++
      return true
    })
    expect(nested).toBe(1)

    const { bytes, xml } = await saveAndUnzip(editor, parsed)
    // nested w:tbl inside a w:tc, and only one top-level table
    expect(/<w:tc>[\s\S]*?<w:tbl>/.test(xml)).toBe(true)
    const reparsed = await parseDocx(bytes)
    const tables = reparsed.blocks.filter((b) => b.type === 'table')
    expect(tables).toHaveLength(1)
    expect(tables[0].table?.rows[0][0].nestedTables).toHaveLength(1)
    expect(tables[0].table?.rows[0][0].nestedTables?.[0].rows).toHaveLength(3)
    editor.destroy()
  })
})

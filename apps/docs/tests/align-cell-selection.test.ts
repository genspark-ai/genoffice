import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, type PmNode } from '../src/renderer/editor/convert'
import { setSelectionAlign } from '../src/renderer/editor/direction'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('A1')}${cell('B1')}${cell('C1')}</w:tr>` +
  `<w:tr>${cell('A2')}${cell('B2')}${cell('C2')}</w:tr></w:tbl>`

async function open(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

function cellPos(editor: Editor, text: string): number {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found < 0 && node.type.name === 'docTableCell' && node.textContent === text) found = pos
  })
  return found
}

const aligns = (editor: Editor) =>
  (editor.getJSON().content![0] as unknown as PmNode).content!.map((row) =>
    row.content!.map((c) => c.content![0].attrs?.align ?? null),
  )

describe('alignment over a table cell selection (Word)', () => {
  it('centers the text of every selected cell, not only the last one', async () => {
    const editor = await open()
    // a row picked from the left margin: one range per cell
    const row = CellSelection.create(editor.state.doc, cellPos(editor, 'A1'), cellPos(editor, 'C1'))
    editor.view.dispatch(editor.state.tr.setSelection(row))
    expect(editor.state.selection.ranges).toHaveLength(3)
    expect(setSelectionAlign(editor, 'center')).toBe(true)
    expect(aligns(editor)).toEqual([
      ['center', 'center', 'center'],
      [null, null, null],
    ])
    editor.destroy()
  })

  it('covers a column and a block of cells too', async () => {
    const editor = await open()
    const column = CellSelection.create(
      editor.state.doc,
      cellPos(editor, 'B1'),
      cellPos(editor, 'B2'),
    )
    editor.view.dispatch(editor.state.tr.setSelection(column))
    expect(setSelectionAlign(editor, 'right')).toBe(true)
    expect(aligns(editor)).toEqual([
      [null, 'right', null],
      [null, 'right', null],
    ])
    editor.destroy()
  })
})

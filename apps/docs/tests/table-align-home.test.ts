import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { generateTableModelXml, parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { alignSelectionOrTable, setSelectionAlign } from '../src/renderer/editor/direction'
import { ParagraphDialog } from '../src/renderer/components/ParagraphDialog'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>`

async function open(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

const tableNode = (editor: Editor) => editor.getJSON().content![0] as unknown as PmNode
const cellAlign = (editor: Editor) =>
  tableNode(editor).content![0].content![0].content![0].attrs?.align ?? null

describe('Home alignment with the whole table selected (Word)', () => {
  it('Center moves the table, leaves the cell text, and keeps the table selected', async () => {
    const editor = await open()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    expect(alignSelectionOrTable(editor, 'center')).toBe(true)
    expect(tableNode(editor).attrs).toMatchObject({ tblAlign: 'center' })
    expect(cellAlign(editor)).toBeNull()
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    expect(computeFormatState(editor).align).toBe('center')
    const xml = generateTableModelXml(pmTableToModel(tableNode(editor)))
    expect(xml).toContain('<w:jc w:val="center"/>')

    // Left takes it back: no w:jc
    expect(alignSelectionOrTable(editor, 'left')).toBe(true)
    expect(tableNode(editor).attrs).toMatchObject({ tblAlign: 'left' })
    expect(generateTableModelXml(pmTableToModel(tableNode(editor)))).not.toContain('<w:jc ')
    editor.destroy()
  })

  it('a caret in a cell still aligns the cell text', async () => {
    const editor = await open()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(4))),
    )
    expect(alignSelectionOrTable(editor, 'center')).toBe(true)
    expect(tableNode(editor).attrs?.tblAlign ?? null).toBeNull()
    expect(cellAlign(editor)).toBe('center')
    editor.destroy()
  })

  it('the Paragraph dialog leaves a selected table where it is', async () => {
    const editor = await open()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    alignSelectionOrTable(editor, 'center')
    // the dialog applies its alignment through setSelectionAlign: paragraphs only
    setSelectionAlign(editor, 'left')
    expect(tableNode(editor).attrs).toMatchObject({ tblAlign: 'center' })

    // and through the dialog itself: OK with only Space After changed
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)
    await act(async () => {
      root.render(createElement(ParagraphDialog, { editor, onClose: () => undefined }))
    })
    const ok = host.querySelector<HTMLButtonElement>('.modal-actions .btn-primary')!
    expect(ok).not.toBeNull()
    await act(async () => ok.click())
    expect(tableNode(editor).attrs).toMatchObject({ tblAlign: 'center' })
    act(() => root.unmount())
    host.remove()
    editor.destroy()
  })
})

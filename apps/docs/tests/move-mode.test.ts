import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { executeOps } from '../src/renderer/ai/ops'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const TABLE =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Selectable cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

const TEXTBOX =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="3000000" cy="1000000"/>' +
  '<a:graphic><a:graphicData><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:ext cx="3000000" cy="1000000"/></a:xfrm></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Selectable textbox text</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback/></mc:AlternateContent></w:r></w:p>'

async function openDocument(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: `${TABLE}${TEXTBOX}` }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return editor
}

function verifyInteractionModes(
  editor: Editor,
  wrapper: HTMLElement,
  textTarget: HTMLElement,
  editableElement: HTMLElement,
) {
  const handle = wrapper.querySelector('.doc-move-handle') as HTMLElement
  expect(handle).not.toBeNull()
  expect(wrapper.classList.contains('doc-content-editing')).toBe(false)
  expect(wrapper.draggable).toBe(true)
  expect(editableElement.getAttribute('contenteditable')).toBe('false')

  textTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
  expect(wrapper.draggable).toBe(true)
  expect(editor.state.selection).toBeInstanceOf(NodeSelection)

  textTarget.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
  expect(wrapper.classList.contains('doc-content-editing')).toBe(true)
  expect(wrapper.draggable).toBe(false)
  expect(editableElement.getAttribute('contenteditable')).toBe('true')

  textTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
  expect(wrapper.draggable).toBe(false)

  document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
  expect(wrapper.classList.contains('doc-content-editing')).toBe(false)
  expect(wrapper.draggable).toBe(true)
  expect(editableElement.getAttribute('contenteditable')).toBe('false')

  textTarget.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
  handle.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
  expect(wrapper.draggable).toBe(true)
  expect(wrapper.classList.contains('doc-content-editing')).toBe(false)
  expect(editor.state.selection).toBeInstanceOf(NodeSelection)
}

describe('object move and text modes', () => {
  it('keeps native table cells editable while the table remains a top-level movable node', async () => {
    const editor = await openDocument()
    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    const cell = table.querySelector('td') as HTMLElement
    let cellPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (cellPos === -1 && node.type.name === 'docTableCell') cellPos = pos
    })

    expect(table).not.toBeNull()
    expect(cell.textContent).toBe('Selectable cell text')
    expect(cellPos).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, cellPos + 2)),
    )
    editor.view.dispatch(editor.state.tr.insertText('+'))
    expect(cell.textContent).toContain('+')
    editor.destroy()
  })

  it('uses a dedicated textbox handle while inner drags remain text selection', async () => {
    const editor = await openDocument()
    const wrapper = editor.view.dom.querySelector('.doc-protected-textboxes') as HTMLElement
    const textTarget = wrapper.querySelector('.doc-textbox') as HTMLElement
    const textEditor = wrapper.querySelector('.doc-textbox-editor') as HTMLElement

    verifyInteractionModes(editor, wrapper, textTarget, textEditor)
    editor.destroy()
  })

  it('persists a moved table/textbox atom in the new document order', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Leading text</w:t></w:r></w:p>' +
        TABLE +
        '<w:p><w:r><w:t>Middle text</w:t></w:r></w:p>' +
        TEXTBOX,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })

    const result = executeOps(editor, [{ op: 'moveBlocks', blockIndexes: [1], afterBlockIndex: 3 }])
    expect(result.ok).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const visibleTypes = reparsed.blocks.filter((block) => !block.hidden).map((block) => block.type)

    // a table moved to the end gets Word's mandatory trailing paragraph
    expect(visibleTypes).toEqual(['paragraph', 'paragraph', 'passthrough', 'table', 'paragraph'])
    expect(reparsed.blocks[2].textboxes?.[0].paras[0].runs[0].text).toBe('Selectable textbox text')
    editor.destroy()
  })
})

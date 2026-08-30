import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// Cells with line breaks, page breaks, tabs and symbol runs: their run texts
// carry \n \f \t and decoded symbol chars, which the parse-side plain-text
// paras cache drops. The save-plan text compare must not read that encoding
// gap as a user edit (it used to send the whole untouched table through the
// lossy cell-text patch on every save).
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold head</w:t></w:r>' +
  '<w:r><w:br/><w:t>plain tail</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t xml:space="preserve">trail </w:t></w:r>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t>ital</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '</w:tbl>'

async function openDoc() {
  const source = await buildDocx({ bodyXml: TABLE })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('table cells with breaks/tabs survive an untouched save', () => {
  it('keeps the untouched table on the original-bytes path', async () => {
    const { editor, parsed } = await openDoc()
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    const tableBlocks = plan.saveBlocks.filter(
      (b) => 'docxIndex' in b || JSON.stringify(b).includes('w:tbl'),
    )
    expect(tableBlocks.every((b) => (b as { kind: string }).kind === 'original')).toBe(true)
  })

  it('a real cell edit patches texts with break/tab elements, not literal chars', async () => {
    const { editor, parsed } = await openDoc()
    // append text to the "bold head" cell paragraph
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'bold head') at = pos + node.text.length
    })
    expect(at).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText('X', at),
    )
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const xml = plan.saveBlocks
      .map((b) => (b as { xml?: string }).xml ?? '')
      .find((x) => x.includes('w:tbl'))
    expect(xml).toBeTruthy()
    expect(xml).toContain('bold headX')
    // the touched cell's line break must come back as an element
    expect(xml).toContain('<w:br/>')
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/)
    // untouched cells keep their original bytes (page break, tab, italic run)
    expect(xml).toContain('<w:br w:type="page"/>')
    expect(xml).toContain('<w:tab/>')
    expect(xml).toContain('<w:i/>')
  })
})

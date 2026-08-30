import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// A real edit inside a table cell used to go through the plain-text cell patch:
// every run collapsed onto the first run's rPr and hyperlinks lost their
// w:hyperlink wrapper. The rich per-paragraph patch must keep both.
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t xml:space="preserve">contact </w:t></w:r>' +
  '<w:hyperlink r:id="rId9"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>mail us</w:t></w:r></w:hyperlink>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> now</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>second para</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>other cell</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '</w:tbl>'

const RELS =
  '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:team@example.com" TargetMode="External"/>'

async function openDoc() {
  const source = await buildDocx({ bodyXml: TABLE, extraRels: RELS })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('rich table cell edits', () => {
  it('keeps hyperlink and per-run formatting when the cell text is edited', async () => {
    const { editor, parsed } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'contact ') at = pos + 'contact'.length
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
    expect(xml).toContain('contactX ')
    // hyperlink wrapper + relationship id survive
    expect(xml).toContain('<w:hyperlink r:id="rId9">')
    expect(xml).toContain('>mail us</w:t>')
    // per-run formatting of the sibling run survives
    expect(xml).toMatch(/<w:b\/>[\s\S]{0,40}<w:t xml:space="preserve"> now<\/w:t>/)
    // untouched second paragraph and neighbor cell keep original bytes
    expect(xml).toContain('<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>second para</w:t></w:r></w:p>')
    expect(xml).toContain('<w:t>other cell</w:t>')
  })
})

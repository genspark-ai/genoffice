import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// Cell runs carrying a comment range / tracked change next to explicit run
// formatting: the parse side lists marks as [comment, docTextStyle] while the
// schema stores them in rank order [docTextStyle, comment]. The formatting
// signature must not read that re-sort as an edit (it used to regenerate the
// whole untouched table, dropping the comment and revision markers).
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const COMMENTS_XML =
  XML_DECL +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="3" w:author="Alice"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment>' +
  '</w:comments>'
const TABLE =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
  '<w:p><w:commentRangeStart w:id="3"/>' +
  '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>noted</w:t></w:r>' +
  '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' +
  '<w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">' +
  '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>added</w:t></w:r></w:ins>' +
  '<w:del w:id="10" w:author="A" w:date="2026-01-01T00:00:00Z">' +
  '<w:r><w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr><w:delText>gone</w:delText></w:r></w:del>' +
  '</w:p></w:tc></w:tr></w:tbl>'

async function openDoc() {
  const bytes = await buildDocx({
    bodyXml: TABLE + '<w:p><w:r><w:t>after</w:t></w:r></w:p>',
    extraParts: [
      {
        path: 'word/comments.xml',
        xml: COMMENTS_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
    ],
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { bytes, parsed, editor }
}

describe('table formatting signature vs schema mark order', () => {
  it('an untouched table with commented and tracked cell runs stays byte-identical', async () => {
    const { bytes, parsed, editor } = await openDoc()
    let noted: readonly string[] = []
    editor.state.doc.descendants((node) => {
      if (node.isText && node.text === 'noted') noted = node.marks.map((m) => m.type.name)
    })
    expect(noted).toContain('comment')
    expect(noted[0]).not.toBe('comment')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.saveBlocks[0]).toEqual({ kind: 'original', docxIndex: 0 })
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(bytes)
    editor.destroy()
  })

  it('a real formatting change in the cell still regenerates the table', async () => {
    const { parsed, editor } = await openDoc()
    let from = -1
    let to = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'noted') {
        from = pos
        to = pos + node.text.length
      }
    })
    editor.view.dispatch(editor.state.tr.addMark(from, to, editor.schema.marks.italic.create()))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    expect(plan.saveBlocks[0].kind).toBe('xml')
    editor.destroy()
  })
})

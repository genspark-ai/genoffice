import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// A style-chain w:lang w:eastAsia reaches every paragraph as a display-only
// ParaFormat field (and every run as a mark attr). The regenerated block must
// read both back, or a paragraph with no pPr of its own signs as "formatted"
// on the parse side and "unformatted" on the editor side: the untouched TOC
// sdt members (title + empty paragraph) were re-serialized on every save.
const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-SG" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>'

const TOC_ENTRY =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>Chapter 1</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:p>'

const BODY =
  '<w:p w:rsidR="00AA0001"><w:r><w:t>Front matter needle</w:t></w:r></w:p>' +
  '<w:sdt><w:sdtPr><w:id w:val="-1"/>' +
  '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>' +
  '</w:sdtPr><w:sdtEndPr/><w:sdtContent>' +
  '<w:p w:rsidR="00AA0002"><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' +
  '<w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>Table of Contents</w:t></w:r></w:p>' +
  '<w:p w:rsidR="00AA0003"/>' +
  TOC_ENTRY +
  '</w:sdtContent></w:sdt>' +
  '<w:p><w:r><w:t>Body after the TOC</w:t></w:r></w:p>'

async function openDoc() {
  const bytes = await buildDocx({ bodyXml: BODY, stylesXml: STYLES_XML })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { bytes, parsed, editor }
}

describe('TOC sdt under a style chain declaring w:lang eastAsia', () => {
  it('parses the inherited East Asian lang onto pPr-less paragraphs', async () => {
    const { parsed } = await openDoc()
    const empty = parsed.blocks.find((b) => b.originalXml === '<w:p w:rsidR="00AA0003"/>')
    expect(empty?.format?.eastAsiaLang).toBe('zh-CN')
  })

  it('an untouched save keeps every block on the original-bytes path', async () => {
    const { bytes, parsed, editor } = await openDoc()
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('an edit outside the sdt leaves the sdt members untouched', async () => {
    const { parsed, editor } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Front matter needle') at = pos + node.text.length
    })
    expect(at).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText('X', at),
    )
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    expect(plan.saveBlocks[0].kind).toBe('generated')
    expect(plan.saveBlocks.slice(1).every((b) => b.kind === 'original')).toBe(true)
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('w:rsidR="00AA0002"')
    expect(xml).toContain('<w:p w:rsidR="00AA0003"/>')
  })

  it('a regenerated paragraph keeps the run-level East Asian lang in the model', async () => {
    const { parsed, editor } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Table of Contents') at = pos + node.text.length
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText('!', at),
    )
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const generated = plan.saveBlocks.find((b) => b.kind === 'generated') as
      { block: { runs: Array<{ text: string; eastAsiaLang?: string }> } } | undefined
    expect(generated?.block.runs[0]?.eastAsiaLang).toBe('ja-JP')
  })
})

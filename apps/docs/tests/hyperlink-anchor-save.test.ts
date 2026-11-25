import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

it('keeps the external hyperlink bookmark and location attributes when editing nearby text', async () => {
  const untouched = '<w:p><w:r><w:t>untouched</w:t></w:r></w:p>'
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:hyperlink r:id="rId11" w:anchor="section2" w:history="1" w:tgtFrame="_blank" w:docLocation="page-two">' +
      '<w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>after</w:t></w:r></w:p>' +
      untouched,
    extraRels:
      '<Relationship Id="rId11" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      'Target="https://example.com/doc.html" TargetMode="External"/>',
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks),
  })
  try {
    editor.view.dispatch(editor.state.tr.insertText('AFTER', 5, 10))
    const plan = pmDocToSavePlan(editor.state.doc.toJSON(), parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const output = await saveDocx(parsed, plan.saveBlocks)
    const reopened = await parseDocx(output)
    const paragraph = reopened.blocks[0]
    if (paragraph.type !== 'paragraph' || !paragraph.runs) {
      throw new Error('Expected a paragraph with runs')
    }
    expect(paragraph.runs.map((run) => run.text).join('')).toBe('linkAFTER')
    expect(paragraph.runs[0].link?.href).toBe('https://example.com/doc.html#section2')

    const zip = await JSZip.loadAsync(output)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml.match(/<w:hyperlink\b[^>]*>/)?.[0]).toContain('w:history="1"')
    expect(xml.match(/<w:hyperlink\b[^>]*>/)?.[0]).toContain('w:tgtFrame="_blank"')
    expect(xml.match(/<w:hyperlink\b[^>]*>/)?.[0]).toContain('w:docLocation="page-two"')
    expect(xml).toContain(untouched)
    const relsPath = 'word/_rels/document.xml.rels'
    const sourceZip = await JSZip.loadAsync(bytes)
    expect(await zip.file(relsPath)!.async('string')).toBe(
      await sourceZip.file(relsPath)!.async('string'),
    )
  } finally {
    editor.destroy()
  }
})

import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// A run holding only a page break still has its own rPr (Word sizes the break
// line by it). The editor models the break as a hardBreak node: without marks
// the round trip produced a bare run, so every break-only paragraph looked
// edited and was regenerated without its formatting.
const BODY =
  '<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:br w:type="page"/></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:i/></w:rPr><w:br w:type="page"/></w:r>' +
  '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Chapter</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>plain</w:t></w:r></w:p>'

async function openDoc() {
  const bytes = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { bytes, parsed, editor }
}

describe('break-only runs keep their rPr', () => {
  it('untouched page-break paragraphs stay on the original-bytes path', async () => {
    const { bytes, parsed, editor } = await openDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(bytes)
    editor.destroy()
  })

  it('an edited paragraph re-emits its break run with the original rPr', async () => {
    const { parsed, editor } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Chapter') at = pos + node.text.length
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText(' 1', at),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const block = plan.saveBlocks[1]
    if (block.kind !== 'generated') throw new Error(`expected generated, got ${block.kind}`)
    expect(block.block.runs.map((r) => [r.text, !!r.italic, r.sizeHalfPoints ?? null])).toEqual([
      ['\f', true, null],
      ['Chapter 1', false, 24],
    ])
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:rPr><w:i/></w:rPr><w:br w:type="page"/>')
    editor.destroy()
  })
})

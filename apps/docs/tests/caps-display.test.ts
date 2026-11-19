import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

/** w:caps / w:smallCaps display (POI capitalized.docx: Word shows CAPITALIZED, we showed lowercase) */

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">The following word is: </w:t></w:r>' +
  '<w:r><w:rPr><w:caps/></w:rPr><w:t>capitalized</w:t></w:r>' +
  '<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>small</w:t></w:r>' +
  '<w:r><w:rPr><w:caps/><w:smallCaps/></w:rPr><w:t>both</w:t></w:r>' +
  '<w:r><w:t>.</w:t></w:r></w:p>'

async function open() {
  const source = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

describe('caps display', () => {
  it('parses w:caps/w:smallCaps into run.caps; w:caps wins when both are set', async () => {
    const { editor, parsed } = await open()
    const runs = parsed.blocks[0].runs!
    expect(runs.map((r) => r.caps)).toEqual([undefined, 'all', 'small', 'all', undefined])
    editor.destroy()
  })

  it('renders text-transform / font-variant-caps without changing the text', async () => {
    const { editor } = await open()
    const spans = [...editor.view.dom.querySelectorAll<HTMLElement>('span[data-doc-style]')]
    const byText = (t: string) => spans.find((s) => s.textContent === t)!
    expect(byText('capitalized').style.textTransform).toBe('uppercase')
    expect(byText('small').style.fontVariantCaps).toBe('small-caps')
    expect(byText('both').style.textTransform).toBe('uppercase')
    editor.destroy()
  })

  it('untouched document stays byte-identical (caps saves via rawRPr)', async () => {
    const { editor, parsed, source } = await open()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })
})

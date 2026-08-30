import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, type Run } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// w:rtl is a run-managed rPr group on the generate side: when a block is
// regenerated after a real edit, the flag must come back from the PM model
// or every rtl run silently turns ltr (Word then reads w:b instead of w:bCs).
const BODY =
  '<w:p><w:r><w:rPr><w:rtl/><w:bCs/></w:rPr><w:t>مرحبا بالعالم</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> plain tail</w:t></w:r></w:p>'

async function openDoc() {
  const source = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('run-level w:rtl survives block regeneration', () => {
  it('an untouched save keeps the block on the original-bytes path', async () => {
    const { editor, parsed } = await openDoc()
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
  })

  it('a real edit regenerates the paragraph with w:rtl intact', async () => {
    const { editor, parsed } = await openDoc()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === ' plain tail') at = pos + node.text.length
    })
    expect(at).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).insertText('X', at),
    )
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const generated = plan.saveBlocks.find((b) => b.kind === 'generated')
    expect(generated).toBeTruthy()
    const runs = (generated as { block: { runs: Run[] } }).block.runs
    expect(runs.map((r) => r.text).join('')).toBe('مرحبا بالعالم plain tailX')
    const rtlRun = runs.find((r) => r.text.startsWith('مرحبا'))
    expect(rtlRun?.rtl).toBe(true)
    expect(rtlRun?.cs).toBe(true)
  })
})

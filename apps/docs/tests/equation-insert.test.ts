import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  BUILTIN_EQUATIONS,
  insertEquationFromLatex,
} from '../src/renderer/components/EquationModal'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openDoc(): Promise<{
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
}> {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('equation insertion', () => {
  it('inserts a display equation that renders native MathML in the editor', async () => {
    const { editor } = await openDoc()
    insertEquationFromLatex(editor, 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}')

    const host = editor.view.dom.querySelector('.doc-formula-math')
    expect(host).not.toBeNull()
    expect(host!.innerHTML).toContain('<mfrac>')
    expect(host!.innerHTML).toContain('<msqrt>')
    // the editable token strip exists alongside as the edit-mode fallback
    expect(editor.view.dom.querySelectorAll('.doc-formula-token').length).toBeGreaterThan(0)
    editor.destroy()
  })

  it('saves the inserted equation and reparses it as a formula block', async () => {
    const { editor, parsed } = await openDoc()
    insertEquationFromLatex(editor, 'a^2 + b^2 = c^2')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(bytes)

    const formulaBlock = reparsed.blocks.find((b) => b.formulaDisplay)
    expect(formulaBlock).toBeDefined()
    expect(formulaBlock!.formulaDisplay!.mathml).toContain('<msup>')
    expect(formulaBlock!.formulaDisplay!.tokens.join('')).toContain('=')
    editor.destroy()
  })

  it('every builtin gallery equation converts without error', async () => {
    const { editor } = await openDoc()
    for (const eq of BUILTIN_EQUATIONS) {
      expect(() => insertEquationFromLatex(editor, eq.latex), eq.name).not.toThrow()
    }
    expect(editor.view.dom.querySelectorAll('.doc-formula-math').length).toBe(
      BUILTIN_EQUATIONS.length,
    )
    editor.destroy()
  })
})

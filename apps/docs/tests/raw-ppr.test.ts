import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for rawPPr passthrough: pPr constructs the format model cannot
 * express (keepNext, tabs, pPrChange revisions, paragraph-mark rPr) survive
 * in-editor edits byte-identically; format edits merge into the original pPr.
 */

const NUM_IDS = { bullet: null, ordered: null }

/** paragraph with exotic pPr: keepNext + tabs + tracked property change */
const EXOTIC_P =
  '<w:p><w:pPr><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>' +
  '<w:spacing w:after="120"/>' +
  '<w:pPrChange w:id="51" w:author="Alice"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>' +
  '</w:pPr><w:r><w:t>Original text content</w:t></w:r></w:p>'

async function openEditor(bodyXml: string) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

describe('rawPPr passthrough', () => {
  it('typing in a paragraph keeps its exotic pPr byte-identical', async () => {
    const { editor, parsed } = await openEditor(EXOTIC_P)
    // in-editor text edit (docxIndex anchor preserved)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)).insertText('edit:'),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const xml = new TextDecoder().decode(
      await (
        await import('jszip')
      ).default
        .loadAsync(saved)
        .then((z) => z.file('word/document.xml')!.async('uint8array')),
    )
    expect(xml).toContain('edit:Original text content')
    expect(xml).toContain(parsed.blocks[0].rawPPr!)
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:pPrChange w:id="51"')
  })

  it('a format command merges into the pPr, keeping unmanaged children', async () => {
    const { editor, parsed } = await openEditor(EXOTIC_P)
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'apply_ops',
        input: {
          ops: [{ op: 'setParagraphFormat', target: { blockIndexes: [0] }, align: 'center' }],
        },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].format?.align).toBe('center')
    const xml = reparsed.internal.documentXml
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>')
    expect(xml).toContain('<w:pPrChange w:id="51"')
    // spacing is format-managed: rebuilt, old bytes replaced
    expect(xml).toContain('<w:spacing w:after="120"/>')
  })

  it('splitting a paragraph does not duplicate the revision record', async () => {
    const { editor, parsed } = await openEditor(EXOTIC_P)
    // split after the 3rd character (pos 1 + 3)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 4)))
    editor.commands.splitBlock()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const reparsed = await parseDocx(saved)
    const changeCount = (reparsed.internal.documentXml.match(/<w:pPrChange/g) ?? []).length
    expect(changeCount).toBe(1)
  })

  it('untouched exotic paragraphs still save byte-identical', async () => {
    const { editor, parsed } = await openEditor(EXOTIC_P)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    editor.destroy()
  })
})

import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { setModuleLang } from '../src/renderer/i18n/locale'

setModuleLang('en')

const RED_PARA =
  '<w:p><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>' +
  '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>review note</w:t></w:r></w:p>'
const UNTOUCHED =
  '<w:p><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>' +
  '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>still red</w:t></w:r></w:p>'

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

async function roundTrip(ops: unknown[]): Promise<string[]> {
  const bytes = await buildDocx({ bodyXml: RED_PARA + UNTOUCHED })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.add(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const outcome = executeOps(editor, ops)
  expect(outcome.ok, outcome.error).toBe(true)
  const saved = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  return xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? []
}

describe('paragraph mark color follows a recolor of the whole run text', () => {
  it('clearing the red runs also clears the red paragraph mark; an untouched paragraph keeps it', async () => {
    const [edited, untouched] = await roundTrip([
      { op: 'setMatchedFont', matchColor: '#FF0000', target: { blockIndexes: [0] }, color: null },
    ])
    expect(edited).not.toContain('FF0000')
    expect(untouched).toContain('<w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>')
  })

  it('recoloring to another color drops the stale mark color too', async () => {
    const [edited] = await roundTrip([
      {
        op: 'setMatchedFont',
        matchColor: '#FF0000',
        target: { blockIndexes: [0] },
        color: '#0000FF',
      },
    ])
    expect(edited).not.toMatch(/<w:pPr>[\s\S]*FF0000[\s\S]*<\/w:pPr>/)
    expect(edited).toContain('w:val="0000FF"')
  })
})

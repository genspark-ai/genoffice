import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  applySectionSettings,
  parseDocx,
  readSections,
  saveDocx,
  type SectionSettings,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeTool } from '../src/renderer/ai/tools'
import { applySectPrRewrites } from '../src/renderer/sectpr-rewrite'

/**
 * A mid-document section-break paragraph that carries text is a normal
 * paragraph: editing its text regenerates it with the old sectPr riding in
 * rawPPr. A Layout / set_page_setup change on that section must still reach
 * the file, whether or not the paragraph itself was edited in the same session.
 */

const NUM_IDS = { bullet: null, ordered: null }

const SECTION_P =
  '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
  '</w:sectPr></w:pPr><w:r><w:t>Last line of section one</w:t></w:r></w:p>'
const PLAIN_P = '<w:p><w:r><w:t>Section two text</w:t></w:r></w:p>'
const BODY = `<w:p><w:r><w:t>Intro</w:t></w:r></w:p>${SECTION_P}${PLAIN_P}`

const LANDSCAPE: Partial<SectionSettings> = {
  pageWidth: 15840,
  pageHeight: 12240,
  orientation: 'landscape',
  marginLeft: 720,
  marginRight: 720,
}

async function saveWithLayoutChange(ops: unknown[]) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  if (ops.length > 0) {
    const exec = await executeTool(editor, { id: 't', name: 'apply_ops', input: { ops } }, NUM_IDS)
    expect(exec.isError).toBeFalsy()
  }
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  editor.destroy()

  const sec = readSections(parsed)[0]!
  const blk = parsed.blocks.find((b) => b.docxIndex === sec.lastBlockIndex)!
  const rewrites = new Map([
    [
      sec.lastBlockIndex,
      {
        from: sec.sectPrXml,
        to: applySectionSettings(sec.sectPrXml, { ...sec.settings, ...LANDSCAPE }),
        originalXml: blk.originalXml!,
      },
    ],
  ])
  const saveBlocks = applySectPrRewrites(plan.saveBlocks, plan.saveBlockIndexByDocx, rewrites)
  const saved = await saveDocx(parsed, saveBlocks)
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('word/document.xml')!.async('string')
  return { plan, xml }
}

const firstSectPr = (xml: string) => xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)![0]

describe('section layout change on the break paragraph', () => {
  it('reaches the file when the paragraph is untouched', async () => {
    const { xml } = await saveWithLayoutChange([])
    const sectPr = firstSectPr(xml)
    expect(sectPr).toContain('w:orient="landscape"')
    expect(sectPr).toContain('w:w="15840" w:h="12240"')
    expect(sectPr).toContain('w:left="720"')
  })

  it('reaches an editor-built xml fragment that lost its docxIndex', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const sec = readSections(parsed)[0]!
    const blk = parsed.blocks.find((b) => b.docxIndex === sec.lastBlockIndex)!
    const to = applySectionSettings(sec.sectPrXml, { ...sec.settings, ...LANDSCAPE })
    const rewrites = new Map([
      [sec.lastBlockIndex, { from: sec.sectPrXml, to, originalXml: blk.originalXml! }],
    ])
    const blocks = applySectPrRewrites(
      [
        { kind: 'original', docxIndex: 0 },
        { kind: 'xml', xml: blk.originalXml! },
      ],
      new Map([
        [0, 0],
        [sec.lastBlockIndex, 1],
      ]),
      rewrites,
    )
    expect(blocks[1]).toMatchObject({ kind: 'xml' })
    expect((blocks[1] as { xml: string }).xml).toContain('w:orient="landscape"')
  })

  it('reaches the file when the paragraph text was edited in the same session', async () => {
    const { plan, xml } = await saveWithLayoutChange([
      {
        op: 'findReplace',
        find: 'Last line of section one',
        replace: 'Last line of section one, edited',
        matchCase: true,
      },
    ])
    expect(plan.saveBlocks[1]?.kind).toBe('generated')
    expect(xml).toContain('Last line of section one, edited')
    const sectPr = firstSectPr(xml)
    expect(sectPr).toContain('w:orient="landscape"')
    expect(sectPr).toContain('w:w="15840" w:h="12240"')
    expect(sectPr).toContain('w:left="720"')
    expect(xml.match(/<w:sectPr/g)).toHaveLength(2)
  })
})

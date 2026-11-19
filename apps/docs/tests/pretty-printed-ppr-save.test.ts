import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Third-party generators (PHPWord, docx4j...) pretty-print document.xml, so
 * whitespace text nodes sit between <w:p> and <w:pPr> and between pPr children.
 * The raw pPr passthrough must still find the pPr, or a text edit rebuilds it
 * from the format model and drops keepNext/keepLines/widowControl and a
 * mid-document w:sectPr (a continuous multicolumn section break).
 */

const NUM_IDS = { bullet: null, ordered: null }

const HEADING_P =
  '<w:p>\n\t\t\t<w:pPr>\n\t\t\t\t<w:keepNext/>\n\t\t\t\t<w:keepLines/>\n\t\t\t\t<w:widowControl w:val="0"/>' +
  '\n\t\t\t\t<w:spacing w:before="360" w:after="240" w:line="360" w:lineRule="auto"/>' +
  '\n\t\t\t\t<w:ind w:left="720"/>\n\t\t\t\t<w:jc w:val="both"/>\n\t\t\t</w:pPr>' +
  '\n\t\t\t<w:r>\n\t\t\t\t<w:t>Formatted paragraph text</w:t>\n\t\t\t</w:r>\n\t\t</w:p>'

const SECTION_P =
  '<w:p>\n\t\t\t<w:pPr>\n\t\t\t\t<w:sectPr>\n\t\t\t\t\t<w:type w:val="continuous"/>' +
  '\n\t\t\t\t\t<w:pgSz w:w="11906" w:h="16838"/>' +
  '\n\t\t\t\t\t<w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" w:header="720" w:footer="720" w:gutter="0"/>' +
  '\n\t\t\t\t\t<w:cols w:num="2" w:space="720"/>\n\t\t\t\t</w:sectPr>\n\t\t\t</w:pPr>' +
  '\n\t\t\t<w:r>\n\t\t\t\t<w:lastRenderedPageBreak/>\n\t\t\t\t<w:t>Multicolumn section text</w:t>\n\t\t\t</w:r>\n\t\t</w:p>'

const PLAIN_P =
  '<w:p>\n\t\t\t<w:r>\n\t\t\t\t<w:t>Plain trailing text</w:t>\n\t\t\t</w:r>\n\t\t</w:p>'

const BODY = `\n\t\t${HEADING_P}\n\t\t${SECTION_P}\n\t\t${PLAIN_P}\n\t\t`

async function roundTrip(ops: unknown[]) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const exec = await executeTool(editor, { id: 't', name: 'apply_ops', input: { ops } }, NUM_IDS)
  expect(exec.isError).toBeFalsy()
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  editor.destroy()
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('word/document.xml')!.async('string')
  return { parsed, plan, xml }
}

const retype = (text: string) => ({
  op: 'findReplace',
  find: text,
  replace: `${text}!`,
  matchCase: true,
})

describe('pretty-printed pPr survives paragraph regeneration', () => {
  it('parses the raw pPr despite whitespace after <w:p>', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(parsed.blocks[0].rawPPr).toContain('<w:keepNext/>')
    expect(parsed.blocks[1].type).toBe('paragraph')
    expect(parsed.blocks[1].rawPPr).toContain('<w:sectPr>')
  })

  it('a text edit keeps keepNext/keepLines/widowControl and the mid-document sectPr', async () => {
    const { plan, xml } = await roundTrip([
      retype('Formatted paragraph text'),
      retype('Multicolumn section text'),
    ])
    expect(plan.changedCount).toBe(2)
    expect(xml).toContain('Multicolumn section text!')
    expect(xml.match(/<w:sectPr[\s>]/g)).toHaveLength(2)
    expect(xml).toContain('<w:cols w:num="2" w:space="720"/>')
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:keepLines/>')
    expect(xml).toContain('<w:widowControl w:val="0"/>')
  })

  it('a format edit merges into the pretty-printed pPr in schema order', async () => {
    const { xml } = await roundTrip([
      { op: 'setParagraphFormat', target: { blockIndexes: [0] }, align: 'center' },
    ])
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)![0]
    const names = [...pPr.matchAll(/<w:([A-Za-z]+)[\s/>]/g)]
      .map((m) => m[1])
      .filter((n) => n !== 'pPr')
    expect(names).toEqual(['keepNext', 'keepLines', 'widowControl', 'spacing', 'ind', 'jc'])
    expect(pPr).toContain('<w:jc w:val="center"/>')
  })
})

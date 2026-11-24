import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocOptions, type PmNode } from '../src/renderer/editor/convert'

// Word probe 2026-11-23: a table inserted into a new Docs document and reopened
// sat 0.19 cm into the left margin here, while Word drew it on the margin with
// its text against the borders. The blank had no compatibility mode (so the
// table indent was read the old way) and no Normal Table style (so Word gave
// the cells no margins).
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Dataset</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'

async function tableIn(wordDefaults: boolean): Promise<PmNode> {
  const zip = await JSZip.loadAsync(await buildBlankDocx({ wordDefaults }))
  const xml = await zip.file('word/document.xml')!.async('string')
  zip.file('word/document.xml', xml.replace('<w:body><w:p/>', `<w:body><w:p/>${TABLE}`))
  const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
  const doc = blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed)) as PmNode
  return doc.content!.find((node) => node.type === 'docTable')!
}

describe('a table in a new Docs document', () => {
  it('stays on the margin after reopening', async () => {
    expect((await tableIn(true)).attrs?.indentTwips ?? 0).toBe(0)
  })

  it('hung into the margin when the blank had no compatibility mode', async () => {
    expect((await tableIn(false)).attrs?.indentTwips).toBe(-108)
  })
})

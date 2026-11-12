import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, patchFieldParagraphXml, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const documentXmlOf = async (bytes: Uint8Array): Promise<string> =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

// caption as Word writes it: "Table <STYLEREF>-<SEQ>" with the hyphen in a run of its own
const CAPTION_PARAGRAPH =
  '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>' +
  '<w:r><w:t xml:space="preserve">Table </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> STYLEREF 1 \\s </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>2</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:noBreakHyphen/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> SEQ Table \\* ARABIC \\s 1 </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t>: Environmental Conditions</w:t></w:r></w:p>'

const TOC_ENTRY_PARAGRAPH =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/>' +
  '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>' +
  '<w:hyperlink w:anchor="_Toc1" w:history="1">' +
  '<w:r><w:t>Table 2</w:t></w:r>' +
  '<w:r><w:noBreakHyphen/><w:t>1: Environmental Conditions</w:t></w:r>' +
  '<w:r><w:tab/></w:r><w:r><w:t>10</w:t></w:r>' +
  '</w:hyperlink></w:p>'

describe('w:noBreakHyphen / w:softHyphen', () => {
  it('a plain run keeps both as U+2011 / U+00AD', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Table 2</w:t><w:noBreakHyphen/><w:t>1 soft</w:t><w:softHyphen/><w:t>ware</w:t></w:r></w:p>',
      }),
    )
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('Table 2‑1 soft­ware')
  })

  it('a field caption with the hyphen in its own run shows it', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CAPTION_PARAGRAPH }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.previewText).toBe('Table 2‑1: Environmental Conditions')
    expect(block.fieldDisplay).toMatchObject({
      kind: 'text',
      left: 'Table 2‑1: Environmental Conditions',
    })
  })

  it('a TOC entry keeps the hyphen in its title', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_ENTRY_PARAGRAPH }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Table 2‑1: Environmental Conditions',
      right: '10',
    })
  })

  it('round-trips both characters back to their elements', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t><w:softHyphen/><w:t>c</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(source)
    const saved = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', runs: doc.blocks[0].runs! } },
    ])
    const xml = await documentXmlOf(saved)
    expect(xml).toContain(
      'a</w:t><w:noBreakHyphen/><w:t xml:space="preserve">b</w:t><w:softHyphen/>',
    )
    expect(xml).not.toMatch(/[‑­]/)
    expect((await parseDocx(saved)).blocks[0].runs!.map((r) => r.text).join('')).toBe('a‑b­c')
  })

  it('editing a TOC title writes the hyphen back as the element, once', async () => {
    const patched = patchFieldParagraphXml(TOC_ENTRY_PARAGRAPH, {
      left: 'Table 2\u20112: Environmental Conditions (rev)',
      right: '11',
    })
    expect(patched.match(/<w:noBreakHyphen\/>/g)).toHaveLength(1)
    expect(patched).not.toMatch(/[\u2011\u00ad]/)
    const doc = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      left: 'Table 2\u20112: Environmental Conditions (rev)',
      right: '11',
    })
  })

  it('a page-number-only patch leaves the hyphen run byte-identical', () => {
    const patched = patchFieldParagraphXml(TOC_ENTRY_PARAGRAPH, { right: '11' })
    expect(patched).toContain('<w:r><w:noBreakHyphen/><w:t>1: Environmental Conditions</w:t></w:r>')
    expect(patched).toContain('<w:t>11</w:t>')
  })

  it('a tracked-deleted hyphen stays out of the live title on edit', async () => {
    const entry = TOC_ENTRY_PARAGRAPH.replace(
      '<w:r><w:t>Table 2</w:t></w:r>',
      '<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Fig. 1</w:delText><w:noBreakHyphen/></w:r></w:del>' +
        '<w:r><w:t>Table 2</w:t></w:r>',
    )
    const patched = patchFieldParagraphXml(entry, { left: 'Table 2\u20113: Edited', right: '10' })
    expect(patched).toContain('<w:r><w:delText>Fig. 1</w:delText><w:noBreakHyphen/></w:r></w:del>')
    expect(patched).toContain('<w:r><w:t>Table 2</w:t></w:r>')
    expect(patched).toContain('<w:r><w:noBreakHyphen/><w:t>3: Edited</w:t></w:r>')
    const doc = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({ left: 'Table 2\u20113: Edited' })
  })
})

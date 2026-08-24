/**
 * settings.xml balanceSingleByteDoubleByteWidth makes rPr w:spacing count
 * double on double-byte characters (Word probe 2026-08-24). The display-only
 * charSpacingTwips scales by each run's wide-glyph mix; raw bytes round-trip
 * untouched through rawRPr.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_PART = {
  path: 'word/settings.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:compat><w:balanceSingleByteDoubleByteWidth/></w:compat></w:settings>',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
}

function spacedPara(text: string, twips: number): string {
  return (
    `<w:p><w:r><w:rPr><w:spacing w:val="${twips}"/></w:rPr>` +
    `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  )
}

const BODY =
  spacedPara('정보화비전및전략수립', -20) + // pure hangul
  spacedPara('RIPC AI', -20) + // pure Latin
  spacedPara('정보RI', -20) // half wide, half narrow

describe('balanceSingleByteDoubleByteWidth character spacing', () => {
  it('doubles spacing on wide glyphs, weighted by the run mix', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY, extraParts: [SETTINGS_PART] }))
    expect(doc.balanceDbcsSpacing).toBe(true)
    expect(doc.blocks[0].runs![0].charSpacingTwips).toBe(-40) // all wide -> x2
    expect(doc.blocks[1].runs![0].charSpacingTwips).toBe(-20) // all narrow -> x1
    expect(doc.blocks[2].runs![0].charSpacingTwips).toBe(-30) // half -> x1.5
  })

  it('reaches runs inside table cells', async () => {
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>' +
      spacedPara('정보화비전', -34) +
      '</w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: tbl, extraParts: [SETTINGS_PART] }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.richParas![0].runs[0].charSpacingTwips).toBe(-68)
  })

  it('leaves everything untouched without the flag', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.balanceDbcsSpacing).toBeUndefined()
    expect(doc.blocks[0].runs![0].charSpacingTwips).toBe(-20)
    expect(doc.blocks[2].runs![0].charSpacingTwips).toBe(-20)
  })

  it('a val="0" flag counts as off', async () => {
    const off = {
      ...SETTINGS_PART,
      xml: SETTINGS_PART.xml.replace(
        '<w:balanceSingleByteDoubleByteWidth/>',
        '<w:balanceSingleByteDoubleByteWidth w:val="0"/>',
      ),
    }
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY, extraParts: [off] }))
    expect(doc.blocks[0].runs![0].charSpacingTwips).toBe(-20)
  })
})

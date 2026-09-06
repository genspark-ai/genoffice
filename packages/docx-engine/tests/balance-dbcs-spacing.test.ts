/**
 * settings.xml balanceSingleByteDoubleByteWidth makes rPr w:spacing count
 * double on double-byte characters (Word probe 2026-08-24), except hangul and
 * kana in modern faces such as Malgun Gothic (Word probe 2026-09-03). The
 * display-only charSpacingTwips scales by each run's doubled-glyph mix; raw
 * bytes round-trip untouched through rawRPr.
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

function spacedPara(text: string, twips: number, font?: string): string {
  const rFonts = font ? `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/>` : ''
  return (
    `<w:p><w:r><w:rPr>${rFonts}<w:spacing w:val="${twips}"/></w:rPr>` +
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

  it('keeps hangul and kana single in Malgun Gothic while Han still doubles', async () => {
    const body =
      spacedPara('정보화비전및전략수립', -20, 'Malgun Gothic') +
      spacedPara('정보화비전및전략수립', -20, '맑은 고딕') +
      spacedPara('あいうえおかきくけこ', -20, 'Malgun Gothic') +
      spacedPara('永和九年歲在癸丑暮春', -20, 'Malgun Gothic') +
      spacedPara('정보화비전및전략수립', -20, '바탕') +
      spacedPara('정보화비전및전략수립', -20, 'HY Missing Font') +
      spacedPara('정보화비전및전략수립', -20, 'NanumGothic')
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraParts: [SETTINGS_PART] }))
    const twips = doc.blocks.filter((b) => !b.hidden).map((b) => b.runs![0].charSpacingTwips)
    expect(twips).toEqual([-20, -20, -20, -40, -40, -40, -20])
  })

  it('falls back to the docDefaults East Asian font for runs without rFonts', async () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: BODY, stylesXml: styles, extraParts: [SETTINGS_PART] }),
    )
    expect(doc.blocks[0].runs![0].charSpacingTwips).toBe(-20)
    expect(doc.blocks[2].runs![0].charSpacingTwips).toBe(-20)
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

  it('a Latin-only rFonts still inherits the default East Asian font for the gate', async () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="Malgun Gothic"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>'
    const body =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:spacing w:val="-20"/></w:rPr>' +
      '<w:t>정보화비전및전략수립</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="바탕"/><w:spacing w:val="-20"/></w:rPr>' +
      '<w:t>정보화비전및전략수립</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: body, stylesXml: styles, extraParts: [SETTINGS_PART] }),
    )
    expect(doc.blocks[0].runs![0].charSpacingTwips).toBe(-20)
    expect(doc.blocks[1].runs![0].charSpacingTwips).toBe(-40)
  })
})

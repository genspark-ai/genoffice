import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const numberingXml =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="&#xF09F;"/><w:pPr><w:ind w:left="440" w:hanging="440"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/><w:sz w:val="16"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="-"/><w:pPr><w:ind w:left="0" w:firstLine="420"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>' +
  '</w:numbering>'

/** textbox whose paragraphs carry the given numPr (null = plain paragraph) */
function textboxXml(numIds: Array<string | null>): string {
  const xml = buildAnchoredTextboxParagraphXml({
    anchor: 'paragraph',
    xEmu: 0,
    yEmu: 0,
    widthEmu: 3000000,
    heightEmu: 1000000,
    insetsEmu: { l: 0, t: 0, r: 0, b: 0 },
    zOrder: 1,
    id: 7,
    paragraphs: numIds.map((_, i) => ({ runs: [{ text: `item ${i + 1}` }] })),
    holderLineTwips: 20,
  })
  const start = xml.indexOf('<w:txbxContent>')
  const end = xml.indexOf('</w:txbxContent>')
  let i = 0
  const content = xml.slice(start, end).replace(/<w:p><w:r>/g, () => {
    const numId = numIds[i++]
    const numPr = numId
      ? `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`
      : ''
    return `<w:p>${numPr}<w:r>`
  })
  return xml.slice(0, start) + content + xml.slice(end)
}

async function textboxParas(numIds: Array<string | null>) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: textboxXml(numIds), numberingXml }))
  return parsed.blocks.flatMap((b) => b.textboxes ?? [])[0]!.paras
}

describe('list markers inside textboxes', () => {
  it('resolves a symbol-font bullet with the level geometry', async () => {
    const paras = await textboxParas(['1', '1', null])
    expect(paras.map((p) => p.listMarker?.text)).toEqual(['•', '•', undefined])
    expect(paras[0].listMarker).toEqual({
      text: '•',
      symbol: true,
      indentLeft: 440,
      hanging: 440,
      szHalfPoints: 16,
    })
  })

  it('carries a level first-line indent instead of a hanging one', async () => {
    const paras = await textboxParas(['3'])
    expect(paras[0].listMarker).toEqual({ text: '-', firstLine: 420 })
  })

  it('continues numbering across content-control (w:sdt) boundaries', async () => {
    let xml = textboxXml(['2', '2', '2'])
    const start = xml.indexOf('<w:txbxContent>') + '<w:txbxContent>'.length
    const second = xml.indexOf('<w:p>', xml.indexOf('<w:p>', start) + 1)
    const third = xml.indexOf('<w:p>', second + 1)
    xml =
      xml.slice(0, second) +
      '<w:sdt><w:sdtContent>' +
      xml.slice(second, third) +
      '</w:sdtContent></w:sdt>' +
      xml.slice(third)
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml, numberingXml }))
    const paras = parsed.blocks.flatMap((b) => b.textboxes ?? [])[0]!.paras
    expect(paras.map((p) => p.listMarker?.text)).toEqual(['1.', '2.', '3.'])
  })

  it('numbers ordered items in document order, skipping plain paragraphs', async () => {
    const paras = await textboxParas(['2', null, '2', '2'])
    expect(paras.map((p) => p.listMarker?.text ?? null)).toEqual(['1.', null, '2.', '3.'])
    expect(paras[0].listMarker!.symbol).toBeUndefined()
  })
})

import { describe, expect, it } from 'vitest'
import { mergeCellParaFormat } from '../src/generate'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

async function parseFirst(pPrXml: string, extraStylesXml?: string) {
  const bytes = await buildDocx({
    bodyXml: `<w:p>${pPrXml ? `<w:pPr>${pPrXml}</w:pPr>` : ''}<w:r><w:t>x</w:t></w:r></w:p>`,
    extraStylesXml,
  })
  return (await parseDocx(bytes)).blocks[0]
}

const NO_HANG =
  '<w:style w:type="paragraph" w:styleId="NoHang"><w:name w:val="No Hang"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:overflowPunct w:val="0"/></w:pPr></w:style>'

describe('Asian typography flags as written on the paragraph', () => {
  it('reads kinsoku, topLinePunct and textAlignment off the paragraph', async () => {
    const block = await parseFirst(
      '<w:kinsoku w:val="0"/><w:topLinePunct/><w:textAlignment w:val="top"/>',
    )
    expect(block.format).toMatchObject({ kinsoku: false, topLinePunct: true, textAlignment: 'top' })
  })

  it('tells a direct w:overflowPunct from one the style chain turned off', async () => {
    const direct = await parseFirst('<w:overflowPunct w:val="0"/>')
    expect(direct.format).toMatchObject({ overflowPunct: false, overflowPunctDirect: false })
    const inherited = await parseFirst('<w:pStyle w:val="NoHang"/>', NO_HANG)
    expect(inherited.format?.overflowPunct).toBe(false)
    expect(inherited.format?.overflowPunctDirect).toBeUndefined()
  })

  it('an inherited off is never written into the paragraph', async () => {
    const block = await parseFirst('<w:pStyle w:val="NoHang"/><w:jc w:val="center"/>', NO_HANG)
    expect(mergePPrFormat(block.rawPPr!, { ...block.format, align: 'left' }, block.format)).toBe(
      '<w:pPr><w:pStyle w:val="NoHang"/><w:jc w:val="left"/></w:pPr>',
    )
  })

  it('writes the three flags between w:tabs and w:spacing, in schema order', () => {
    const raw =
      '<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:spacing w:after="0"/></w:pPr>'
    expect(
      mergePPrFormat(raw, { autoSpaceDN: false, overflowPunctDirect: false, autoSpaceDE: true }),
    ).toBe(
      '<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:overflowPunct w:val="0"/>' +
        '<w:autoSpaceDE/><w:autoSpaceDN w:val="0"/></w:pPr>',
    )
  })

  it('a header or footer cell paragraph takes a changed flag and keeps the rest', () => {
    const raw = '<w:pPr><w:keepNext/><w:ind w:right="240"/></w:pPr>'
    expect(mergeCellParaFormat(raw, { autoSpaceDE: false }, {})).toBe(
      '<w:pPr><w:keepNext/><w:autoSpaceDE w:val="0"/><w:ind w:right="240"/></w:pPr>',
    )
    expect(mergeCellParaFormat(raw, { autoSpaceDE: false }, { autoSpaceDE: false })).toBe(raw)
  })
})

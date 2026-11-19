import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const OFF = '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
const STYLES_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** Normal carrying the Korean HWP-export defaults: character-level breaking + autospace off */
const KR_STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${STYLES_NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>' +
  `<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/>${OFF}</w:pPr></w:style>` +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="a"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Loose"><w:name w:val="Loose"/></w:style>' +
  '</w:styles>'

async function parseFirst(pPrXml: string, stylesXml?: string) {
  const bytes = await buildDocx({
    bodyXml: `<w:p>${pPrXml ? `<w:pPr>${pPrXml}</w:pPr>` : ''}<w:r><w:t>가맹점</w:t></w:r></w:p>`,
    stylesXml,
  })
  return (await parseDocx(bytes)).blocks[0]
}

describe('w:wordWrap parsing', () => {
  it('absent means Word default (word wrapping): no field', async () => {
    expect((await parseFirst('')).format?.wordWrap).toBeUndefined()
  })

  it('w:val="0" parses to false, bare element to true', async () => {
    expect((await parseFirst('<w:wordWrap w:val="0"/>')).format?.wordWrap).toBe(false)
    expect((await parseFirst('<w:wordWrap/>')).format?.wordWrap).toBe(true)
  })
})

describe('default-style inheritance (Word probe 2026-09-03: pStyle-less paragraphs follow Normal)', () => {
  it('a pStyle-less paragraph inherits wordWrap and autoSpace off from the default style', async () => {
    const block = await parseFirst('', KR_STYLES)
    expect(block.format?.wordWrap).toBe(false)
    expect(block.format?.autoSpace).toBe(false)
  })

  it('inherits through basedOn but not into unrelated styles', async () => {
    const body = await parseFirst('<w:pStyle w:val="Body"/>', KR_STYLES)
    expect(body.format?.wordWrap).toBe(false)
    expect(body.format?.autoSpace).toBe(false)
    const loose = await parseFirst('<w:pStyle w:val="Loose"/>', KR_STYLES)
    expect(loose.format?.wordWrap).toBeUndefined()
    expect(loose.format?.autoSpace).toBeUndefined()
  })

  it('an explicit on in the pPr overrides the style-level off', async () => {
    const block = await parseFirst('<w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/>', KR_STYLES)
    expect(block.format?.wordWrap).toBe(true)
    expect(block.format?.autoSpace).toBe(true)
  })

  it('table cell paragraphs inherit the default style flags too', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
        '<w:p><w:r><w:t>가맹점 PG사</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      stylesXml: KR_STYLES,
    })
    const table = (await parseDocx(bytes)).blocks.find((b) => b.type === 'table')
    const para = table?.table?.rows[0][0].richParas?.[0]
    expect(para?.wordWrap).toBe(false)
    expect(para?.autoSpace).toBe(false)
  })
})

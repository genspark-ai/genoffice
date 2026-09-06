/**
 * Word truth (probe 2026-09-05, MS Mincho 10.5pt, doNotCompress and
 * compressPunctuation): kinsoku, hanging line-end punctuation and punctuation
 * compression apply only when the run's effective w:lang w:eastAsia is a CJK
 * language. Precedence: run rPr > character style > paragraph style chain >
 * docDefaults; the paragraph mark's rPr does not count; no w:lang anywhere
 * keeps the East Asian rules.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="JP"><w:name w:val="JP"/><w:basedOn w:val="a"/>' +
  '<w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="JPChild"><w:name w:val="JPChild"/><w:basedOn w:val="JP"/></w:style>' +
  '<w:style w:type="character" w:styleId="ZH"><w:name w:val="ZH"/><w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr></w:style>' +
  '</w:styles>'

async function parse(bodyXml: string, stylesXml = STYLES) {
  return parseDocx(await buildDocx({ bodyXml, stylesXml }))
}

describe('effective w:lang w:eastAsia', () => {
  it('reads docDefaults and resolves it through the style basedOn chain', async () => {
    const doc = await parse('<w:p><w:r><w:t>x</w:t></w:r></w:p>')
    expect(doc.docDefaults?.eastAsiaLang).toBe('en-US')
    expect(doc.styles.get('JP')?.display?.eastAsiaLang).toBe('ja-JP')
    expect(doc.styles.get('JPChild')?.display?.eastAsiaLang).toBe('ja-JP')
    expect(doc.styles.get('a')?.display?.eastAsiaLang).toBeUndefined()
  })

  it('paragraphs inherit the style chain value; the paragraph mark rPr does not count', async () => {
    const doc = await parse(
      '<w:p><w:pPr><w:pStyle w:val="JPChild"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:pPr><w:r><w:t>y</w:t></w:r></w:p>',
    )
    expect(doc.blocks[0].format?.eastAsiaLang).toBe('ja-JP')
    expect(doc.blocks[1].format?.eastAsiaLang).toBeUndefined()
    expect(doc.blocks[1].runs?.[0].eastAsiaLang).toBeUndefined()
  })

  it('runs carry their own value or the character style value', async () => {
    const doc = await parse(
      '<w:p><w:pPr><w:pStyle w:val="JP"/></w:pPr>' +
        '<w:r><w:rPr><w:lang w:eastAsia="en-US"/></w:rPr><w:t>a</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="ZH"/></w:rPr><w:t>b</w:t></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>c</w:t></w:r>' +
        '</w:p>',
    )
    const runs = doc.blocks[0].runs!
    expect(runs[0].eastAsiaLang).toBe('en-US')
    expect(runs[1].eastAsiaLang).toBe('zh-CN')
    expect(runs[2].eastAsiaLang).toBeUndefined()
  })

  it('leaves everything undefined when no w:lang exists', async () => {
    const doc = await parse(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>`,
    )
    expect(doc.docDefaults?.eastAsiaLang).toBeUndefined()
    expect(doc.blocks[0].format?.eastAsiaLang).toBeUndefined()
  })
})

/**
 * Word reads a fractional w:line (Google Docs exports "278.00000000000006")
 * and truncates it (Word for Mac probe 2026-09-24: 278.5 and 278.9 lay out
 * as 278, docDefaults and style values alike); dropping the value collapsed
 * a 1.16-spaced document to single spacing.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const para = (spacing: string) =>
  `<w:p><w:pPr><w:spacing ${spacing}/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`

async function firstPara(spacing: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: para(spacing) }))
  return parsed.blocks[0].format ?? {}
}

describe('fractional twips in w:spacing', () => {
  it('truncates a fractional w:line like Word', async () => {
    expect(await firstPara('w:line="278.00000000000006" w:lineRule="auto"')).toMatchObject({
      lineRawTwips: 278,
      lineSpacing: 1.16,
      lineRule: 'auto',
    })
    expect(await firstPara('w:line="278.9" w:lineRule="auto"')).toMatchObject({
      lineRawTwips: 278,
    })
    expect(await firstPara('w:line="300.5" w:lineRule="exact"')).toMatchObject({
      lineRawTwips: 300,
      lineRule: 'exact',
    })
  })

  it('still ignores a non-numeric w:line', async () => {
    const f = await firstPara('w:line="abc" w:lineRule="auto"')
    expect(f.lineRawTwips).toBeUndefined()
    expect(f.lineSpacing).toBeUndefined()
  })

  it('truncates fractional before/after too', async () => {
    expect(await firstPara('w:before="600.5" w:after="600.9"')).toMatchObject({
      spaceBefore: 600,
      spaceAfter: 600,
    })
  })

  it('reads a fractional w:line from docDefaults and from a style', async () => {
    const stylesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
      '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278.00000000000006" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Fl"><w:name w:val="Fl"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:line="300.5" w:lineRule="auto"/></w:pPr></w:style></w:styles>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:pPr><w:pStyle w:val="Fl"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
        stylesXml,
      }),
    )
    expect(parsed.docDefaults).toMatchObject({ lineRawTwips: 278, lineRule: 'auto' })
    expect(parsed.docDefaults?.lineSpacing).toBeCloseTo(278 / 240, 5)
    expect(parsed.styles.get('Fl')?.display).toMatchObject({ lineRawTwips: 300, lineSpacing: 1.25 })
  })
})

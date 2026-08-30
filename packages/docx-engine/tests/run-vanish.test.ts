import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:styleId="HiddenPara"><w:name w:val="HiddenPara"/>' +
  '<w:rPr><w:vanish/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="HiddenChar"><w:name w:val="HiddenChar"/>' +
  '<w:rPr><w:vanish/></w:rPr></w:style>' +
  '</w:styles>'

const DEFAULT_VANISH_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normlny"><w:name w:val="Normlny"/>' +
  '<w:rPr><w:vanish/></w:rPr></w:style>' +
  '</w:styles>'

const runsOf = async (bodyXml: string) => {
  const doc = await parseDocx(await buildDocx({ bodyXml, stylesXml: STYLES }))
  return doc.blocks[0].runs!
}

describe('run-level hidden text (w:vanish)', () => {
  it('an explicit run vanish hides the run', async () => {
    const runs = await runsOf(
      '<w:p><w:r><w:t>shown</w:t></w:r>' +
        '<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>',
    )
    expect(runs.map((r) => [r.text, r.vanish ?? false])).toEqual([
      ['shown', false],
      ['hidden', true],
    ])
  })

  it('w:specVanish (style separator) stays visible', async () => {
    const runs = await runsOf(
      '<w:p><w:r><w:rPr><w:vanish/><w:specVanish/></w:rPr><w:t>separator</w:t></w:r></w:p>',
    )
    expect(runs[0].vanish).toBeUndefined()
  })

  it('a paragraph style vanish inherits into runs; an explicit off un-hides', async () => {
    const runs = await runsOf(
      '<w:p><w:pPr><w:pStyle w:val="HiddenPara"/></w:pPr>' +
        '<w:r><w:t>inherited hidden</w:t></w:r>' +
        '<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>unhidden</w:t></w:r></w:p>',
    )
    expect(runs.map((r) => [r.text, r.vanish ?? false])).toEqual([
      ['inherited hidden', true],
      ['unhidden', false],
    ])
  })

  it('a vanish-carrying default paragraph style hides style-less runs (form docs)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>inherited hidden</w:t></w:r>' +
          '<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>unhidden</w:t></w:r></w:p>',
        stylesXml: DEFAULT_VANISH_STYLES,
      }),
    )
    expect(doc.blocks[0].runs!.map((r) => [r.text, r.vanish ?? false])).toEqual([
      ['inherited hidden', true],
      ['unhidden', false],
    ])
  })

  it('a character style vanish inherits into its runs', async () => {
    const runs = await runsOf(
      '<w:p><w:r><w:rPr><w:rStyle w:val="HiddenChar"/></w:rPr><w:t>char hidden</w:t></w:r>' +
        '<w:r><w:t>plain</w:t></w:r></w:p>',
    )
    expect(runs.map((r) => [r.text, r.vanish ?? false])).toEqual([
      ['char hidden', true],
      ['plain', false],
    ])
  })
})

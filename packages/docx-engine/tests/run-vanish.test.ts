import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { computeListMarkers } from '../src/list-markers'
import { staysVanished } from '../src/parse-props'
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

describe('staysVanished falsy variants', () => {
  it('treats none, uppercase, and single-quoted falsy vals as visible', () => {
    expect(staysVanished('<w:p><w:rPr><w:vanish w:val="none"/></w:rPr></w:p>')).toBe(false)
    expect(staysVanished('<w:p><w:rPr><w:vanish w:val="False"/></w:rPr></w:p>')).toBe(false)
    expect(staysVanished("<w:p><w:rPr><w:vanish w:val='off'/></w:rPr></w:p>")).toBe(false)
    expect(staysVanished('<w:p><w:rPr><w:vanish/></w:rPr></w:p>')).toBe(true)
  })
})

describe('hidden empty numbered paragraph', () => {
  it('collapses to an invisible block that still consumes its list number', async () => {
    const hidden =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' +
      '<w:rPr><w:vanish/></w:rPr></w:pPr></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          hidden +
          hidden +
          '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>' +
          '<w:r><w:t>third</w:t></w:r></w:p>',
        numberingXml:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
          '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
          '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>',
      }),
    )
    const blocks = doc.blocks.slice(0, 3)
    expect(blocks.map((b) => [b.type, b.invisibleMarker ?? false])).toEqual([
      ['passthrough', true],
      ['passthrough', true],
      ['listItem', false],
    ])
    expect(blocks[0].strayList).toEqual({ numId: '2', ilvl: 0 })
    const refs = blocks.map((b) => b.list ?? b.strayList!)
    expect(computeListMarkers(refs, doc.numbering)).toEqual(['1.', '2.', '3.'])
  })
})

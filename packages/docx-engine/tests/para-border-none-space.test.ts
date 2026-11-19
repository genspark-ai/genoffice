import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** docDefaults pBdr with none sides that keep a w:space (HTML-import generators write this) */
const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr><w:pBdr>` +
  '<w:top w:val="none" w:sz="0" w:space="30" w:color="000000"/>' +
  '<w:left w:val="none" w:sz="0" w:space="31" w:color="000000"/>' +
  '<w:bottom w:val="none" w:sz="0" w:space="30" w:color="000000"/>' +
  '<w:right w:val="none" w:sz="0" w:space="31" w:color="000000"/>' +
  '</w:pBdr></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Web"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pBdr><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="H2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="3" w:color="EEEEEE"/></w:pBdr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '</w:styles>'

const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:hdr ${W}>` +
  '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="0" w:color="00AEEF"/></w:pBdr></w:pPr><w:r><w:t>Brand</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="0" w:color="00AEEF"/></w:pBdr></w:pPr></w:p>' +
  '</w:hdr>'

async function build(bodyXml: string) {
  return parseDocx(
    await buildDocx({
      bodyXml,
      stylesXml: STYLES_XML,
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER_XML,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    }),
  )
}

describe('w:pBdr none sides keep their w:space as paragraph padding', () => {
  it('roots every paragraph style at the docDefaults pBdr; a space=0 override removes it', async () => {
    const doc = await build('<w:p><w:r><w:t>x</w:t></w:r></w:p>')
    expect(doc.docDefaults?.borderSides).toEqual({
      t: { none: true, spacePt: 30 },
      l: { none: true, spacePt: 31 },
      b: { none: true, spacePt: 30 },
      r: { none: true, spacePt: 31 },
    })
    expect(doc.styles.get('Normal')!.display?.borderSides).toEqual(doc.docDefaults?.borderSides)
    expect(doc.styles.get('Web')!.display?.borderSides).toEqual({
      t: null,
      l: null,
      b: null,
      r: null,
    })
    // the drawn side keeps its own space, the inherited none sides keep theirs
    expect(doc.styles.get('H2')!.display?.borderSides).toEqual({
      t: { none: true, spacePt: 30 },
      l: { none: true, spacePt: 31 },
      b: { color: 'EEEEEE', szPt: 0.75, spacePt: 3 },
      r: { none: true, spacePt: 31 },
    })
  })

  it('direct none sides: reset stays, the space becomes borderPad (display-only)', async () => {
    const doc = await build(
      '<w:p><w:pPr><w:pBdr><w:top w:val="none" w:sz="0" w:space="12" w:color="auto"/>' +
        '<w:bottom w:val="nil"/></w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    )
    const [p] = doc.blocks.filter((b) => !b.hidden)
    expect(p.format?.borders).toBeUndefined()
    expect(p.format?.borderReset).toBe('tb')
    expect(p.format?.borderPad).toEqual({ t: 12 })
  })

  it('header paragraphs resolve the docDefaults padding and keep their own drawn side', async () => {
    const doc = await build('<w:p><w:r><w:t>x</w:t></w:r></w:p>')
    const paras = doc.headerParas!
    expect(paras).toHaveLength(2)
    for (const para of paras) {
      expect(para.borders).toBe('b')
      expect(para.borderLines).toEqual({ b: { color: '00AEEF', szPt: 0.5 } })
      expect(para.borderPad).toEqual({ t: 30, l: 31, r: 31 })
    }
  })
})

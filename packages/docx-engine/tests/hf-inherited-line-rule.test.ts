import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * A header paragraph's line rule should survive wherever it was declared:
 * directly, through a named style, through the implicit Normal, or through the
 * document defaults. The strip base only carries an automatic multiplier, so an
 * inherited exact/atLeast rule has to travel inline like a direct one.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const HEADER = '<w:p><w:r><w:t>Header</w:t></w:r></w:p>'

const NORMAL =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'

/** docDefaults pPr and, when given, the default paragraph style's own pPr */
function styles(docDefaultsPPr: string, normalPPr?: string): string {
  const normal =
    normalPPr === undefined
      ? NORMAL
      : '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
        `<w:pPr>${normalPPr}</w:pPr></w:style>`
  return (
    `${XML_DECL}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:pPrDefault><w:pPr>${docDefaultsPPr}</w:pPr></w:pPrDefault></w:docDefaults>` +
    normal +
    '</w:styles>'
  )
}

async function headerDoc(stylesXml: string, headerInner: string) {
  return parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
      stylesXml,
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: `${XML_DECL}<w:hdr ${NS}>${headerInner}</w:hdr>`,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    }),
  )
}

/** the header's single flow paragraph */
async function headerPara(stylesXml: string, headerInner = HEADER) {
  const doc = await headerDoc(stylesXml, headerInner)
  return doc.headerParas![0]!
}

describe('header line spacing inherited from the document defaults', () => {
  it('keeps an exact rule and its twips', async () => {
    const para = await headerPara(styles('<w:spacing w:line="280" w:lineRule="exact"/>'))
    expect(para).toMatchObject({ lineRule: 'exact', lineRawTwips: 280 })
    // an absolute rule is not a multiplier: the auto-facing keys stay absent
    expect(para.lineSpacing).toBeUndefined()
  })

  it('keeps an atLeast rule and its twips', async () => {
    const para = await headerPara(styles('<w:spacing w:line="280" w:lineRule="atLeast"/>'))
    expect(para).toMatchObject({ lineRule: 'atLeast', lineRawTwips: 280 })
  })

  it('leaves an auto rule to the strip base', async () => {
    // the base carries the automatic multiplier already; emitting the same
    // multiplier inline would only restate it
    const para = await headerPara(styles('<w:spacing w:line="480" w:lineRule="auto"/>'))
    expect(para.lineRule).toBeUndefined()
    expect(para.lineRawTwips).toBeUndefined()
  })
})

describe('header line spacing inherited from the implicit Normal', () => {
  it('keeps an exact rule the paragraph never names', async () => {
    // the paragraph carries no w:pStyle, yet Normal is the style it resolves to
    const para = await headerPara(styles('', '<w:spacing w:line="280" w:lineRule="exact"/>'))
    expect(para).toMatchObject({ lineRule: 'exact', lineRawTwips: 280 })
  })

  it('keeps an atLeast rule the paragraph never names', async () => {
    const para = await headerPara(styles('', '<w:spacing w:line="300" w:lineRule="atLeast"/>'))
    expect(para).toMatchObject({ lineRule: 'atLeast', lineRawTwips: 300 })
  })
})

describe('header line spacing declared directly', () => {
  it('still reaches the paragraph', async () => {
    const para = await headerPara(
      styles(''),
      '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>',
    )
    expect(para).toMatchObject({ lineRule: 'exact', lineRawTwips: 360 })
  })
})

describe('header table cell line spacing', () => {
  it('keeps an exact rule inherited from the document defaults', async () => {
    // cells resolve their line through the same source list
    const row =
      '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'
    const doc = await headerDoc(styles('<w:spacing w:line="280" w:lineRule="exact"/>'), row)
    const cell = doc.headerParas!.find((p) => p.cells)!.cells![0]
    expect(cell.paraProps![0]).toMatchObject({ lineRule: 'exact', lineRawTwips: 280 })
  })
})

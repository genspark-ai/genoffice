import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const BODY = '<w:p><w:r><w:t>Generated body text</w:t></w:r></w:p>'

const THEME_PART = {
  path: 'word/theme/theme1.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
    '<a:themeElements><a:fontScheme name="T">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>',
  contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
}

function stylesXml(rPrDefault: string) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
    `<w:docDefaults>${rPrDefault}<w:pPrDefault/></w:docDefaults>` +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
    '<w:rPr><w:sz w:val="56"/></w:rPr></w:style></w:styles>'
  )
}

describe('docDefaults without any Latin font slot', () => {
  it('empty rPrDefault, no theme: Word falls back to Times New Roman', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: BODY, stylesXml: stylesXml('<w:rPrDefault/>') }),
    )
    expect(doc.docDefaults?.asciiFont).toBe('Times New Roman')
    expect(doc.docDefaults?.sizeHalfPoints).toBe(20)
  })

  it('an explicit docDefaults size wins over the 10 pt built-in', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: stylesXml('<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'),
      }),
    )
    expect(doc.docDefaults?.asciiFont).toBe('Times New Roman')
    expect(doc.docDefaults?.sizeHalfPoints).toBe(24)
  })

  it('rPr without rFonts, no theme: same fallback', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: stylesXml('<w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'),
      }),
    )
    expect(doc.docDefaults?.asciiFont).toBe('Times New Roman')
  })

  it('an explicit ascii face wins', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: stylesXml(
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault>',
        ),
      }),
    )
    expect(doc.docDefaults?.asciiFont).toBe('Arial')
  })

  it('a theme part keeps the current resolution (no Times fallback)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: stylesXml('<w:rPrDefault/>'),
        extraParts: [THEME_PART],
      }),
    )
    expect(doc.docDefaults?.asciiFont).toBeUndefined()
    expect(doc.docDefaults?.sizeHalfPoints).toBeUndefined()
  })

  it('theme reference without a theme part is left unresolved', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: stylesXml(
          '<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault>',
        ),
      }),
    )
    expect(doc.docDefaults?.asciiFont).toBeUndefined()
  })
})

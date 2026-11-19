import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

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

const NORMAL_ARIAL =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:style>'

function footerPart(rFonts: string) {
  return {
    path: 'word/footer1.xml',
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:p><w:r><w:rPr>${rFonts}</w:rPr><w:t>Version 1.0</w:t></w:r></w:p>` +
      '</w:ftr>',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  }
}

async function parseWithFooter(rFonts: string, withTheme = true) {
  const bytes = await buildDocx({
    bodyXml: `<w:p><w:r><w:rPr>${rFonts}</w:rPr><w:t>body</w:t></w:r></w:p>`,
    extraStylesXml: NORMAL_ARIAL,
    extraRels:
      '<Relationship Id="rId62" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: withTheme ? [footerPart(rFonts), THEME_PART] : [footerPart(rFonts)],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId62"/>',
  })
  return parseDocx(bytes)
}

describe('header/footer theme font references', () => {
  const THEMED =
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsia="Arial"/>'

  it('resolves w:asciiTheme in footer runs exactly like body runs', async () => {
    const doc = await parseWithFooter(THEMED)
    const body = doc.blocks[0].runs![0]
    const footer = doc.footerParas![0].runs[0]
    expect(body.fontAscii).toBe('Calibri')
    expect(footer.fontAscii).toBe('Calibri')
    expect(footer.font).toBe('Arial')
    expect(footer.themeRFonts).toEqual({ fontAscii: 'Calibri' })
  })

  it('leaves the style chain in charge when no theme part exists', async () => {
    const doc = await parseWithFooter(THEMED, false)
    expect(doc.footerParas![0].runs[0].fontAscii).toBeUndefined()
  })

  it('keeps explicit footer fonts untouched', async () => {
    const doc = await parseWithFooter('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>')
    expect(doc.footerParas![0].runs[0].fontAscii).toBe('Georgia')
    expect(doc.footerParas![0].runs[0].themeRFonts).toBeUndefined()
  })
})

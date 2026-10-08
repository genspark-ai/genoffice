import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { buildDocx } from './helpers/build-docx'
import { parseDocx } from '../src/index'

const THEME = {
  path: 'word/theme/theme1.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
    '<a:themeElements><a:fontScheme name="T">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>' +
    '<a:font script="Arab" typeface="Arial"/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>',
  contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
}

const SETTINGS = {
  path: 'word/settings.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:themeFontLang w:val="en-US" w:bidi="ar-SA"/></w:settings>',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
}

const ARABIC = 'كريم حسن (معلم)'

async function parseWithDocDefaults(bodyXml: string, rPrDefault: string) {
  const zip = await JSZip.loadAsync(await buildDocx({ bodyXml, extraParts: [THEME, SETTINGS] }))
  const stylesXml = await zip.file('word/styles.xml')!.async('string')
  zip.file(
    'word/styles.xml',
    stylesXml.replace(
      /(<w:styles[^>]*>)/,
      `$1<w:docDefaults><w:rPrDefault><w:rPr>${rPrDefault}</w:rPr></w:rPrDefault></w:docDefaults>`,
    ),
  )
  return parseDocx(await zip.generateAsync({ type: 'uint8array' }))
}

const DD_RFONTS =
  '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'

describe('complex-script runs without a cs font', () => {
  it('take the docDefaults cs face, an empty theme slot resolving through the script table', async () => {
    const doc = await parseWithDocDefaults(
      `<w:p><w:r><w:rPr><w:rFonts w:hint="cs"/><w:b/><w:rtl/></w:rPr><w:t>${ARABIC}</w:t></w:r>` +
        '<w:r><w:rPr><w:rFonts w:hint="cs"/><w:rtl/></w:rPr><w:t>J1</w:t></w:r></w:p>',
      DD_RFONTS,
    )
    expect(doc.docDefaults?.csFont).toBe('Arial')
    const runs = doc.blocks[0].runs!
    expect(runs[0].csFont).toBe('Arial')
    expect(runs[0].fontCs).toBeUndefined()
    expect(runs[1].csFont).toBeUndefined()
  })

  it('keep an explicit run cs font over the document default', async () => {
    const doc = await parseWithDocDefaults(
      `<w:p><w:r><w:rPr><w:rFonts w:cs="Times New Roman" w:hint="cs"/><w:rtl/></w:rPr><w:t>${ARABIC}</w:t></w:r></w:p>`,
      DD_RFONTS,
    )
    expect(doc.blocks[0].runs![0].csFont).toBe('Times New Roman')
  })

  it('prefer the paragraph style cs face over docDefaults', async () => {
    const zip = await JSZip.loadAsync(
      await buildDocx({
        bodyXml: `<w:p><w:pPr><w:pStyle w:val="Ar"/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>${ARABIC}</w:t></w:r></w:p>`,
        extraParts: [THEME, SETTINGS],
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Ar"><w:name w:val="Ar"/><w:rPr><w:rFonts w:cs="Sakkal Majalla"/></w:rPr></w:style>',
      }),
    )
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    zip.file(
      'word/styles.xml',
      stylesXml.replace(
        /(<w:styles[^>]*>)/,
        `$1<w:docDefaults><w:rPrDefault><w:rPr>${DD_RFONTS}</w:rPr></w:rPrDefault></w:docDefaults>`,
      ),
    )
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.blocks[0].runs![0].csFont).toBe('Sakkal Majalla')
  })
  it('leave a run without w:rtl on its ascii/hAnsi face even when the text is Arabic', async () => {
    const doc = await parseWithDocDefaults(
      `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic"/><w:sz w:val="31"/></w:rPr><w:t>${ARABIC}</w:t></w:r>` +
        `<w:r><w:t>${ARABIC}</w:t></w:r></w:p>`,
      DD_RFONTS,
    )
    const runs = doc.blocks[0].runs!
    expect(runs[0].csFont).toBeUndefined()
    expect(runs[0].font).toBe('Traditional Arabic')
    expect(runs[0].sizeHalfPoints).toBe(31)
    expect(runs[1].csFont).toBeUndefined()
  })
})

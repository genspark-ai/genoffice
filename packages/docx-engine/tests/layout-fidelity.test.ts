import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

const settingsPart = (inner: string) => ({
  path: 'word/settings.xml',
  xml:
    XML_DECL +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:settings>`,
  contentType: SETTINGS_CT,
})

describe('w:t whitespace handling', () => {
  it('trims leading/trailing XML whitespace unless xml:space="preserve"', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>\n\t\t\tL1: \n\t\t</w:t></w:r>' +
        '<w:r><w:t>\n  \t\tcancelled\n\t</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> kept </w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('L1:cancelled kept ')
  })

  it('xml:space="preserve" on the document root covers every w:t (inherited XML scope)', async () => {
    const bytes = await buildDocx({
      docRootExtraAttrs: 'xml:space="preserve"',
      bodyXml:
        '<w:p><w:r><w:t>Memorandum</w:t></w:r>' +
        '<w:r><w:rPr><w:spacing w:val="-49"/></w:rPr><w:t> </w:t></w:r>' +
        '<w:r><w:t>of</w:t></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="default">\n\t x \n</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    // whitespace-only run survives; an explicit xml:space="default" still trims
    expect(doc.blocks[0].runs!.map((r) => r.text)).toEqual(['Memorandum', ' ', 'of', 'x'])
  })

  it('xml:space="preserve" on a footer part root covers its w:t too', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId70" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [
        {
          path: 'word/footer1.xml',
          xml:
            XML_DECL +
            '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve">' +
            '<w:p><w:r><w:t>END</w:t></w:r>' +
            '<w:r><w:rPr><w:spacing w:val="-4"/></w:rPr><w:t> </w:t></w:r>' +
            '<w:r><w:t>OF</w:t></w:r></w:p></w:ftr>',
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
        },
      ],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId70"/>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.footerParas![0].runs.map((r) => r.text)).toEqual(['END', ' ', 'OF'])
  })
})

describe('hex color tolerance', () => {
  it("accepts a leading '#' on w:color and w:shd fills (tdf#57589)", async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:shd w:val="clear" w:fill="#e6e6e6"/></w:pPr>' +
        '<w:r><w:rPr><w:color w:val="#004080"/><w:shd w:val="clear" w:fill="#ff00ff"/></w:rPr>' +
        '<w:t>colored</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format?.shadingFill).toBe('e6e6e6')
    expect(doc.blocks[0].runs![0].color).toBe('004080')
    expect(doc.blocks[0].runs![0].shading).toBe('ff00ff')
  })
})

describe('default style resolution', () => {
  it('the last w:default="1" paragraph style wins', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="paragraph" w:default="1" w:styleId="Title"><w:name w:val="Title"/>' +
        '<w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>',
    })
    const doc = await parseDocx(bytes)
    const defaults = [...doc.styles.values()].filter((s) => s.isDefault && s.type === 'paragraph')
    expect(defaults.map((s) => s.styleId)).toEqual(['Title'])
  })

  it('without any w:default, a style id/named "Normal" is the default (not the first style)', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      stylesXml:
        XML_DECL +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>' +
        '<w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Style7"><w:name w:val="normal"/>' +
        '<w:rPr><w:color w:val="FF00FF"/><w:sz w:val="48"/></w:rPr></w:style>' +
        '</w:styles>',
    })
    const doc = await parseDocx(bytes)
    const defaults = [...doc.styles.values()].filter((s) => s.isDefault && s.type === 'paragraph')
    expect(defaults.map((s) => s.styleId)).toEqual(['Style7'])
    expect(defaults[0].display?.color).toBe('FF00FF')
  })

  it('without any w:default or Normal style, no paragraph style is the default', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      stylesXml:
        XML_DECL +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>' +
        '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="MyStyle"><w:name w:val="MyStyle"/></w:style>' +
        '</w:styles>',
    })
    const doc = await parseDocx(bytes)
    const defaults = [...doc.styles.values()].filter((s) => s.isDefault)
    expect(defaults).toEqual([])
  })
})

describe('settings.xml layout flags', () => {
  it('parses w:autoHyphenation and w:defaultTabStop (including 0)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [settingsPart('<w:autoHyphenation/><w:defaultTabStop w:val="0"/>')],
      }),
    )
    expect(doc.autoHyphenation).toBe(true)
    expect(doc.defaultTabStopTwips).toBe(0)
  })

  it('leaves both unset without a settings part', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.autoHyphenation).toBeUndefined()
    expect(doc.defaultTabStopTwips).toBeUndefined()
  })

  it('w:autoHyphenation w:val="false" counts as off', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [
          settingsPart('<w:autoHyphenation w:val="false"/><w:defaultTabStop w:val="709"/>'),
        ],
      }),
    )
    expect(doc.autoHyphenation).toBeUndefined()
    expect(doc.defaultTabStopTwips).toBe(709)
  })
})

describe('w:suppressAutoHyphens', () => {
  const stylesWith = (docDefaultsPPr: string, styleExtra = '') =>
    XML_DECL +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:docDefaults><w:pPrDefault><w:pPr>${docDefaultsPPr}</w:pPr></w:pPrDefault></w:docDefaults>` +
    styleExtra +
    '</w:styles>'

  it('pPrDefault w:suppressAutoHyphens lands on docDefaults', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [settingsPart('<w:autoHyphenation/>')],
        stylesXml: stylesWith('<w:suppressAutoHyphens/>'),
      }),
    )
    expect(doc.autoHyphenation).toBe(true)
    expect(doc.docDefaults?.suppressAutoHyphens).toBe(true)
  })

  it('explicit off at pPrDefault stays unset', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        stylesXml: stylesWith('<w:suppressAutoHyphens w:val="0"/>'),
      }),
    )
    expect(doc.docDefaults?.suppressAutoHyphens).toBeUndefined()
  })

  it('style-level value is tri-state and survives the basedOn merge', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        stylesXml: stylesWith(
          '',
          '<w:style w:type="paragraph" w:styleId="NoHyph"><w:name w:val="NoHyph"/>' +
            '<w:pPr><w:suppressAutoHyphens/></w:pPr></w:style>' +
            '<w:style w:type="paragraph" w:styleId="ReHyph"><w:name w:val="ReHyph"/>' +
            '<w:basedOn w:val="NoHyph"/>' +
            '<w:pPr><w:suppressAutoHyphens w:val="false"/></w:pPr></w:style>',
        ),
      }),
    )
    expect(doc.styles.get('NoHyph')?.display?.suppressAutoHyphens).toBe(true)
    expect(doc.styles.get('ReHyph')?.display?.suppressAutoHyphens).toBe(false)
  })
})

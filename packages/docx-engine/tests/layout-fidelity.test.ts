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

  it('without any w:default, the first style of the type is the default', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      stylesXml:
        XML_DECL +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>' +
        '<w:rPr><w:color w:val="FF00FF"/><w:sz w:val="48"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="MyStyle"><w:name w:val="MyStyle"/></w:style>' +
        '</w:styles>',
    })
    const doc = await parseDocx(bytes)
    const defaults = [...doc.styles.values()].filter((s) => s.isDefault && s.type === 'paragraph')
    expect(defaults.map((s) => s.styleId)).toEqual(['Normal'])
    expect(defaults[0].display?.color).toBe('FF00FF')
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

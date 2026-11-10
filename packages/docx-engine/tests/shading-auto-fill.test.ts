/**
 * w:shd w:fill="auto" is Word's "no shading": at paragraph level it cancels the
 * shading inherited from the paragraph style instead of falling through to it.
 * w:themeFill is re-resolved against the theme (the w:fill hex is only a cache).
 */
import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<w:style w:type="paragraph" w:styleId="Boxed"><w:name w:val="Boxed"/>' +
  '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="B8CCE4"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Unboxed"><w:name w:val="Unboxed"/>' +
  '<w:basedOn w:val="Boxed"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="StillBoxed"><w:name w:val="Still Boxed"/>' +
  '<w:basedOn w:val="Boxed"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>'

const AUTO_SHD = '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'

describe('paragraph w:shd fill=auto', () => {
  it('an explicit auto on the paragraph clears the style fill', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:pPr><w:pStyle w:val="Boxed"/>${AUTO_SHD}</w:pPr><w:r><w:t>clear</w:t></w:r></w:p>` +
          '<w:p><w:pPr><w:pStyle w:val="Boxed"/></w:pPr><w:r><w:t>inherit</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
      }),
    )
    const [cleared, inherited] = doc.blocks
    expect(cleared.format?.shadingFill).toBeUndefined()
    expect(cleared.format?.shadingDisplay).toBeUndefined()
    expect(cleared.format?.shadingClear).toBe(true)
    expect(inherited.format?.shadingClear).toBeUndefined()
    // the raw w:shd round-trips untouched
    expect(mergePPrFormat(cleared.rawPPr!, cleared.format!)).toContain(AUTO_SHD)
  })

  it('a style-level auto yields no fill and cancels the basedOn fill', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>', extraStylesXml: STYLES }),
    )
    expect(doc.styles.get('Boxed')?.display?.shadingFill).toBe('B8CCE4')
    expect(doc.styles.get('Unboxed')?.display?.shadingFill).toBe('auto')
    expect(doc.styles.get('StillBoxed')?.display?.shadingFill).toBe('B8CCE4')
  })

  it('a paragraph-mark rPr shading is not paragraph shading', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:rPr><w:shd w:val="clear" w:fill="C0C0C0"/></w:rPr></w:pPr>' +
          '<w:r><w:rPr><w:shd w:val="clear" w:fill="C0C0C0"/></w:rPr></w:r></w:p>',
      }),
    )
    const format = doc.blocks[0].format
    expect(format?.shadingFill).toBeUndefined()
    expect(format?.shadingDisplay).toBeUndefined()
    expect(format?.shadingClear).toBeUndefined()
  })
})

describe('w:themeFill resolution', () => {
  it('re-resolves the fill from the theme for display and keeps the cached hex for save', async () => {
    // LibreOffice export: background1 (white) with a zero tint - Word paints white, not the cached D9D9D9
    const shd =
      '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1"' +
      ' w:themeFillTint="0" w:themeFillShade="d9"/>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p><w:pPr>${shd}</w:pPr><w:r><w:t>R</w:t></w:r></w:p>` }),
    )
    const format = doc.blocks[0].format!
    expect(format.shadingFill).toBe('D9D9D9')
    expect(format.shadingDisplay).toBe('FFFFFF')
    expect(format.shadingClear).toBeUndefined()
    expect(mergePPrFormat(doc.blocks[0].rawPPr!, format)).toContain(shd)
  })

  it('a plain themed fill that matches the cache sets no display twin', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFill="background1"/></w:pPr>' +
          '<w:r><w:t>R</w:t></w:r></w:p>',
      }),
    )
    expect(doc.blocks[0].format?.shadingFill).toBe('FFFFFF')
    expect(doc.blocks[0].format?.shadingDisplay).toBeUndefined()
  })

  it('table cell shading keeps the cached fill', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr>' +
          '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="d9"/>' +
          '</w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      }),
    )
    expect(doc.blocks[0].table?.rows[0][0].fill).toBe('D9D9D9')
  })
})

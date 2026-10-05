import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

const withBlip = (inner: string): string =>
  IMAGE_PARAGRAPH_XML.replace(
    '<a:blip r:embed="rId10"/>',
    `<a:blip r:embed="rId10">${inner}</a:blip>`,
  )

async function effectsOf(bodyXml: string) {
  const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return doc.blocks[0].imageEffects
}

describe('picture recolor (a:lum / a:grayscl / a:biLevel)', () => {
  it('reads brightness and contrast as fractions', async () => {
    expect(await effectsOf(withBlip('<a:lum bright="40000" contrast="40000"/>'))).toEqual({
      bright: 0.4,
      contrast: 0.4,
    })
    expect(await effectsOf(withBlip('<a:lum contrast="-20000"/>'))).toEqual({ contrast: -0.2 })
  })

  it('reads grayscale and the bi-level threshold (default 50%)', async () => {
    expect(await effectsOf(withBlip('<a:grayscl/>'))).toEqual({ grayscale: true })
    expect(await effectsOf(withBlip('<a:biLevel thresh="25000"/>'))).toEqual({
      biLevelThresh: 0.25,
    })
    expect(await effectsOf(withBlip('<a:biLevel/>'))).toEqual({ biLevelThresh: 0.5 })
    expect(await effectsOf(withBlip('<a:biLevel thresh="0"/>'))).toEqual({ biLevelThresh: 0 })
  })

  it('does not read a later picture through a self-closing first blip', async () => {
    const second = withBlip('<a:lum bright="40000"/>').replace(/^<w:p>|<\/w:p>$/g, '')
    const xml = IMAGE_PARAGRAPH_XML.replace('</w:p>', `${second}</w:p>`)
    expect(await effectsOf(xml)).toBeUndefined()
  })

  it('leaves a plain blip without effects', async () => {
    expect(await effectsOf(IMAGE_PARAGRAPH_XML)).toBeUndefined()
    expect(await effectsOf(withBlip('<a:lum/>'))).toBeUndefined()
  })
})

describe('gradFill schemeClr stops without a usable theme', () => {
  const BOX = (fillXml: string): string =>
    '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="4897120" cy="520700"/><wp:wrapNone/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/office/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.openxmlformats.org/office/2010/wordprocessingShape">' +
    `<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml}<a:ln><a:noFill/></a:ln></wps:spPr>` +
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>JUIN 2026</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

  it('falls back to the built-in Office palette when theme1.xml carries no clrScheme', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BOX(
          '<a:gradFill><a:gsLst>' +
            '<a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>' +
            '<a:gs pos="100000"><a:schemeClr val="hlink"/></a:gs>' +
            '</a:gsLst></a:gradFill>',
        ),
        extraParts: [
          {
            path: 'word/theme/theme1.xml',
            xml:
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
              '<a:themeElements><a:fontScheme name="T"><a:majorFont/><a:minorFont/></a:fontScheme></a:themeElements></a:theme>',
            contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
          },
        ],
      }),
    )
    // accent1 4472C4 + hlink 0563C1 from DEFAULT_THEME_COLORS, equal-weight average
    expect(doc.blocks[0].textboxes?.[0].fill).toBe('256BC3')
  })
})

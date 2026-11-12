/**
 * Slide-level <a:themeOverride> part (keep-source-formatting pastes): the slide's
 * schemeClr refs resolve against the override's clrScheme instead of the master
 * theme's — PowerPoint paints such a slide's bg2 background in the override's colour.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { openPptx } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OVERRIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:clrScheme name="o">' +
  '<a:dk1><a:srgbClr val="2E404D"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="19A3FC"/></a:dk2><a:lt2><a:srgbClr val="070F26"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="005B96"/></a:accent1><a:accent2><a:srgbClr val="0071BC"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="00DFED"/></a:accent3><a:accent4><a:srgbClr val="00CB5D"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="FFC300"/></a:accent5><a:accent6><a:srgbClr val="FF7900"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="E32600"/></a:hlink><a:folHlink><a:srgbClr val="38F990"/></a:folHlink>' +
  '</a:clrScheme></a:themeOverride>'

const BG =
  '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg2"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'

async function deckWithBg2Slide(withOverride: boolean): Promise<Buffer> {
  const zip = await JSZip.loadAsync(fx('01_standard_business.pptx'))
  const slidePath = 'ppt/slides/slide1.xml'
  const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
  const slide = await zip.file(slidePath)!.async('string')
  zip.file(slidePath, slide.replace(/<p:cSld([^>]*)>/, `<p:cSld$1>${BG}`))
  if (withOverride) {
    const rels = await zip.file(relsPath)!.async('string')
    zip.file(
      relsPath,
      rels.replace(
        '</Relationships>',
        '<Relationship Id="rIdThemeOvr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>',
      ),
    )
    zip.file('ppt/theme/themeOverride1.xml', OVERRIDE)
  }
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('slide-level themeOverride', () => {
  it('resolves the slide background against the override clrScheme', async () => {
    const { deck } = await openPptx(await deckWithBg2Slide(true))
    expect(deck.slides[0]!.background).toEqual({ type: 'solid', color: '#070F26' })
  })

  it('keeps the master theme colours when the slide has no override part', async () => {
    const { deck } = await openPptx(await deckWithBg2Slide(false))
    const base = (deck.slides[0]!.background as any).color
    expect(base).toMatch(/^#[0-9A-F]{6}$/)
    expect(base).not.toBe('#070F26')
  })
})

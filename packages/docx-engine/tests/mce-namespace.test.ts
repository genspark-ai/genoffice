import { describe, expect, it } from 'vitest'
import {
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
} from '../src/generate'

/**
 * ECMA-376 Part 3 (MCE): every prefix named in mc:Choice Requires must be
 * resolvable at the AlternateContent scope. Word treats an unresolvable
 * prefix as file corruption ("unreadable content"), so the
 * builders must declare the wps namespace on mc:AlternateContent itself.
 */

const BUILDERS: Array<[string, string]> = [
  ['wordart', buildWordArtParagraphXml({ text: 'Hi' })],
  ['textbox', buildTextboxParagraphXml()],
  ['shape', buildShapeParagraphXml({ prst: 'ellipse' })],
]

describe('MCE Requires prefixes resolve at AlternateContent scope', () => {
  for (const [name, xml] of BUILDERS) {
    it(`${name} declares every Requires prefix on mc:AlternateContent`, () => {
      const acTag = /<mc:AlternateContent[^>]*>/.exec(xml)?.[0]
      expect(acTag).toBeDefined()
      for (const m of xml.matchAll(/Requires="([^"]+)"/g)) {
        for (const prefix of m[1].split(/\s+/)) {
          expect(acTag, `prefix ${prefix} in ${name}`).toContain(`xmlns:${prefix}=`)
        }
      }
    })
  }

  for (const [name, xml] of BUILDERS) {
    it(`${name} emits the required wps:cNvSpPr child`, () => {
      const wsp = /<wps:wsp[^>]*>([\s\S]*?)<wps:spPr>/.exec(xml)?.[1] ?? ''
      expect(wsp, name).toContain('<wps:cNvSpPr')
    })
  }

  it('VML fallback sizes are points, not pixels', () => {
    const xml = buildWordArtParagraphXml({ widthEmu: 2700000, heightEmu: 720000 })
    const style = /<v:rect[^>]*style="([^"]*)"/.exec(xml)?.[1] ?? ''
    expect(style).toContain(`width:${2700000 / 12700}pt`)
    expect(style).toContain(`height:${720000 / 12700}pt`)
  })
})

/**
 * Tests for normAutofit fontScale write-back (patchBodyPrAutofit):
 * attribute replace/insert/clear, and stacking with patchTextElementXml (bodyPr kept as-is).
 */
import { describe, it, expect } from 'vitest'
import { patchBodyPrAutofit, patchTextElementXml } from '../src/index'
import type { TextElement } from '../src/types'

const SP = (bodyPr: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>` +
  `<p:txBody>${bodyPr}<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp>`

describe('patchBodyPrAutofit', () => {
  it('writes fontScale/lnSpcReduction onto bare normAutofit', () => {
    const out = patchBodyPrAutofit(SP('<a:bodyPr><a:normAutofit/></a:bodyPr>'), 0.625, 0.2)
    expect(out).toContain('<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>')
  })

  it('replaces existing attributes', () => {
    const out = patchBodyPrAutofit(
      SP('<a:bodyPr><a:normAutofit fontScale="90000" lnSpcReduction="10000"/></a:bodyPr>'),
      0.5,
    )
    expect(out).toContain('<a:normAutofit fontScale="50000"/>')
    expect(out).not.toContain('90000')
  })

  it('clears attributes when scale back to 1', () => {
    const out = patchBodyPrAutofit(SP('<a:bodyPr><a:normAutofit fontScale="62500"/></a:bodyPr>'), 1)
    expect(out).toContain('<a:normAutofit/>')
    expect(out).not.toContain('fontScale')
  })

  it('handles paired-tag normAutofit and leaves non-autofit xml untouched', () => {
    const paired = patchBodyPrAutofit(
      SP('<a:bodyPr><a:normAutofit></a:normAutofit></a:bodyPr>'),
      0.8,
    )
    expect(paired).toContain('<a:normAutofit fontScale="80000"/>')
    const noFit = SP('<a:bodyPr><a:spAutoFit/></a:bodyPr>')
    expect(patchBodyPrAutofit(noFit, 0.8)).toBe(noFit)
  })

  it('survives patchTextElementXml (bodyPr preserved through text regen)', () => {
    const xml = patchBodyPrAutofit(SP('<a:bodyPr><a:normAutofit/></a:bodyPr>'), 0.75)
    const el = {
      id: 'e',
      type: 'shape',
      anchor: { spIndex: 0, originalXml: xml, range: [0, xml.length] },
      transform: { offset: { x: 0, y: 0, cx: 100, cy: 100 }, rot: 0, flipH: false, flipV: false },
      text: {
        autofit: 'shrink',
        fontScale: 0.75,
        paragraphs: [{ runs: [{ text: 'edited', fontSize: 18 }] }],
      },
    } as unknown as TextElement
    const out = patchTextElementXml(el, xml)
    expect(out).toContain('fontScale="75000"')
    expect(out).toContain('edited')
  })
})

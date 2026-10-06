import { describe, expect, it } from 'vitest'
import { parseSlide, setElementRedaction } from '../src/index'
import { hasRedactExtIn, REDACT_EXT_URI } from '../src/redaction-xml'
import type { PictureElement, Slide, TextElement } from '../src/types'

/**
 * The two halves of "withhold this", and they are deliberately different:
 *
 * - **text** rides the existing edit pipeline as one more run attribute, so
 *   'setText' — which is journaled, undoable and already rebuilds the runs —
 *   carries it for free. A 'redact' of 'null' stops withholding, which is why
 *   the field is three-state like 'link' rather than a plain string.
 * - **a picture** has no runs, so it needs its own byte surgery on the element's
 *   'spPr', the same way 'setPictureOpacity' does for 'a:alphaModFix'.
 *
 * In both cases the words stay and the file still opens.
 */

const LABEL = 'client phone'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const SP = (rPr: string, text: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="900000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>${rPr}<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

const PIC_SP_PR =
  '<p:spPr><a:xfrm><a:off x="100000" y="1000000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>'

const PIC = (spPr = PIC_SP_PR) =>
  '<p:pic><p:nvPicPr><p:cNvPr id="3" name="p"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  spPr +
  '</p:pic>'

const load = (body: string): Slide =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(body), ctx: {} })

/** a text element's id, used to prove a non-picture target is refused */
const shapeOf = (s: Slide) => s.elements[0] as TextElement
const picOf = (s: Slide) => s.elements.find((e) => e.type === 'picture') as PictureElement

describe('withholding a picture', () => {
  it('puts the mark in the element spPr and on the model', () => {
    const s = load(PIC())
    expect(setElementRedaction(s, picOf(s).id, LABEL)).toBe(true)
    const pic = picOf(s)
    expect(pic.redact).toBe(LABEL)
    expect(pic.anchor.originalXml).toContain('go:redact')
    expect(pic.anchor.originalXml).toContain(REDACT_EXT_URI)
    // the picture's own bytes are otherwise untouched
    expect(pic.anchor.originalXml).toContain('r:embed="rId2"')
  })

  it('reopening the file reads the mark back', () => {
    const s = load(PIC())
    setElementRedaction(s, picOf(s).id, LABEL)
    const reopened = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(PIC(picOf(s).anchor.originalXml.replace(/^<p:pic>|<\/p:pic>$/g, ''))),
      ctx: {},
    })
    expect((reopened.elements[0] as PictureElement).redact).toBe(LABEL)
  })

  it('clears the mark and leaves no empty extension list', () => {
    const s = load(PIC())
    setElementRedaction(s, picOf(s).id, LABEL)
    expect(setElementRedaction(s, picOf(s).id, null)).toBe(true)
    const pic = picOf(s)
    expect(pic.redact).toBeUndefined()
    expect(hasRedactExtIn(pic.anchor.originalXml)).toBe(false)
    expect(pic.anchor.originalXml).not.toContain('extLst')
  })

  it('re-applying replaces rather than stacking', () => {
    const s = load(PIC())
    setElementRedaction(s, picOf(s).id, 'first')
    setElementRedaction(s, picOf(s).id, 'second')
    const xml = picOf(s).anchor.originalXml
    expect(xml.match(/go:redact/g)).toHaveLength(1)
    expect(xml).toContain('label="second"')
  })

  it('refuses an id that is not a picture', () => {
    const s = load(SP('<a:rPr lang="en-US"/>', 'text'))
    expect(setElementRedaction(s, shapeOf(s).id, LABEL)).toBe(false)
    expect(setElementRedaction(s, 'no-such-id', LABEL)).toBe(false)
  })

  it('marking does not disturb the picture geometry', () => {
    const s = load(PIC())
    const before = picOf(s).anchor.originalXml
    setElementRedaction(s, picOf(s).id, LABEL)
    const after = picOf(s).anchor.originalXml
    for (const frag of [
      '<a:off x="100000" y="1000000"/>',
      '<a:ext cx="2000000" cy="1500000"/>',
      '<a:prstGeom prst="rect"/>',
    ]) {
      expect(before).toContain(frag)
      expect(after).toContain(frag)
    }
  })
})

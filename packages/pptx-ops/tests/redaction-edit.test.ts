import { describe, expect, it } from 'vitest'
import { parseSlide, patchTextElementXml } from '@genoffice/pptx-engine'
import { applyEditParagraphs } from '../src/edit-text'
import type { Slide, TextElement } from '@genoffice/pptx-engine'

/**
 * Withholding text rides the existing edit pipeline as one more run attribute:
 * `setText` — journaled, undoable, and already rebuilding the runs — carries it
 * for free. Three states, like `link`: absent keeps the original mark, `null`
 * clears it, a string withholds under that name. The words stay in the file.
 */

const LABEL = 'client phone'
const SECRET = '13800138000'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const SP = (rPr: string, text: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="900000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>${rPr}<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

const load = (body: string): Slide =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(body), ctx: {} })

const shapeOf = (s: Slide) => s.elements[0] as TextElement

describe('withholding text through the edit pipeline', () => {
  it('a run marked in an edit lands on the model', () => {
    const before = shapeOf(load(SP('<a:rPr lang="en-US"/>', SECRET))).text!.paragraphs
    const after = applyEditParagraphs(before, [
      { runs: [{ text: SECRET, srcRun: 0, redact: LABEL, underline: true }] },
    ])
    expect(after[0]!.runs[0]!.redact).toBe(LABEL)
    expect(after[0]!.runs[0]!.text).toBe(SECRET)
  })

  it('the words are untouched and the file still holds them', () => {
    const s = load(SP('<a:rPr lang="en-US"/>', SECRET))
    const el = shapeOf(s)
    el.text = {
      paragraphs: applyEditParagraphs(el.text!.paragraphs, [
        { runs: [{ text: SECRET, srcRun: 0, redact: LABEL, underline: true }] },
      ]),
    }
    const xml = patchTextElementXml(el, el.anchor.originalXml)
    expect(xml).toContain(SECRET)
    expect(xml).not.toContain('{{')
    expect(xml).toContain('go:redact')
  })

  it('an unmarked run in the same edit keeps the original state', () => {
    const before = shapeOf(load(SP('<a:rPr lang="en-US"/>', 'plain'))).text!.paragraphs
    const after = applyEditParagraphs(before, [{ runs: [{ text: 'plain', srcRun: 0 }] }])
    expect(after[0]!.runs[0]!.redact).toBeUndefined()
  })

  it('an edit that says nothing about the mark leaves it alone', () => {
    const before = shapeOf(load(SP('<a:rPr lang="en-US"/>', SECRET))).text!.paragraphs
    const marked = applyEditParagraphs(before, [
      { runs: [{ text: SECRET, srcRun: 0, redact: LABEL }] },
    ])
    // a later edit that omits 'redact' entirely must not clear it
    const again = applyEditParagraphs(marked, [{ runs: [{ text: 'changed', srcRun: 0 }] }])
    expect(again[0]!.runs[0]!.redact).toBe(LABEL)
  })

  it('null stops withholding, and the file loses the mark', () => {
    const before = shapeOf(load(SP('<a:rPr lang="en-US"/>', SECRET))).text!.paragraphs
    const marked = applyEditParagraphs(before, [
      { runs: [{ text: SECRET, srcRun: 0, redact: LABEL }] },
    ])
    const cleared = applyEditParagraphs(marked, [
      { runs: [{ text: SECRET, srcRun: 0, redact: null }] },
    ])
    expect(cleared[0]!.runs[0]!.redact).toBeUndefined()
    const el = shapeOf(load(SP('<a:rPr lang="en-US"/>', SECRET)))
    el.text = { paragraphs: cleared }
    expect(patchTextElementXml(el, el.anchor.originalXml)).not.toContain('go:redact')
  })

  it('an empty string clears too, matching how the editor DOM sends "off"', () => {
    const before = shapeOf(load(SP('<a:rPr lang="en-US"/>', SECRET))).text!.paragraphs
    const marked = applyEditParagraphs(before, [
      { runs: [{ text: SECRET, srcRun: 0, redact: LABEL }] },
    ])
    const cleared = applyEditParagraphs(marked, [
      { runs: [{ text: SECRET, srcRun: 0, redact: '' }] },
    ])
    expect(cleared[0]!.runs[0]!.redact).toBeUndefined()
  })
})

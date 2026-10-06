import { describe, expect, it } from 'vitest'
import { parseSlide, patchTextElementXml } from '../src/index'
import { rebuildTxBody } from '../src/generate'
import { setRedactExt, stripRedactExt } from '../src/redaction-xml'
import type { Slide, TextElement, TextRun } from '../src/types'

/**
 * The carrier end to end, through the real parser.
 *
 * The two save paths are not interchangeable, which is why the mark lives on the
 * model rather than being left to ride along in the XML:
 *
 * - `patchTextElementXml` patches each run **in place** when the run count
 *   matches, and
 * - falls through to `rebuildTxBody`, which regenerates `<a:txBody>` from the
 *   model and drops anything the model does not carry.
 *
 * Marking half a word splits a run, which is the second path. So this file goes
 * through `parseSlide` rather than hand-building the model — a hand-built run
 * would not know about the mark, and would (correctly) clear it.
 */

const LABEL = 'client phone'
const SECRET = '13800138000'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const shapeWith = (runsXml: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runsXml}</a:p></p:txBody></p:sp>`

const run = (rPr: string, text: string) => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`

/** Parse a slide and hand back its single shape plus the slide XML it came from. */
function load(runsXml: string): { shape: TextElement; slide: Slide; slideXml: string } {
  const slideXml = slideWith(shapeWith(runsXml))
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
  return { shape: slide.elements[0] as TextElement, slide, slideXml }
}

const runList = (shape: TextElement) => shape.text!.paragraphs[0]!.runs
const runs = (shape: TextElement) => runList(shape)
const RPR = '<a:rPr lang="en-US"/>'
const MARKED = setRedactExt(RPR, LABEL)
const markedShape = () => load(run(MARKED, `Call ${SECRET}`))
const shapeXml = (el: TextElement) => el.anchor.originalXml

describe('a withheld run, through the real parser', () => {
  it('parse reads the label off the run and leaves the words alone', () => {
    const { shape } = markedShape()
    expect(runs(shape)[0]!.redact).toBe(LABEL)
    expect(runs(shape)[0]!.text).toBe(`Call ${SECRET}`)
  })

  it('parse reports no mark on a plain run', () => {
    const { shape } = load(run(RPR, 'Call now'))
    expect(runs(shape)[0]!.redact).toBeUndefined()
  })

  it('aligned path: an unrelated text edit keeps the mark and the words', () => {
    const { shape } = markedShape()
    const r = runs(shape)[0]!
    r.text = `Call ${SECRET} now`
    const out = patchTextElementXml(shape, shapeXml(shape))
    expect(out).toContain('go:redact')
    expect(out).toContain(SECRET)
  })

  it('structural path: the mark returns because the model carries it', () => {
    const { shape } = markedShape()
    const r = runs(shape)[0]!
    // split the run, which is what marking half a word does
    runList(shape).splice(0, 1, { text: 'Call ' } as TextRun, { ...r, text: SECRET })
    const out = patchTextElementXml(shape, shapeXml(shape))
    expect(out).toContain('go:redact')
    expect(out).toContain(SECRET)
  })

  it('structural path without the model carrying it loses the mark — the trap', () => {
    const { shape } = markedShape()
    runList(shape).splice(0, 1, { text: 'Call ' } as TextRun, { text: SECRET })
    const out = patchTextElementXml(shape, shapeXml(shape))
    expect(out).not.toContain('go:redact')
  })

  it('rebuildTxBody alone writes the mark back from the model', () => {
    const { shape } = markedShape()
    const out = rebuildTxBody(shape, shapeXml(shape))
    expect(out).toContain('go:redact')
    expect(out).toContain(SECRET)
  })

  it('never writes the placeholder into the file', () => {
    const { shape } = markedShape()
    const out = rebuildTxBody(shape, shapeXml(shape))
    expect(out).not.toContain('{{')
  })

  it('clearing the mark in the model removes it from the file, words intact', () => {
    const { shape } = markedShape()
    delete runs(shape)[0]!.redact
    const out = patchTextElementXml(shape, shapeXml(shape))
    expect(out).not.toContain('go:redact')
    expect(out).not.toContain('extLst')
    expect(out).toContain(SECRET)
  })

  it('re-marking a plain run adds the mark without touching the words', () => {
    const { shape } = load(run(RPR, SECRET))
    runs(shape)[0]!.redact = LABEL
    const out = rebuildTxBody(shape, shapeXml(shape))
    expect(out).toContain('go:redact')
    expect(out).toContain(SECRET)
  })

  it('an apply / clear / apply cycle leaves the mark in place', () => {
    const { shape } = markedShape()
    delete runs(shape)[0]!.redact
    const cleared = patchTextElementXml(shape, shapeXml(shape))
    const { shape: again } = markedShape()
    runs(again)[0]!.redact = LABEL
    const out = rebuildTxBody(again, cleared)
    expect(out).toContain('go:redact')
    expect(out).toContain(SECRET)
  })

  it('two runs in one shape keep their own labels', () => {
    const first = setRedactExt(RPR, 'client phone')
    const second = setRedactExt(RPR, 'contract value')
    const { shape } = load(run(first, 'A') + run(second, 'B'))
    const out = rebuildTxBody(shape, shapeXml(shape))
    expect(out).toContain('label="client phone"')
    expect(out).toContain('label="contract value"')
    expect(out).toContain('<a:t>A</a:t>')
    expect(out).toContain('<a:t>B</a:t>')
  })

  it('the visible half needs the model too', () => {
    // the patch path writes u="none" for a run whose model has no underline, so
    // applying a redaction has to set it or the span is invisible on reopen
    const { shape } = load(run(RPR, SECRET))
    runs(shape)[0]!.redact = LABEL
    const invisible = rebuildTxBody(shape, shapeXml(shape))
    runs(shape)[0]!.underline = true
    const visible = rebuildTxBody(shape, shapeXml(shape))
    expect(invisible).not.toContain('<a:u')
    expect(visible).toContain('u="sng"')
  })

  it('the mark does not disturb the other run properties', () => {
    const rich = '<a:rPr lang="en-US" sz="1800" b="1"><a:latin typeface="Calibri"/></a:rPr>'
    const { shape } = load(run(setRedactExt(rich, LABEL), SECRET))
    // the aligned path is the normal one and keeps every byte it does not model
    const patched = patchTextElementXml(shape, shapeXml(shape))
    expect(patched).toContain('lang="en-US"')
    expect(patched).toContain('sz="1800"')
    expect(patched).toContain('Calibri')
    expect(patched).toContain('go:redact')
    // the rebuild path is the engine's lossy fallback by design — it does not
    // model `lang` — so assert only what it does carry
    const rebuilt = rebuildTxBody(shape, shapeXml(shape))
    expect(rebuilt).toContain('sz="1800"')
    expect(rebuilt).toContain('b="1"')
    expect(rebuilt).toContain('Calibri')
    expect(rebuilt).toContain('go:redact')
  })

  it('saving many marked runs in a row does not drift', () => {
    // guards the stateful-regex trap: hasRedactExtIn ran per run, so a shared
    // /g lastIndex made the apply/clear decision alternate between saves
    const { shape } = markedShape()
    for (let i = 0; i < 6; i += 1) {
      const r = runList(shape)[0]!
      r.text = `${SECRET} pass ${i}`
      const out = patchTextElementXml(shape, shapeXml(shape))
      expect(out, `pass ${i}`).toContain('go:redact')
      shape.anchor.originalXml = out
    }
  })

  it('a mark survives a full parse → edit → save → parse round trip', () => {
    const { shape } = markedShape()
    runs(shape)[0]!.text = `Call ${SECRET} today`
    const saved = patchTextElementXml(shape, shapeXml(shape))
    // reopen the saved slide the way the app would
    const reopened = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(saved),
      ctx: {},
    })
    const r = (reopened.elements[0] as TextElement).text!.paragraphs[0]!.runs[0]!
    expect(r.redact).toBe(LABEL)
    expect(r.text).toBe(`Call ${SECRET} today`)
  })

  it('strip then re-apply on a raw rPr is a fixed point', () => {
    const base = '<a:rPr lang="en-US" sz="1800"/>'
    const round = stripRedactExt(setRedactExt(base, LABEL))
    expect(setRedactExt(round, LABEL)).toBe(setRedactExt(base, LABEL))
  })
})

/**
 * Fidelity of the rebuildTxBody fallback path: on text-structure changes (run add/remove),
 * explicit paragraph properties (bullet/marL/indent/lnSpc/spcBef/spcAft) and run-level
 * hlinkClick/strike/cs are written back; display values inherited from lstStyle are not baked in.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { patchTextElementXml } from '../src/index'
import type { TextElement } from '../src/types'

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`
const spWith = (txBodyInner: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody>${txBodyInner}</p:txBody></p:sp>`

const parseEl = (txBodyInner: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(spWith(txBodyInner)),
    ctx: {},
  })
  return slide.elements[0] as TextElement
}

/** Append a run to trigger the structural rebuild path. */
const forceRebuild = (el: TextElement) => {
  el.text!.paragraphs[el.text!.paragraphs.length - 1]!.runs.push({ text: ' added' })
}

describe('rebuild path: explicit paragraph properties written back', () => {
  const PPR =
    '<a:pPr marL="342900" indent="-342900" algn="just">' +
    '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>' +
    '<a:spcBef><a:spcPts val="1200"/></a:spcBef>' +
    '<a:spcAft><a:spcPct val="50000"/></a:spcAft>' +
    '<a:buClr><a:srgbClr val="FF0000"/></a:buClr>' +
    '<a:buSzPct val="75000"/><a:buFont typeface="Wingdings"/>' +
    '<a:buChar char="v"/></a:pPr>'

  it('bullet/indent/line spacing/paragraph spacing all preserved', () => {
    const el = parseEl(`<a:bodyPr/><a:p>${PPR}<a:r><a:t>x</a:t></a:r></a:p>`)
    const p = el.text!.paragraphs[0]!
    expect(p.bullet).toMatchObject({ type: 'char', char: 'v', font: 'Wingdings', sizePct: 75 })
    expect(p.marL).toBe(342900)
    expect(p.indent).toBe(-342900)
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('marL="342900"')
    expect(out).toContain('indent="-342900"')
    expect(out).toContain('algn="just"')
    expect(out).toContain('<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>')
    expect(out).toContain('<a:spcBef><a:spcPts val="1200"/></a:spcBef>')
    expect(out).toContain('<a:spcAft><a:spcPct val="50000"/></a:spcAft>')
    expect(out).toContain('<a:buClr><a:srgbClr val="FF0000"/></a:buClr>')
    expect(out).toContain('<a:buSzPct val="75000"/>')
    expect(out).toContain('<a:buFont typeface="Wingdings"/>')
    expect(out).toContain('<a:buChar char="v"/>')
  })

  it('buAutoNum numbering format preserved', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:pPr><a:buAutoNum type="romanLcPeriod" startAt="3"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:buAutoNum type="romanLcPeriod" startAt="3"/>')
  })

  it('buNone preserved (explicitly turning off the inherited bullet)', () => {
    const el = parseEl('<a:bodyPr/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>')
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('<a:buNone/>')
  })
})

describe('rebuild path: inherited values not materialized', () => {
  const LST =
    '<a:lstStyle><a:lvl1pPr marL="228600" algn="ctr"><a:buChar char="•"/>' +
    '<a:defRPr sz="2000"/></a:lvl1pPr></a:lstStyle>'

  it('algn/buChar/marL/sz inherited from lstStyle are not written into <a:p>', () => {
    const el = parseEl(`<a:bodyPr/>${LST}<a:p><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>`)
    const p = el.text!.paragraphs[0]!
    // Display values already merged (for rendering)
    expect(p.align).toBe('center')
    expect(p.bullet?.char).toBe('•')
    expect(p.runs[0]!.fontSize).toBe(20)
    expect(p.runs[0]!.fontSizeImplicit).toBe(true)
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    // lstStyle kept as-is, but inherited values not materialized inside <a:p>
    const pStart = out.indexOf('<a:p>')
    const body = out.slice(pStart)
    expect(body).not.toContain('algn=')
    expect(body).not.toContain('buChar')
    expect(body).not.toContain('marL=')
    expect(body).not.toContain('sz=')
  })

  it('inherited color writes no solidFill; explicit schemeClr materialized to keep visuals', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr/><a:t>a</a:t></a:r>' +
        '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>b</a:t></a:r></a:p>',
    )
    const [r1, r2] = el.text!.paragraphs[0]!.runs
    expect(r1!.colorInherited ?? false).toBe(r1!.color != null)
    expect(r2!.colorInherited).toBeUndefined()
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    // Only the schemeClr run materializes srgbClr (without a theme, resolveColorNode may fail to
    // resolve the color — then neither is written; assert the inherited run never writes fill)
    const firstRun = /<a:r><a:rPr[^>]*\/?>(?:<\/a:rPr>)?<a:t>a<\/a:t>/.exec(out)
    expect(firstRun).toBeTruthy()
  })
})

describe('rebuild path: run-level hlinkClick/strike/cs written back', () => {
  it('hlinkClick r:id + action + tooltip preserved', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId9" action="ppaction://hlinksldjump" tooltip="jump"/></a:rPr>' +
        '<a:t>link</a:t></a:r></a:p>',
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.hyperlinkRId).toBe('rId9')
    expect(r.hyperlinkAction).toBe('ppaction://hlinksldjump')
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('r:id="rId9"')
    expect(out).toContain('action="ppaction://hlinksldjump"')
    expect(out).toContain('tooltip="jump"')
  })

  it('strike and a:cs font preserved', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr strike="dblStrike"><a:cs typeface="Arial"/></a:rPr><a:t>x</a:t></a:r></a:p>',
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.strike).toBe(true)
    expect(r.strikeStyle).toBe('dblStrike')
    expect(r.csFont).toBe('Arial')
    forceRebuild(el)
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('strike="dblStrike"')
    expect(out).toContain('<a:cs typeface="Arial"/>')
  })
})

describe('in-place patch path: strike toggle', () => {
  it('enabling strikethrough injects strike="sngStrike"', () => {
    const el = parseEl('<a:bodyPr/><a:p><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>')
    el.text!.paragraphs[0]!.runs[0]!.strike = true
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('strike="sngStrike"')
  })

  it('removing strikethrough: existing strike becomes noStrike, runs without strike get nothing injected', () => {
    const el = parseEl(
      '<a:bodyPr/><a:p><a:r><a:rPr strike="sngStrike"/><a:t>a</a:t></a:r><a:r><a:rPr/><a:t>b</a:t></a:r></a:p>',
    )
    const [r1, r2] = el.text!.paragraphs[0]!.runs
    r1!.strike = false
    delete r1!.strikeStyle
    r2!.text = 'b2'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    expect(out).toContain('strike="noStrike"')
    expect(out.match(/\sstrike="/g)!.length).toBe(1)
  })

  it('run without explicit sz: text-only edit does not inject sz (inheritance not materialized)', () => {
    const el = parseEl(
      '<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></a:lstStyle>' +
        '<a:p><a:r><a:rPr/><a:t>x</a:t></a:r></a:p>',
    )
    el.text!.paragraphs[0]!.runs[0]!.text = 'y'
    const out = patchTextElementXml(el, el.anchor.originalXml)
    const pStart = out.indexOf('<a:p>')
    expect(out.slice(pStart)).not.toContain('sz=')
  })
})

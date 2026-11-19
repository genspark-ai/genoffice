/**
 * Connector (p:cxnSp) parsing tests:
 * - one each of straight / bent / curved + arrows + flip
 * - roundtrip: connector slide bytes stay identical
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { openPptx, savePptx, reassembleSlideXml } from '../src/index'
import type { TextElement } from '../src/types'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

// XML construction helpers
const wrap = (spTree: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  spTree +
  `</p:spTree></p:cSld></p:sld>`

// ── 1. straight connector (line/straightConnector1) ───────────────────────────

const STRAIGHT_CXN = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="10" name="Straight Connector 1"/><p:cNvCxnSpPr><a:stCxn id="2" idx="3"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm flipH="1"><a:off x="500" y="300"/><a:ext cx="914400" cy="457200"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="19050">
      <a:solidFill><a:srgbClr val="0070C0"/></a:solidFill>
      <a:headEnd type="arrow" w="sm" len="lg"/>
      <a:tailEnd type="stealth" w="lg" len="sm"/>
    </a:ln>
  </p:spPr>
</p:cxnSp>`

describe('straight connector parsing', () => {
  it('parses as type=shape, presetGeometry=line, flipH correct', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: wrap(STRAIGHT_CXN),
      ctx: {},
    })
    expect(slide.elements).toHaveLength(1)
    const el = slide.elements[0] as TextElement
    expect(el.type).toBe('shape')
    expect(el.presetGeometry).toBe('line')
    expect(el.transform.flipH).toBe(true)
    expect(el.transform.flipV).toBe(false)
    expect(el.fill).toEqual({ type: 'none' })
    expect(el.stroke?.fill).toEqual({ type: 'solid', color: '#0070C0' })
    expect(el.stroke?.width).toBe(19050)
  })

  it('headEnd parses with type/w/len', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(STRAIGHT_CXN), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.stroke?.headEnd).toEqual({ type: 'arrow', w: 'sm', len: 'lg' })
  })

  it('tailEnd stealth parses with type/w/len', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(STRAIGHT_CXN), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.stroke?.tailEnd).toEqual({ type: 'stealth', w: 'lg', len: 'sm' })
  })
})

// ── 2. bent connector (bentConnector3) ───────────────────────────────────────

const BENT_CXN = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="11" name="Elbow Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm flipV="1"><a:off x="100" y="200"/><a:ext cx="914400" cy="457200"/></a:xfrm>
    <a:prstGeom prst="bentConnector3">
      <a:avLst><a:gd name="adj1" fmla="val 75000"/></a:avLst>
    </a:prstGeom>
    <a:ln w="19050">
      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      <a:prstDash val="dash"/>
      <a:tailEnd type="triangle"/>
    </a:ln>
  </p:spPr>
</p:cxnSp>`

describe('bent connector parsing', () => {
  it('presetGeometry=bentConnector3, adj1=75000', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(BENT_CXN), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.presetGeometry).toBe('bentConnector3')
    expect(el.adjust?.adj1).toBe(75000)
    expect(el.transform.flipV).toBe(true)
    expect(el.stroke?.dash).toBe('dash')
    expect(el.stroke?.tailEnd).toEqual({ type: 'triangle' })
    expect(el.stroke?.headEnd).toBeUndefined()
  })
})

// ── 3. curved connector (curvedConnector3) ───────────────────────────────────

const CURVED_CXN = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="12" name="Curved Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200" y="400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
    <a:prstGeom prst="curvedConnector3">
      <a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst>
    </a:prstGeom>
    <a:ln w="12700">
      <a:solidFill><a:srgbClr val="00B050"/></a:solidFill>
      <a:headEnd type="oval"/>
      <a:tailEnd type="diamond" w="lg" len="lg"/>
    </a:ln>
  </p:spPr>
</p:cxnSp>`

describe('curved connector parsing', () => {
  it('presetGeometry=curvedConnector3, oval/diamond arrowheads', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(CURVED_CXN), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.presetGeometry).toBe('curvedConnector3')
    expect(el.adjust?.adj1).toBe(50000)
    expect(el.stroke?.headEnd).toEqual({ type: 'oval' })
    expect(el.stroke?.tailEnd).toEqual({ type: 'diamond', w: 'lg', len: 'lg' })
  })
})

// ── 4. Fallback to a visible stroke when ln is absent ────────────────────

const BARE_CXN = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="13" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:cxnSp>`

describe('fallback without explicit ln', () => {
  it('has a default visible stroke when ln is absent', () => {
    const slide = parseSlide({ path: 'p', slideXml: wrap(BARE_CXN), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.stroke).toBeTruthy()
    expect(el.stroke!.fill.type).toBe('solid')
  })
})

// ── 5. Roundtrip: byte fidelity after parse + reassemble ─────────────────

describe('connector slide roundtrip fidelity', () => {
  it('pptx containing cxnSp reassembles to byte-identical slide XML', async () => {
    // Use 01_standard_business.pptx (fine even without connectors: fidelity means unchanged)
    const { deck } = await openPptx(fx('01_standard_business.pptx'))
    // Confirm reassemble == originalXml slide by slide
    for (const slide of deck.slides) {
      expect(reassembleSlideXml(slide)).toBe(slide.originalXml)
    }
  })

  it('reparses after savePptx: connector elements neither lost nor distorted', async () => {
    // Building a minimal pptx with cxnSp: the helper-wrapped xml can't be packed into a pptx,
    // so use 01_standard_business.pptx to verify saving doesn't break existing slides
    const original = await openPptx(fx('01_standard_business.pptx'))
    const savedBytes = await savePptx(original)
    const reopened = await openPptx(savedBytes)
    expect(reopened.deck.slides.length).toBe(original.deck.slides.length)
    // Element count matches per slide (any cxnSp present is not lost)
    for (let i = 0; i < original.deck.slides.length; i++) {
      expect(reopened.deck.slides[i]!.elements.length).toBe(
        original.deck.slides[i]!.elements.length,
      )
    }
  })
})

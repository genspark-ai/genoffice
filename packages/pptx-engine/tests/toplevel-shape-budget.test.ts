/**
 * Top-level shape budget: the group path already bounds p:grpSp descendants
 * (MAX_GROUP_DEPTH / MAX_GROUP_DESCENDANTS), but a slide's own <p:spTree>
 * children were unbounded — a slide carrying a million top-level shapes parsed
 * into a million model elements. The top level now spends the same descendant
 * budget, and the surplus stays one byte-preserving passthrough so a save still
 * replays it verbatim.
 */
import { describe, it, expect } from 'vitest'
import { reassembleSlideXml } from '../src/index'
import { parseSlide } from '../src/parse'

// The budget the group path spends on p:grpSp descendants (parse.ts MAX_GROUP_DESCENDANTS).
const GROUP_DESCENDANT_BUDGET = 10_000

const SP = (id: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>`

const slideWithShapes = (n: number) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  Array.from({ length: n }, (_, i) => SP(i + 2)).join('') +
  '</p:spTree></p:cSld></p:sld>'

describe('top-level shape budget', () => {
  it('keeps a slide with more top-level shapes than the group budget bounded', () => {
    const over = GROUP_DESCENDANT_BUDGET + 50
    const slideXml = slideWithShapes(over)
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })

    // Only the budget is turned into model elements; the rest is a single passthrough.
    const passthroughs = slide.elements.filter((e) => e.type === 'passthrough')
    expect(slide.elements.length - passthroughs.length).toBeLessThanOrEqual(GROUP_DESCENDANT_BUDGET)
    expect(passthroughs).toHaveLength(1)
    expect(slide.elements.length).toBeLessThan(over)
  })

  it('replays the shapes past the budget verbatim (no bytes lost)', () => {
    const slideXml = slideWithShapes(GROUP_DESCENDANT_BUDGET + 50)
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })

    expect(reassembleSlideXml(slide)).toBe(slideXml)
  })

  it('leaves a slide inside the budget untouched', () => {
    const slideXml = slideWithShapes(12)
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })

    expect(slide.elements).toHaveLength(12)
    expect(slide.elements.some((e) => e.type === 'passthrough')).toBe(false)
    expect(reassembleSlideXml(slide)).toBe(slideXml)
  })
})

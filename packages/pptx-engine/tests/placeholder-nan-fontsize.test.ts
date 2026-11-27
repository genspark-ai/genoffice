/**
 * Placeholder fontSize chain vs malformed sz regression:
 *
 * parseDefRPrStyle fed sz through a bare parseInt, so a malformed value
 * (sz="auto", garbage) became NaN in LevelTextStyle.fontSize. NaN is != null,
 * so mergeTextStyleChain adopted it and blocked every lower layer's real size
 * (layout → master ph → master txStyles); the run then carried fontSize NaN,
 * which the write path clamps to the 1pt floor (sz="100", invisible text).
 * A malformed sz now stays undefined: the chain keeps looking, and with no
 * valid size anywhere the run simply keeps fontSizeImplicit (no sz written).
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import {
  parseDefRPrStyle,
  parseLstStyleLevels,
  parseMasterTextStyles,
  parsePlaceholderMap,
} from '../src/placeholder'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

// layout body placeholder whose lstStyle carries a malformed sz
const LAYOUT =
  '<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="body ph"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="auto"/></a:lvl1pPr>' +
  '<a:lvl2pPr><a:defRPr sz="auto"/></a:lvl2pPr></a:lstStyle><a:p/></p:txBody></p:sp>' +
  '</p:spTree></p:cSld></p:sldLayout>'

// master txStyles body family with real sizes for lvl1/lvl2
const MASTER =
  '<?xml version="1.0"?><p:sldMaster xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree/></p:cSld>' +
  '<p:txStyles><p:bodyStyle>' +
  '<a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>' +
  '<a:lvl2pPr><a:defRPr sz="1600"/></a:lvl2pPr>' +
  '</p:bodyStyle></p:txStyles></p:sldMaster>'

const bodyPh = (paraXml: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="5" name="ph"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:p>${paraXml}</a:p></p:txBody></p:sp>`

const run = (rPr: string, text: string) => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`

describe('parseDefRPrStyle malformed sz', () => {
  it('sz="auto" alone yields no style instead of fontSize NaN', () => {
    expect(parseDefRPrStyle({ '@_sz': 'auto' })).toBeUndefined()
  })

  it('malformed sz stays undefined while sibling attrs survive', () => {
    const st = parseDefRPrStyle({ '@_sz': 'auto', '@_b': '1' })
    expect(st).toBeDefined()
    expect(st!.fontSize).toBeUndefined()
    expect(st!.bold).toBe(true)
  })

  it('valid sz still parses to pt', () => {
    expect(parseDefRPrStyle({ '@_sz': '2400' })!.fontSize).toBe(24)
  })

  it('lstStyle level drops the malformed sz but keeps sibling attrs', () => {
    const st = parseLstStyleLevels({
      'a:lvl1pPr': { 'a:defRPr': { '@_sz': 'auto', '@_b': '1' } },
    })
    expect(st!.levels[0]!.fontSize).toBeUndefined()
    expect(st!.levels[0]!.bold).toBe(true)
    // a level whose defRPr carried nothing but sz="auto" has no style at all
    expect(parseLstStyleLevels({ 'a:lvl1pPr': { 'a:defRPr': { '@_sz': 'auto' } } })).toBeUndefined()
  })
})

describe('placeholder chain falls through malformed sz', () => {
  const ctx = () => ({
    layoutPlaceholders: parsePlaceholderMap(LAYOUT),
    masterTextStyles: parseMasterTextStyles(MASTER),
  })

  it('run inherits the master txStyles size the NaN used to block', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(bodyPh(run('<a:rPr lang="en-US"/>', 'x'))),
      ctx: ctx(),
    })
    const r = (slide.elements[0] as any).text.paragraphs[0].runs[0]
    expect(r.fontSize).toBe(18)
    expect(Number.isFinite(r.fontSize)).toBe(true)
  })

  it('level-specific: lvl2 run takes the master lvl2 size', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(bodyPh('<a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>x</a:t></a:r>')),
      ctx: ctx(),
    })
    expect((slide.elements[0] as any).text.paragraphs[0].runs[0].fontSize).toBe(16)
  })

  it('no valid size anywhere → fontSize stays undefined (implicit), never NaN', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(bodyPh(run('<a:rPr lang="en-US"/>', 'x'))),
      ctx: { layoutPlaceholders: parsePlaceholderMap(LAYOUT) },
    })
    const r = (slide.elements[0] as any).text.paragraphs[0].runs[0]
    expect(r.fontSize).toBeUndefined()
    expect(r.fontSizeImplicit).toBe(true)
  })

  it('an explicit run sz still wins over the chain', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(bodyPh(run('<a:rPr lang="en-US" sz="2800"/>', 'x'))),
      ctx: ctx(),
    })
    expect((slide.elements[0] as any).text.paragraphs[0].runs[0].fontSize).toBe(28)
  })
})

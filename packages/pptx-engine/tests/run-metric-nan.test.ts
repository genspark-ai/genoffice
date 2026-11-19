/**
 * `sz` and `spc` are integer hundredths of a point. The reader divided them
 * unconditionally, so `sz="abc"` put NaN into the run: `?? DEFAULT_SIZE_PT`
 * does not catch NaN, and every glyph width and line height derived from it
 * was NaN too. A malformed attribute must read as absent and inherit.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const runWith = (rPr: string): any => {
  const shape =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
    '<a:prstGeom prst="rect"/></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr ${rPr}/><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>`
  const element: any = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(shape),
    ctx: {},
  }).elements[0]
  return element.text.paragraphs[0].runs[0]
}

describe('run metric attributes', () => {
  it('reads well-formed values as hundredths of a point', () => {
    const run = runWith('sz="1800" spc="-100"')
    expect(run.fontSize).toBe(18)
    expect(run.letterSpacing).toBe(-1)
  })

  it('inherits instead of producing NaN from a malformed size', () => {
    const run = runWith('sz="abc"')
    expect(run.fontSize).toBeUndefined()
    expect(run.styleSrc.fontSize).not.toBe('run')
    expect(run.fontSizeImplicit).toBe(true)
  })

  it('drops a malformed letter spacing', () => {
    const run = runWith('spc="wide"')
    expect(run.letterSpacing).toBeUndefined()
  })
})

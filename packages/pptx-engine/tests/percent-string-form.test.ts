/**
 * ST_Percentage has two spellings: the schema's own integer in 1/1000 %
 * ("50000") and the suffixed form Google Slides writes ("50%"). Reading the
 * suffixed form as a bare integer made every such value a thousandth of what
 * the file meant — a 50% gradient stop collapsed to the left edge, a 62.5%
 * font scale to 0.06%, a 150% line height to 0.15%.
 */
import { describe, it, expect } from 'vitest'
import { applyColorMods } from '../src/color'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const first = (shape: string): any =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(shape), ctx: {} }).elements[0]

const shapeWithFill = (fill: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="R"/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  `<a:prstGeom prst="rect"/>${fill}</p:spPr></p:sp>`

const shapeWithText = (bodyPrChildren: string, paragraph: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr>${bodyPrChildren}</a:bodyPr><a:lstStyle/><a:p>${paragraph}</a:p></p:txBody></p:sp>`

const stopsOf = (stops: string): Array<{ pos: number; color: string }> =>
  first(shapeWithFill(`<a:gradFill><a:gsLst>${stops}</a:gsLst></a:gradFill>`)).fill.stops

describe('a percentage written with a % suffix', () => {
  it('places a gradient stop where the file says', () => {
    expect(
      stopsOf(
        '<a:gs pos="0%"><a:srgbClr val="FF0000"/></a:gs>' +
          '<a:gs pos="50%"><a:srgbClr val="00FF00"/></a:gs>' +
          '<a:gs pos="100%"><a:srgbClr val="0000FF"/></a:gs>',
      ).map((stop) => stop.pos),
    ).toEqual([0, 0.5, 1])
  })

  it('keeps the integer form, where 50000 is also 50%', () => {
    expect(stopsOf('<a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>')[0]!.pos).toBe(0.5)
  })

  it('shrinks text by the font scale the file carries', () => {
    const scaled = first(
      shapeWithText('<a:normAutofit fontScale="62.5%" lnSpcReduction="20%"/>', ''),
    )
    expect(scaled.text.fontScale).toBeCloseTo(0.625)
    expect(scaled.text.lnSpcReduction).toBeCloseTo(0.2)
    const integer = first(
      shapeWithText('<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>', ''),
    )
    expect(integer.text.fontScale).toBeCloseTo(0.625)
  })

  it('spaces lines and paragraphs by the percentage the file carries', () => {
    const paragraph = first(
      shapeWithText(
        '',
        '<a:pPr><a:lnSpc><a:spcPct val="150%"/></a:lnSpc>' +
          '<a:spcBef><a:spcPct val="50%"/></a:spcBef></a:pPr>',
      ),
    ).text.paragraphs[0]
    expect(paragraph.lineHeight).toBe(150)
    expect(paragraph.spaceBeforePct).toBe(50)
  })

  it('offsets a superscript by the percentage the file carries', () => {
    const run = first(shapeWithText('', '<a:r><a:rPr baseline="50%"/><a:t>2</a:t></a:r>')).text
      .paragraphs[0].runs[0]
    expect(run.baseline).toBe(50)
  })

  it('applies colour modifiers that use it', () => {
    // 75% lumMod is the same as 75000, which already rendered #2F5496
    expect(applyColorMods('#4472C4', { 'a:lumMod': { '@_val': '75%' } } as any)).toBe(
      applyColorMods('#4472C4', { 'a:lumMod': { '@_val': '75000' } } as any),
    )
    expect(applyColorMods('#FF0000', { 'a:alpha': { '@_val': '50%' } } as any)).toBe('#FF000080')
  })
})

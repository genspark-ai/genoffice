import { describe, it, expect } from 'vitest'
import { resolveColorNode } from '../src/color'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const shapeWithFill = (fill: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="R"/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  `<a:prstGeom prst="rect"/>${fill}</p:spPr></p:sp>`

const fillOf = (fill: string): unknown =>
  (
    parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(shapeWithFill(fill)), ctx: {} })
      .elements[0] as any
  ).fill

describe('a:srgbClr with no resolvable val', () => {
  // '#' + String(undefined) is the truthy string '#UNDEFINED', so it slipped past the
  // `if (!base)` guard and hexToRgb parsed 'UN'/'FI' as 0,0 — painting a colour nobody authored.
  it('is unresolved when val is absent', () => {
    expect(resolveColorNode({ 'a:srgbClr': {} }, undefined)).toBeUndefined()
    expect(
      resolveColorNode({ 'a:srgbClr': { 'a:lumMod': { '@_val': '50000' } } }, undefined),
    ).toBeUndefined()
  })

  it('is unresolved when val is present but empty', () => {
    expect(resolveColorNode({ 'a:srgbClr': { '@_val': '' } }, undefined)).toBeUndefined()
  })

  it('matches the sibling prstClr/schemeClr missing-attribute branches', () => {
    expect(resolveColorNode({ 'a:prstClr': {} }, undefined)).toBeUndefined()
    expect(resolveColorNode({ 'a:schemeClr': {} }, undefined)).toBeUndefined()
  })

  it('still resolves a present val, with or without modifiers', () => {
    expect(resolveColorNode({ 'a:srgbClr': { '@_val': 'ff0000' } }, undefined)).toBe('#FF0000')
    expect(
      resolveColorNode(
        { 'a:srgbClr': { '@_val': 'FF0000', 'a:alpha': { '@_val': '50000' } } },
        undefined,
      ),
    ).toBe('#FF000080')
  })
})

describe('shape fill with an unresolvable srgbClr', () => {
  it('falls through to inheritance instead of painting an invented colour', () => {
    // A truncated/corrupt fill. It resolved to { type: 'solid', color: '#000000' } for the
    // empty val and '#00DE0F80' for the modifier-only form; the correct outcome is no
    // direct fill, so the shape inherits like any other unresolved fill.
    expect(fillOf('<a:solidFill><a:srgbClr val=""/></a:solidFill>')).toBeUndefined()
    expect(fillOf('<a:solidFill><a:srgbClr><a:alpha val="50000"/></a:srgbClr></a:solidFill>')).toBe(
      undefined,
    )
  })

  it('leaves a well-formed solidFill alone', () => {
    expect(fillOf('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')).toEqual({
      type: 'solid',
      color: '#FF0000',
    })
  })
})

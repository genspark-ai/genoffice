import { describe, expect, it } from 'vitest'
import { applyShapeWrapAt, applyShapeZOrderAt, buildShapeParagraphXml } from '../src/generate'

const SHAPE_XML = buildShapeParagraphXml({ prst: 'rect', withTextbox: true })
const POSITION_H = SHAPE_XML.match(/<wp:positionH[\s\S]*?<\/wp:positionH>/)![0]
const POSITION_V = SHAPE_XML.match(/<wp:positionV[\s\S]*?<\/wp:positionV>/)![0]

describe('shape wrap/rank patching', () => {
  it('switches a shape to behind without touching its position', () => {
    const out = applyShapeWrapAt(SHAPE_XML, 0, 'behind')
    expect(out).toContain('behindDoc="1"')
    expect(out).toContain('<wp:wrapNone/>')
    expect(out).toContain(POSITION_H)
    expect(out).toContain(POSITION_V)
  })

  it('switches a shape to square wrap and preserves its position', () => {
    const out = applyShapeWrapAt(SHAPE_XML, 0, 'square-left')
    expect(out).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(out).toContain(POSITION_H)
    expect(out).toContain(POSITION_V)
  })

  it('writes topBottom wrap together with a rank', () => {
    const out = applyShapeWrapAt(SHAPE_XML, 0, 'topBottom', 3)
    expect(out).toContain('<wp:wrapTopAndBottom/>')
    expect(out).toContain('relativeHeight="251658243"')
  })

  it('re-encodes only relativeHeight for a pure rank change', () => {
    const out = applyShapeZOrderAt(SHAPE_XML, 0, -2)
    expect(out).toContain('relativeHeight="251658238"')
    expect(out.replace('relativeHeight="251658238"', 'relativeHeight="251658240"')).toBe(SHAPE_XML)
  })

  it('converts an anchored shape to inline', () => {
    const out = applyShapeWrapAt(SHAPE_XML, 0, null)
    expect(out).toContain('<wp:inline')
    expect(out).not.toContain('wp:positionH')
    expect(out).not.toContain('wp:wrapSquare')
  })

  it('leaves the paragraph untouched for an out-of-range box index', () => {
    expect(applyShapeWrapAt(SHAPE_XML, 5, 'behind')).toBe(SHAPE_XML)
    expect(applyShapeZOrderAt(SHAPE_XML, 5, 3)).toBe(SHAPE_XML)
  })
})

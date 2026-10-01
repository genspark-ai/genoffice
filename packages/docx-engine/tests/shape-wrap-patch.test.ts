import { describe, expect, it } from 'vitest'
import { applyShapeWrapAt, applyShapeZOrderAt, buildShapeParagraphXml } from '../src/generate'

const SHAPE_XML = buildShapeParagraphXml({ prst: 'rect', withTextbox: true })
const POSITION_H = SHAPE_XML.match(/<wp:positionH[\s\S]*?<\/wp:positionH>/)![0]
const POSITION_V = SHAPE_XML.match(/<wp:positionV[\s\S]*?<\/wp:positionV>/)![0]

describe('shape wrap/rank patching', () => {
  it('switches a shape to behind without touching its position', () => {
    const out = applyShapeWrapAt(SHAPE_XML, { boxIndex: 0 }, 'behind')
    expect(out).toContain('behindDoc="1"')
    expect(out).toContain('<wp:wrapNone/>')
    expect(out).toContain(POSITION_H)
    expect(out).toContain(POSITION_V)
  })

  it('switches a shape to square wrap and preserves its position', () => {
    const out = applyShapeWrapAt(SHAPE_XML, { boxIndex: 0 }, 'square-left')
    expect(out).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(out).toContain(POSITION_H)
    expect(out).toContain(POSITION_V)
  })

  it('writes topBottom wrap together with a rank', () => {
    const out = applyShapeWrapAt(SHAPE_XML, { boxIndex: 0 }, 'topBottom', 3)
    expect(out).toContain('<wp:wrapTopAndBottom/>')
    expect(out).toContain('relativeHeight="251658243"')
  })

  it('re-encodes only relativeHeight for a pure rank change', () => {
    const out = applyShapeZOrderAt(SHAPE_XML, { boxIndex: 0 }, -2)
    expect(out).toContain('relativeHeight="251658238"')
    expect(out.replace('relativeHeight="251658238"', 'relativeHeight="251658240"')).toBe(SHAPE_XML)
  })

  it('converts an anchored shape to inline', () => {
    const out = applyShapeWrapAt(SHAPE_XML, { boxIndex: 0 }, null)
    expect(out).toContain('<wp:inline')
    expect(out).not.toContain('wp:positionH')
    expect(out).not.toContain('wp:wrapSquare')
  })

  it('removes a self-closing tight/through wrap element', () => {
    const through = SHAPE_XML.replace(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:wrapThrough wrapText="bothSides"/>',
    )
    const out = applyShapeWrapAt(through, { boxIndex: 0 }, 'behind')
    expect(out).toContain('<wp:wrapNone/>')
    expect(out).not.toContain('<wp:wrapThrough')
    expect(out.match(/<wp:wrap/g)).toHaveLength(1)
    const inline = applyShapeWrapAt(through, { boxIndex: 0 }, null)
    expect(inline).not.toContain('<wp:wrap')
    expect(inline).toContain('<wp:inline')
  })

  it('leaves the paragraph untouched for an out-of-range box index', () => {
    expect(applyShapeWrapAt(SHAPE_XML, { boxIndex: 5 }, 'behind')).toBe(SHAPE_XML)
    expect(applyShapeZOrderAt(SHAPE_XML, { boxIndex: 5 }, 3)).toBe(SHAPE_XML)
  })
})

describe('shape targeting by cNvPr id', () => {
  // a textless first shape is exactly the drawing boxDrawingSegments skips:
  // ordinal targeting would hit the second drawing, id targeting must not
  const withId = (xml: string, id: number): string =>
    xml.replace('<wps:cNvSpPr/>', `<wps:cNvPr id="${id}"/><wps:cNvSpPr/>`)
  const first = withId(buildShapeParagraphXml({ prst: 'rect', withTextbox: false }), 7)
  const second = withId(buildShapeParagraphXml({ prst: 'rect', withTextbox: true }), 8)
  const two =
    first.slice(0, first.indexOf('</w:p>')) + second.slice(second.indexOf('<w:p>') + '<w:p>'.length)
  const secondBody = second.slice(second.indexOf('<w:p>') + '<w:p>'.length)

  it('changes only the drawing carrying the requested id', () => {
    const out = applyShapeWrapAt(two, { shapeId: '7', boxIndex: 0 }, 'behind')
    expect(out.match(/behindDoc="1"/g)).toHaveLength(1)
    expect(out.match(/behindDoc="0"/g)).toHaveLength(1)
    expect(out.match(/<wp:wrapNone\/>/g)).toHaveLength(1)
    expect(out.match(/<wp:wrapSquare wrapText="bothSides"\/>/g)).toHaveLength(1)
    expect(out.endsWith(secondBody)).toBe(true)
  })

  it('re-encodes the rank of the drawing carrying the requested id', () => {
    const out = applyShapeZOrderAt(two, { shapeId: '7', boxIndex: 0 }, 3)
    expect([...out.matchAll(/relativeHeight="(\d+)"/g)].map((m) => m[1])).toEqual([
      '251658243',
      '251658240',
    ])
    expect(out.endsWith(secondBody)).toBe(true)
  })

  it('leaves the paragraph untouched for an unknown shape id', () => {
    expect(applyShapeWrapAt(two, { shapeId: '99', boxIndex: 0 }, 'behind')).toBe(two)
    expect(applyShapeZOrderAt(two, { shapeId: '99', boxIndex: 0 }, 3)).toBe(two)
  })
})

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

  it('addresses a photo drawing by its pic:cNvPr id, leaving the textbox drawing alone', () => {
    const photo =
      `<w:r><w:drawing><wp:anchor behindDoc="1" relativeHeight="251659264"><wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:posOffset>-1</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>-2</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="10" cy="10"/><wp:wrapNone/><wp:docPr id="9" name="Picture 9"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="Picture 9"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill><pic:spPr/></pic:pic>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
    const para = SHAPE_XML.replace(/<\/w:p>$/, `${photo}</w:p>`)
    const out = applyShapeWrapAt(para, { shapeId: '9', boxIndex: 1 }, 'front', 5)
    const photoOut = out.slice(
      out.indexOf('<wp:docPr id="9"') - 400,
      out.indexOf('<wp:docPr id="9"'),
    )
    expect(photoOut).toContain('behindDoc="0"')
    expect(photoOut).toContain('relativeHeight="251658245"')
    expect(out.startsWith(SHAPE_XML.replace(/<\/w:p>$/, ''))).toBe(true)
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

  it('removes a self-closing wrapTight element cleanly', () => {
    const tight = SHAPE_XML.replace(
      '<wp:wrapSquare wrapText="bothSides"/>',
      '<wp:wrapTight wrapText="bothSides"/>',
    )
    const out = applyShapeWrapAt(tight, { boxIndex: 0 }, 'behind')
    expect(out).toContain('<wp:wrapNone/>')
    expect(out).not.toContain('<wp:wrapTight')
    expect(out.match(/<wp:wrap/g)).toHaveLength(1)
  })

  it('leaves the paragraph untouched for an out-of-range box index', () => {
    expect(applyShapeWrapAt(SHAPE_XML, { boxIndex: 5 }, 'behind')).toBe(SHAPE_XML)
    expect(applyShapeZOrderAt(SHAPE_XML, { boxIndex: 5 }, 3)).toBe(SHAPE_XML)
  })
})

describe('shape targeting by cNvPr id', () => {
  // a textless first shape is exactly the drawing boxDrawingSegments skips:
  // ordinal targeting would hit the second drawing, id targeting must not
  const first = buildShapeParagraphXml({ prst: 'rect', withTextbox: false, id: 7 })
  const second = buildShapeParagraphXml({ prst: 'rect', withTextbox: true, id: 8 })
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

  it('does not throw on a shape id with regex metacharacters', () => {
    expect(applyShapeWrapAt(two, { shapeId: 'a(1', boxIndex: 0 }, 'behind')).toBe(two)
  })
})

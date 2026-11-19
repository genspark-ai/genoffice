/**
 * Group scale bound: a group's child coordinate system (<a:chOff>/<a:chExt>)
 * maps child EMU onto slide EMU by multiplying by ext/chExt, and that product
 * lands straight in the layout tree. The write path already bounds every
 * coordinate it emits (generate.ts posEmuAttr clamps a:ext to COORD_MAX, the
 * ST_PositiveCoordinate ceiling, and refuses non-finite values), so the read
 * path applies the same bound: a child coordinate system whose scale is not
 * representable in that range is dropped, leaving the 1:1 fallback every
 * consumer already handles. ext=2^31 with chExt=1 is the hostile case — an
 * ordinary 1e6 EMU child ends up ~2e11 px away.
 */
import { describe, it, expect } from 'vitest'
import { reassembleSlideXml } from '../src/index'
import { parseSlide } from '../src/parse'
import type { GroupElement } from '../src/types'

/** A 400-digit attribute parses to Infinity: out of the int64 range the schema allows. */
const OVER_INT64 = '1' + '0'.repeat(400)

const SP =
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr></p:sp>'

const groupSlideXml = (extCx: string, chCx: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extCx}" cy="${chCx}"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="${chCx}" cy="${chCx}"/></a:xfrm></p:grpSpPr>` +
  SP +
  '</p:grpSp>' +
  '</p:spTree></p:cSld></p:sld>'

/** A p:cxnSp child: what PowerPoint puts in a zero-height connector group. */
const CXNSP =
  '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
  '<p:spPr><a:xfrm><a:off x="1200000" y="2000000"/><a:ext cx="2600000" cy="0"/></a:xfrm>' +
  '<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>'

/** A group slide carrying one group whose a:xfrm is given verbatim (asymmetric axes). */
const groupSlideXmlWith = (xfrm: string, child = SP) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  `<p:grpSpPr>${xfrm}</p:grpSpPr>` +
  child +
  '</p:grpSp>' +
  '</p:spTree></p:cSld></p:sld>'

const groupOf = (slideXml: string): GroupElement => {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
  const el = slide.elements[0]
  expect(el?.type).toBe('group')
  return el as GroupElement
}

/**
 * The per-axis scale the model hands the layout tree, or null when the group
 * carries no child coordinate system (the 1:1 mapping consumers fall back to).
 * A zero axis is scale 1 in every consumer (pptx-ops: `ch?.cx ? … : 1`,
 * pptx-render: `ch?.cx || …`), which is what a dropped chOff gives too.
 */
const axisScales = (g: GroupElement): { x: number; y: number } | null =>
  g.childOffset
    ? {
        x: g.childOffset.cx > 0 ? g.transform.offset.cx / g.childOffset.cx : 1,
        y: g.childOffset.cy > 0 ? g.transform.offset.cy / g.childOffset.cy : 1,
      }
    : null

const childSpaceScale = (g: GroupElement): number | null => axisScales(g)?.x ?? null

describe('group child-coordinate scale bound', () => {
  it('drops a child coordinate system whose scale walks out of the write-path range', () => {
    // ext=2^31, chExt=1 → scale 2.1e9; a 1e6 EMU child lands ~2e11 px away.
    const g = groupOf(groupSlideXml('2147483648', '1'))

    expect(g.childOffset).toBeUndefined()
    expect(childSpaceScale(g)).toBeNull()
    // The group itself is untouched: children still parsed, bytes still replayed.
    expect(g.children).toHaveLength(1)
  })

  it('replays a group with an out-of-range scale verbatim (no bytes lost)', () => {
    const slideXml = groupSlideXml('2147483648', '1')
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })

    expect(reassembleSlideXml(slide)).toBe(slideXml)
  })

  it('keeps a real scaled group (2x) on its child coordinate system', () => {
    const g = groupOf(groupSlideXml('7620000', '3810000'))

    expect(g.childOffset).toEqual({ x: 0, y: 0, cx: 3810000, cy: 3810000 })
    expect(childSpaceScale(g)).toBe(2)
  })

  it('rejects a non-finite group extent instead of scaling by Infinity', () => {
    // parseInt of a 400-digit attribute is Infinity: the write path refuses to
    // emit it (clampInt sends non-finite to the floor), so the read path drops
    // the child coordinate system rather than hand Infinity to the layout tree.
    const g = groupOf(groupSlideXml(OVER_INT64, '1'))

    expect(g.childOffset).toBeUndefined()
    expect(childSpaceScale(g)).toBeNull()
  })

  it('rejects a non-finite child extent (no zero/infinite scale factor)', () => {
    const g = groupOf(groupSlideXml('2147483648', OVER_INT64))

    expect(g.childOffset).toBeUndefined()
    expect(childSpaceScale(g)).toBeNull()
  })

  it('keeps a zero-height connector group on its child coordinate system', () => {
    // PowerPoint writes ext cy=0 / chExt cy=0 for a horizontal connector group.
    // A zero axis is scale 1, but chOff still positions the child, so dropping the
    // child coordinate system would move the connector.
    const g = groupOf(
      groupSlideXmlWith(
        '<a:xfrm><a:off x="500000" y="4000000"/><a:ext cx="3000000" cy="0"/>' +
          '<a:chOff x="1000000" y="2000000"/><a:chExt cx="3000000" cy="0"/></a:xfrm>',
        CXNSP,
      ),
    )

    expect(g.childOffset).toEqual({ x: 1000000, y: 2000000, cx: 3000000, cy: 0 })
    expect(axisScales(g)).toEqual({ x: 1, y: 1 })
    // The cxnSp child is parsed as usual: a stroke-only shape, line preset.
    expect(g.children).toHaveLength(1)
    expect(g.children[0].type).toBe('shape')
    expect((g.children[0] as { presetGeometry?: string }).presetGeometry).toBe('line')
  })

  it('scales the healthy axis of a zero-height group and leaves the zero axis at 1', () => {
    const g = groupOf(
      groupSlideXmlWith(
        '<a:xfrm><a:off x="0" y="4000000"/><a:ext cx="7620000" cy="0"/>' +
          '<a:chOff x="0" y="2000000"/><a:chExt cx="3810000" cy="0"/></a:xfrm>',
        CXNSP,
      ),
    )

    // A zero cy must neither zero nor invalidate the 2x cx.
    expect(g.childOffset).toEqual({ x: 0, y: 2000000, cx: 3810000, cy: 0 })
    expect(axisScales(g)).toEqual({ x: 2, y: 1 })
  })

  it('still drops a group whose healthy axis overflows, next to a zero axis', () => {
    // The cy axis is judged on its own and cannot mask the cx overflow.
    const g = groupOf(
      groupSlideXmlWith(
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2147483648" cy="0"/>' +
          '<a:chOff x="0" y="0"/><a:chExt cx="1" cy="0"/></a:xfrm>',
        CXNSP,
      ),
    )

    expect(g.childOffset).toBeUndefined()
    expect(axisScales(g)).toBeNull()
  })

  it('rejects a non-finite chOff instead of mapping a child onto NaN', () => {
    const g = groupOf(
      groupSlideXmlWith(
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="3000000"/>' +
          `<a:chOff x="${OVER_INT64}" y="0"/><a:chExt cx="3000000" cy="3000000"/></a:xfrm>`,
      ),
    )

    expect(g.childOffset).toBeUndefined()
    expect(axisScales(g)).toBeNull()
  })

  it('keeps a zero child extent on both axes (scale 1 on each)', () => {
    // A zero chExt is legitimate input, not a reason to lose the child coordinate
    // system: every consumer already maps a zero axis to 1, the same mapping the
    // 1:1 fallback gives.
    const g = groupOf(groupSlideXml('2147483648', '0'))

    expect(g.childOffset).toEqual({ x: 0, y: 0, cx: 0, cy: 0 })
    expect(axisScales(g)).toEqual({ x: 1, y: 1 })
  })

  it('leaves a non-finite group box to the write-path clamp (both ends agree)', () => {
    // Not in scope here: parseXfrm carries the group's own out-of-int64 a:ext into
    // transform.offset for every element, and generate.ts's clampPosEmu already
    // neutralises it on write (clampInt sends non-finite to the floor). What must
    // not survive is a scale built on top of it.
    const g = groupOf(groupSlideXml(OVER_INT64, '1'))

    expect(g.transform.offset.cx).toBe(Infinity)
    expect(childSpaceScale(g)).toBeNull()
  })
})

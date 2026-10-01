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

const groupOf = (slideXml: string): GroupElement => {
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
  const el = slide.elements[0]
  expect(el?.type).toBe('group')
  return el as GroupElement
}

/**
 * The scale the model hands the layout tree, or null when the group carries no
 * child coordinate system (the 1:1 mapping consumers fall back to).
 */
const childSpaceScale = (g: GroupElement): number | null =>
  g.childOffset && g.childOffset.cx > 0 ? g.transform.offset.cx / g.childOffset.cx : null

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

  it('rejects a zero child extent (no zero denominator in the model)', () => {
    const g = groupOf(groupSlideXml('2147483648', '0'))

    expect(g.childOffset).toBeUndefined()
    expect(childSpaceScale(g)).toBeNull()
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

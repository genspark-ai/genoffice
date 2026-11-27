/**
 * ungroupElement + group rot (quarter turns) regression:
 *
 * OOXML applies the group's rot AFTER the chOff/chExt scale, turning the scaled
 * result about the group centre. ungroupElement used to ignore grp.transform.rot
 * entirely, so ungrouping a group with rot=90/180/270 put every child back at
 * the unrotated position (and dropped the group's rotation from the child's own
 * rot) — wrong x/y/cx/cy for any rotated group, and doubly wrong combined with
 * non-uniform ext/chExt scaling (the scaled extents must swap for 90/270).
 *
 * Expected values hand-derived by rotating the child box corners about the group
 * centre (rotation of a box = bbox of its rotated corners).
 */
import { describe, it, expect } from 'vitest'
import { appendRawElements, createBlankPptx, openPptx, ungroupElement } from '../src/index'
import type { GroupElement } from '../src/types'

const sp = (name: string, x: number, y: number, cx: number, cy: number, rot?: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm${rot ? ` rot="${rot}"` : ''}><a:off x="${x}" y="${y}"/>` +
  `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`

// rot in 60000ths on the group xfrm + non-uniform scale (ext 4000x8000 vs chExt 2000x2000)
const grpXml = (rot: number, childrenXml: string) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm rot="${rot}"><a:off x="1000" y="2000"/><a:ext cx="4000" cy="8000"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="2000" cy="2000"/></a:xfrm></p:grpSpPr>` +
  childrenXml +
  `</p:grpSp>`

async function ungroupRot(groupRot: number, childrenXml: string) {
  const opened = await openPptx(await createBlankPptx())
  const appended = appendRawElements(opened, 0, [grpXml(groupRot, childrenXml)])!
  const grpId = appended.elementIds[appended.elementIds.length - 1]!
  const fresh = ungroupElement(opened, 0, grpId)
  expect(fresh).not.toBeNull()
  return fresh!
}

describe('ungroupElement quarter-turn group rotation', () => {
  it('rot=90: child boxes rotate about the group centre, extents swap (non-uniform scale)', async () => {
    // scale x2/x4: children fill the top-left / bottom-left quadrant of the group
    const fresh = await ungroupRot(
      5400000,
      sp('A', 0, 0, 1000, 1000) + sp('B', 0, 1000, 1000, 1000),
    )
    const a = fresh.elements.find((e) => e.name === 'A')!
    const b = fresh.elements.find((e) => e.name === 'B')!
    // top-left quadrant rotated 90° CW lands in the top-right quadrant of the rotated bbox
    expect(a.transform.offset).toEqual({ x: 3000, y: 4000, cx: 4000, cy: 2000 })
    expect(a.transform.rot).toBe(5400000)
    expect(b.transform.offset).toEqual({ x: -1000, y: 4000, cx: 4000, cy: 2000 })
    expect(b.transform.rot).toBe(5400000)
  })

  it('rot=180: child boxes mirror about the group centre, extents keep', async () => {
    const fresh = await ungroupRot(10800000, sp('A', 0, 0, 1000, 1000))
    const a = fresh.elements.find((e) => e.name === 'A')!
    // top-left quadrant mirrored to the bottom-right quadrant of the group box
    expect(a.transform.offset).toEqual({ x: 3000, y: 6000, cx: 2000, cy: 4000 })
    expect(a.transform.rot).toBe(10800000)
  })

  it('rot=270: child boxes rotate about the group centre, extents swap', async () => {
    const fresh = await ungroupRot(16200000, sp('A', 0, 0, 1000, 1000))
    const a = fresh.elements.find((e) => e.name === 'A')!
    // top-left quadrant rotated 90° CCW lands in the bottom-left quadrant of the rotated bbox
    expect(a.transform.offset).toEqual({ x: -1000, y: 6000, cx: 4000, cy: 2000 })
    expect(a.transform.rot).toBe(16200000)
  })

  it('child rot adds the group rotation, normalized into 0..21600000', async () => {
    const fresh = await ungroupRot(
      5400000,
      sp('C90', 0, 0, 1000, 1000, 5400000) + sp('CNEG', 0, 1000, 1000, 1000, -5400000),
    )
    const c90 = fresh.elements.find((e) => e.name === 'C90')!
    const cneg = fresh.elements.find((e) => e.name === 'CNEG')!
    expect(c90.transform.rot).toBe(10800000) // 90° child + 90° group = 180°
    expect(cneg.transform.rot).toBe(0) // -90° + 90° = 0 → no rot attribute
    // position math is independent of the child's own rot
    expect(c90.transform.offset).toEqual({ x: 3000, y: 4000, cx: 4000, cy: 2000 })
  })

  it('non-quarter rot keeps the unrotated mapping (not exactly representable)', async () => {
    // rot=600000 (10°): falls back to the scale-only conversion, child rot unchanged
    const fresh = await ungroupRot(600000, sp('A', 0, 0, 1000, 1000))
    const a = fresh.elements.find((e) => e.name === 'A')!
    expect(a.transform.offset).toEqual({ x: 1000, y: 2000, cx: 2000, cy: 4000 })
    expect(a.transform.rot).toBe(0)
  })

  it('rot=0 groups keep the plain scale mapping', async () => {
    const fresh = await ungroupRot(0, sp('A', 0, 0, 1000, 1000))
    const a = fresh.elements.find((e) => e.name === 'A')!
    expect(a.transform.offset).toEqual({ x: 1000, y: 2000, cx: 2000, cy: 4000 })
    expect(a.transform.rot).toBe(0)
  })

  it('nested group child inherits the rotation too', async () => {
    const opened = await openPptx(await createBlankPptx())
    // outer group 1:1 (ext == chExt) so the inner group's kept-verbatim child
    // coordinate space keeps mapping its own children exactly
    const outer =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm rot="5400000"><a:off x="1000" y="2000"/><a:ext cx="1000" cy="1000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>` +
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="21" name="inner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>` +
      sp('X', 0, 0, 500, 500) +
      `</p:grpSp></p:grpSp>`
    const appended = appendRawElements(opened, 0, [outer])!
    const grpId = appended.elementIds[appended.elementIds.length - 1]!
    const fresh = ungroupElement(opened, 0, grpId)
    expect(fresh).not.toBeNull()
    // the inner group is lifted with the mapped box + accumulated rotation
    const innerEl = fresh!.elements.find((e) => e.type === 'group') as GroupElement
    expect(innerEl).toBeDefined()
    // inner fills the outer box → rotated 90° CW with it, extents swap
    expect(innerEl.transform.offset).toEqual({ x: 1000, y: 2000, cx: 1000, cy: 1000 })
    expect(innerEl.transform.rot).toBe(5400000)
    // its own child space is kept verbatim: the inner group's rot carries the turn
    expect(innerEl.children[0]!.transform.offset).toEqual({ x: 0, y: 0, cx: 500, cy: 500 })
    expect(innerEl.children[0]!.transform.rot).toBe(0)
  })
})

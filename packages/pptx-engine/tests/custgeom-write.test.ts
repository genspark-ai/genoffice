import { describe, it, expect } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  savePptx,
  setShapeAdjustValues,
  setShapeCustomGeometry,
  custGeomXml,
  parseCustGeom,
  validCustGeomPath,
  type CustGeomPath,
} from '../src/index'
import type { TextElement } from '../src/types'

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }

const tri: CustGeomPath = {
  w: OFF.cx,
  h: OFF.cy,
  cmds: [
    { op: 'M', pts: [0, OFF.cy] },
    { op: 'L', pts: [OFF.cx / 2, 0] },
    { op: 'C', pts: [OFF.cx * 0.75, 0, OFF.cx, OFF.cy / 2, OFF.cx, OFF.cy] },
    { op: 'Q', pts: [OFF.cx / 2, OFF.cy * 1.2, 0, OFF.cy] },
    { op: 'Z', pts: [] },
  ],
}

describe('custGeomXml', () => {
  it('emits a single-path custGeom whose parse equals the input', () => {
    const xml = custGeomXml(tri)
    expect(xml).toContain('<a:path w="1828800" h="914400">')
    expect(xml).toContain('<a:moveTo><a:pt x="0" y="914400"/></a:moveTo>')
    expect(xml).toContain('<a:cubicBezTo>')
    expect(xml).toContain('<a:quadBezTo>')
    expect(xml).toMatch(/<a:close\/><\/a:path>/)
    const geo = parseCustGeom(`<p:spPr>${xml}</p:spPr>`, OFF.cx, OFF.cy)
    expect(geo?.path).toBe('M 0 1 L 0.5 0 C 0.75 0 1 0.5 1 1 Q 0.5 1.2 0 1 Z')
    expect(geo?.fillPath).toBeUndefined()
    expect(geo?.strokePath).toBeUndefined()
  })

  it('rounds coordinates to integers', () => {
    const xml = custGeomXml({ w: 10.4, h: 10.6, cmds: [{ op: 'M', pts: [1.5, 2.49] }] })
    expect(xml).toContain('<a:path w="10" h="11">')
    expect(xml).toContain('<a:pt x="2" y="2"/>')
  })
})

describe('validCustGeomPath', () => {
  it('requires a leading moveTo, matching point counts and finite numbers', () => {
    expect(validCustGeomPath(tri)).toBe(true)
    expect(validCustGeomPath({ ...tri, cmds: [] })).toBe(false)
    expect(validCustGeomPath({ ...tri, cmds: [{ op: 'L', pts: [0, 0] }] })).toBe(false)
    expect(validCustGeomPath({ ...tri, cmds: [{ op: 'M', pts: [0] }] })).toBe(false)
    expect(validCustGeomPath({ ...tri, cmds: [{ op: 'M', pts: [0, Number.NaN] }] })).toBe(false)
    expect(validCustGeomPath({ ...tri, w: 0 })).toBe(false)
    expect(validCustGeomPath({ ...tri, cmds: [{ op: 'X', pts: [] }] })).toBe(false)
  })
})

describe('setShapeCustomGeometry (Edit Points)', () => {
  it('replaces the preset, clears adjust, and round-trips through save', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'roundRect', offset: { ...OFF } })
    setShapeAdjustValues(slide, el.id, { adj: 33333 })

    expect(setShapeCustomGeometry(slide, el.id, tri)).toBe(true)
    const shape = el as TextElement
    expect(shape.presetGeometry).toBeUndefined()
    expect(shape.adjust).toBeUndefined()
    expect(shape.customGeometry?.path).toBe('M 0 1 L 0.5 0 C 0.75 0 1 0.5 1 1 Q 0.5 1.2 0 1 Z')
    expect(el.anchor.originalXml).not.toContain('<a:prstGeom')
    expect(el.anchor.originalXml.match(/<a:custGeom>/g)).toHaveLength(1)

    const reopened = await openPptx(await savePptx(opened))
    const saved = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(saved.presetGeometry).toBeUndefined()
    expect(saved.customGeometry).toEqual(shape.customGeometry)
  })

  it('rewrites an existing custGeom in place', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    expect(setShapeCustomGeometry(slide, el.id, tri)).toBe(true)
    const moved: CustGeomPath = {
      ...tri,
      cmds: tri.cmds.map((c, i) => (i === 1 ? { op: 'L', pts: [OFF.cx / 4, 0] } : c)),
    }
    expect(setShapeCustomGeometry(slide, el.id, moved)).toBe(true)
    expect((el as TextElement).customGeometry?.path).toContain('L 0.25 0')
    expect(el.anchor.originalXml.match(/<a:custGeom>/g)).toHaveLength(1)
  })

  it('rejects unknown ids, non-shape elements and invalid paths', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    expect(setShapeCustomGeometry(slide, 'nope', tri)).toBe(false)
    expect(setShapeCustomGeometry(slide, el.id, { ...tri, cmds: [] })).toBe(false)
    expect((el as TextElement).presetGeometry).toBe('rect')
  })
})

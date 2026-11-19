/** Connector attachment parsing + endpoint follow after a connected shape moves. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { updateConnectorsForMoved } from '../src/index'

const SP = (id: number, x: number, y: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="1000" cy="600"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'

const CXN =
  '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9" name="C"/>' +
  '<p:cNvCxnSpPr><a:stCxn id="2" idx="3"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>' +
  '<p:spPr><a:xfrm><a:off x="1000" y="300"/><a:ext cx="1000" cy="0"/></a:xfrm>' +
  '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>'

const slideXml =
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  `<p:nvGrpSpPr/><p:grpSpPr/>${SP(2, 0, 0)}${SP(3, 2000, 0)}${CXN}</p:spTree></p:cSld></p:sld>`

describe('connector following', () => {
  it('parses stCxn/endCxn attachment relations', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const cxn = slide.elements.find((e) => e.connection)!
    expect(cxn.connection).toEqual({ start: { id: 2, idx: 3 }, end: { id: 3, idx: 1 } })
  })

  it('endpoints reflow after a connected shape moves (idx 3=right-edge midpoint → idx 1=left-edge midpoint)', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const s3 = slide.elements.find((e) => e.name === 'S3')!
    // Move S3 to (4000, 2000)
    s3.transform.offset = { x: 4000, y: 2000, cx: 1000, cy: 600 }
    const n = updateConnectorsForMoved(slide, [s3.id])
    expect(n).toBe(1)
    const cxn = slide.elements.find((e) => e.connection)!
    // Start = midpoint of S2's right edge (1000, 300); end = midpoint of S3's left edge (4000, 2300)
    expect(cxn.transform.offset).toEqual({ x: 1000, y: 300, cx: 3000, cy: 2000 })
    expect(cxn.transform.flipH).toBe(false)
    expect(cxn.transform.flipV).toBe(false)
    expect(cxn.dirtyTransform).toBe(true)
  })

  it('flip expresses direction when the target moves above-left of the start', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const s3 = slide.elements.find((e) => e.name === 'S3')!
    s3.transform.offset = { x: -3000, y: -2000, cx: 1000, cy: 600 }
    updateConnectorsForMoved(slide, [s3.id])
    const cxn = slide.elements.find((e) => e.connection)!
    expect(cxn.transform.flipH).toBe(true)
    expect(cxn.transform.flipV).toBe(true)
  })

  it('a rotated target snaps to the rotated side (90deg rect: left midpoint lands on the visual top)', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const s3 = slide.elements.find((e) => e.name === 'S3')!
    s3.transform.offset = { x: 4000, y: 2000, cx: 1000, cy: 600 }
    s3.transform.rot = 90 * 60000
    updateConnectorsForMoved(slide, [s3.id])
    const cxn = slide.elements.find((e) => e.connection)!
    // centre (4500, 2300); left midpoint offset (-500, 0) rotates to (0, -500)
    expect(cxn.transform.offset).toEqual({ x: 1000, y: 300, cx: 3500, cy: 1500 })
    expect(cxn.transform.flipH).toBe(false)
    expect(cxn.transform.flipV).toBe(false)
  })

  it('flipH on the target mirrors the site to the opposite edge', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const s3 = slide.elements.find((e) => e.name === 'S3')!
    s3.transform.offset = { x: 4000, y: 2000, cx: 1000, cy: 600 }
    s3.transform.flipH = true
    updateConnectorsForMoved(slide, [s3.id])
    const cxn = slide.elements.find((e) => e.connection)!
    // idx 1 (left) mirrored → right edge midpoint (5000, 2300)
    expect(cxn.transform.offset).toEqual({ x: 1000, y: 300, cx: 4000, cy: 2000 })
  })

  it('moving an unattached element does not affect connectors', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const cxnBefore = JSON.stringify(slide.elements.find((e) => e.connection)!.transform)
    expect(updateConnectorsForMoved(slide, ['nonexistent'])).toBe(0)
    expect(JSON.stringify(slide.elements.find((e) => e.connection)!.transform)).toBe(cxnBefore)
  })
})

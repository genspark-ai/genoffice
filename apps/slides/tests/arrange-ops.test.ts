/**
 * alignElements / distributeElements / addConnector, built-in table styles by
 * name, and layout-by-name slide ops — against a real in-memory deck.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  addElement,
  addTable,
  createBlankPptx,
  openPptx,
  savePptx,
  type OpenedPptx,
  type SlideElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'

let opened: OpenedPptx
let a: string
let b: string
let c: string

const box = (x: number, y: number, cx = 914400, cy = 457200) => ({ x, y, cx, cy })
const slide = () => opened.deck.slides[0]!
const find = (id: string): SlideElement => slide().elements.find((e) => e.id === id)!
const apply = (ops: object[]) => runTxn(opened, { ops: ops as never })

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
  a = addElement(slide(), { kind: 'rect', offset: box(0, 0) }).id
  b = addElement(slide(), { kind: 'ellipse', offset: box(3000000, 1000000, 600000, 600000) }).id
  c = addElement(slide(), { kind: 'roundRect', offset: box(6000000, 2500000) }).id
})

describe('alignElements', () => {
  it('aligns to the selection bounding box and to the slide', () => {
    const r = apply([{ op: 'alignElements', target: { slide: 0 }, els: [a, b, c], mode: 'top' }])
    expect(r.applied).toBe(true)
    expect([a, b, c].map((id) => find(id).transform.offset.y)).toEqual([0, 0, 0])
    expect(find(b).dirtyTransform).toBe(true)

    const right = apply([
      { op: 'alignElements', target: { slide: 0 }, els: [a], mode: 'right', to: 'slide' },
    ])
    expect(right.applied).toBe(true)
    const o = find(a).transform.offset
    expect(o.x + o.cx).toBe(opened.deck.size.cx)
  })

  it('needs two elements for selection alignment and refuses group children', () => {
    const one = apply([{ op: 'alignElements', target: { slide: 0 }, els: [a], mode: 'left' }])
    expect(one.applied).toBe(false)
    expect(one.failures![0]!.error).toContain('at least 2')
    const grouped = apply([{ op: 'groupElements', target: { slide: 0 }, els: [a, b] }])
    expect(grouped.applied).toBe(true)
    const childId = (
      slide().elements.find((e) => e.type === 'group') as { children: SlideElement[] }
    ).children[0]!.id
    const r = apply([
      { op: 'alignElements', target: { slide: 0 }, els: [childId, c], mode: 'left' },
    ])
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('inside a group')
  })
})

describe('distributeElements', () => {
  it('keeps the outer elements and evens the inner gaps', () => {
    const r = apply([
      { op: 'distributeElements', target: { slide: 0 }, els: [a, b, c], axis: 'horizontal' },
    ])
    expect(r.applied).toBe(true)
    const [oa, ob, oc] = [a, b, c].map((id) => find(id).transform.offset)
    expect(oa!.x).toBe(0)
    expect(oc!.x).toBe(6000000)
    const gap1 = ob!.x - (oa!.x + oa!.cx)
    const gap2 = oc!.x - (ob!.x + ob!.cx)
    expect(Math.abs(gap1 - gap2)).toBeLessThanOrEqual(1)
    expect(ob!.y).toBe(1000000)
  })

  it('to slide spaces the edges evenly and accepts fewer than three', () => {
    const r = apply([
      {
        op: 'distributeElements',
        target: { slide: 0 },
        els: [a, c],
        axis: 'vertical',
        to: 'slide',
      },
    ])
    expect(r.applied).toBe(true)
    const [oa, oc] = [a, c].map((id) => find(id).transform.offset)
    const gap = (opened.deck.size.cy - oa!.cy - oc!.cy) / 3
    expect(Math.abs(oa!.y - gap)).toBeLessThanOrEqual(1)
    expect(Math.abs(oc!.y - (oa!.y + oa!.cy + gap))).toBeLessThanOrEqual(1)
    const two = apply([
      { op: 'distributeElements', target: { slide: 0 }, els: [a, c], axis: 'vertical' },
    ])
    expect(two.applied).toBe(false)
    expect(two.failures![0]!.error).toContain('at least 3')
  })
})

describe('addConnector', () => {
  it('glues the closest sides, writes arrows and follows a moved ellipse', async () => {
    const r = apply([
      {
        op: 'addConnector',
        target: { slide: 0 },
        from: a,
        to: b,
        arrow: 'both',
        line: { color: '#FF0000', widthPt: 2, dash: 'dash' },
      },
    ])
    expect(r.applied).toBe(true)
    const rec = r.records![0]!
    expect(rec.after).toMatchObject({ fromSide: 'right', toSide: 'left' })
    const cxn = find(rec.created![0]!)
    const xml = cxn.anchor.originalXml
    expect(xml).toContain('<p:cxnSp>')
    expect(xml).toMatch(/<a:stCxn id="\d+" idx="3"\/>/)
    // ellipse edge sites sit at even indexes: left = 2
    expect(xml).toMatch(/<a:endCxn id="\d+" idx="2"\/>/)
    expect(xml).toContain('<a:headEnd type="triangle"')
    expect(xml).toContain('<a:tailEnd type="triangle"')
    expect(xml).toContain('<a:prstDash val="dash"/>')
    expect(xml).toContain('w="25400"')
    expect(xml).toContain('val="FF0000"')
    const o = cxn.transform.offset
    expect(o.x).toBe(914400)
    expect(o.y).toBe(228600)
    expect(o.cx).toBe(3000000 - 914400)

    // move the ellipse: the connector end re-lays to its new left midpoint
    const moved = apply([
      {
        op: 'setTransform',
        target: { slide: 0, el: b },
        box: box(4000000, 2000000, 600000, 600000),
      },
    ])
    expect(moved.applied).toBe(true)
    const after = find(cxn.id).transform.offset
    expect(after.x + after.cx).toBe(4000000)
    expect(after.y + after.cy).toBe(2300000)

    const reopened = await openPptx(await savePptx(opened))
    const saved = reopened.deck.slides[0]!.elements.find((e) => e.connection)!
    expect(saved.connection).toMatchObject({ start: { idx: 3 }, end: { idx: 2 } })
  })

  it('honours pinned sides and refuses bad endpoints', () => {
    const r = apply([
      {
        op: 'addConnector',
        target: { slide: 0 },
        from: a,
        to: c,
        kind: 'elbow',
        fromSide: 'bottom',
        arrow: 'none',
      },
    ])
    expect(r.applied).toBe(true)
    expect(r.records![0]!.after).toMatchObject({ fromSide: 'bottom' })
    const xml = find(r.records![0]!.created![0]!).anchor.originalXml
    expect(xml).toContain('prst="bentConnector3"')
    expect(xml).not.toContain('tailEnd')
    // right-to-left: the flip is baked into the saved xfrm, not just the model
    const back = apply([{ op: 'addConnector', target: { slide: 0 }, from: c, to: a }])
    expect(back.applied).toBe(true)
    const backXml = find(back.records![0]!.created![0]!).anchor.originalXml
    expect(backXml).toMatch(/<a:xfrm[^>]*flipH="1"/)
    expect(backXml).toMatch(/<a:xfrm[^>]*flipV="1"/)
    const same = apply([{ op: 'addConnector', target: { slide: 0 }, from: a, to: a }])
    expect(same.failures![0]!.error).toContain('must differ')
    const missing = apply([{ op: 'addConnector', target: { slide: 0 }, from: a, to: 'nope' }])
    expect(missing.failures![0]!.error).toContain('Available:')
    const side = apply([
      { op: 'addConnector', target: { slide: 0 }, from: a, to: b, toSide: 'middle' },
    ])
    expect(side.failures![0]!.error).toContain('top/left/bottom/right')
  })
})

describe('setTableStyle with a built-in style', () => {
  it('writes the GUID and region flags, reads the name back after save', async () => {
    const tbl = {
      id: addTable(opened, 0, { rows: 3, cols: 2, offset: box(0, 3000000, 3000000, 1000000) })!
        .elementId,
    }
    const r = apply([
      {
        op: 'setTableStyle',
        target: { slide: 0, el: tbl.id },
        styleId: 'medium style 2 - accent 1',
        firstRow: true,
        bandRow: true,
        lastRow: true,
      },
    ])
    expect(r.applied).toBe(true)
    const xml = find(tbl.id).anchor.originalXml
    expect(xml).toContain('<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>')
    expect(xml).toMatch(/<a:tblPr[^>]*firstRow="1"/)
    expect(xml).toMatch(/<a:tblPr[^>]*lastRow="1"/)
    expect(xml).toMatch(/<a:tblPr[^>]*bandRow="1"/)
    expect((xml.match(/<a:tableStyleId>/g) ?? []).length).toBe(1)

    const reopened = await openPptx(await savePptx(opened))
    const saved = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table')!
    expect(saved).toMatchObject({
      styleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
      styleFlags: { firstRow: true, bandRow: true, lastRow: true, firstCol: false },
    })
  })

  it('lists the gallery when the name is unknown and accepts a GUID', () => {
    const tbl = {
      id: addTable(opened, 0, { rows: 2, cols: 2, offset: box(0, 3000000) })!.elementId,
    }
    const bad = apply([
      { op: 'setTableStyle', target: { slide: 0, el: tbl.id }, styleId: 'Medium Style 9' },
    ])
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('Light Style 1 - Accent 1')
    const guid = apply([
      {
        op: 'setTableStyle',
        target: { slide: 0, el: tbl.id },
        styleId: '{e8034e78-7f5d-4c2e-b375-fc64b27bc917}',
        keepFormatting: true,
      },
    ])
    expect(guid.applied).toBe(true)
    expect(find(tbl.id).anchor.originalXml).toContain('{E8034E78-7F5D-4C2E-B375-FC64B27BC917}')
  })
})

describe('slides by layout name', () => {
  it('addSlideWithLayout by name or index, setSlideLayout by name', () => {
    const r = apply([{ op: 'addSlideWithLayout', layout: 'blank' }])
    expect(r.applied).toBe(true)
    expect(opened.deck.slides).toHaveLength(2)
    expect(r.records![0]!.created![0]).toMatch(/^s_/)
    const byIndex = apply([{ op: 'addSlideWithLayout', target: { slide: 0 }, layout: 0 }])
    expect(byIndex.applied).toBe(true)
    expect(byIndex.records![0]!.after).toMatchObject({ index: 1 })
    expect(opened.deck.slides).toHaveLength(3)
    const bad = apply([{ op: 'addSlideWithLayout', layout: 'Title and Content' }])
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('0: "Blank"')
    const oor = apply([{ op: 'addSlideWithLayout', layout: 4 }])
    expect(oor.failures![0]!.error).toContain('out of range (0-0)')
    const relink = apply([{ op: 'setSlideLayout', target: { slide: 1 }, layout: 'Blank' }])
    expect(relink.applied).toBe(true)
    expect(relink.records![0]!.after).toBe('ppt/slideLayouts/slideLayout1.xml')
  })
})

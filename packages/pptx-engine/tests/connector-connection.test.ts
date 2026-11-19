/**
 * Phase 2: Connector shape attachments (a:stCxn/a:endCxn) —
 * setElementConnection byte surgery + move-following via updateConnectorsForMoved.
 */
import { describe, it, expect } from 'vitest'
import {
  openPptx,
  savePptx,
  createBlankPptx,
  setElementConnection,
  updateConnectorsForMoved,
  elementSpid,
} from '../src/index'
import { addElement } from '../src/insert'
import type { TextElement } from '../src/types'

async function setup() {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const a = addElement(slide, {
    kind: 'rect',
    offset: { x: 0, y: 0, cx: 1000000, cy: 500000 },
    fillColor: '#4472C4',
  })
  const b = addElement(slide, {
    kind: 'rect',
    offset: { x: 3000000, y: 2000000, cx: 1000000, cy: 500000 },
    fillColor: '#ED7D31',
  })
  const cxn = addElement(slide, {
    kind: 'lineArrow',
    offset: { x: 1000000, y: 250000, cx: 2000000, cy: 2000000 },
  })
  return { opened, slide, a, b, cxn }
}

describe('setElementConnection', () => {
  it('writes stCxn/endCxn into cNvCxnSpPr and updates the model', async () => {
    const { slide, a, b, cxn } = await setup()
    const aId = elementSpid(a)!
    const bId = elementSpid(b)!
    expect(
      setElementConnection(slide, cxn.id, {
        start: { id: aId, idx: 3 },
        end: { id: bId, idx: 0 },
      }),
    ).toBe(true)
    const xml = cxn.anchor.originalXml
    expect(xml).toMatch(
      new RegExp(
        `<p:cNvCxnSpPr[^>]*><a:stCxn id="${aId}" idx="3"/><a:endCxn id="${bId}" idx="0"/></p:cNvCxnSpPr>`,
      ),
    )
    expect(cxn.connection).toEqual({ start: { id: aId, idx: 3 }, end: { id: bId, idx: 0 } })
  })

  it('undefined keeps the other end, null detaches', async () => {
    const { slide, a, b, cxn } = await setup()
    const aId = elementSpid(a)!
    const bId = elementSpid(b)!
    setElementConnection(slide, cxn.id, { start: { id: aId, idx: 3 }, end: { id: bId, idx: 0 } })
    // Re-attach only the end; start untouched
    setElementConnection(slide, cxn.id, { end: { id: bId, idx: 1 } })
    expect(cxn.connection).toEqual({ start: { id: aId, idx: 3 }, end: { id: bId, idx: 1 } })
    expect(cxn.anchor.originalXml).toContain(`<a:stCxn id="${aId}" idx="3"/>`)
    expect(cxn.anchor.originalXml).toContain(`<a:endCxn id="${bId}" idx="1"/>`)
    // Detach the start
    setElementConnection(slide, cxn.id, { start: null })
    expect(cxn.connection).toEqual({ end: { id: bId, idx: 1 } })
    expect(cxn.anchor.originalXml).not.toContain('<a:stCxn')
  })

  it('attached connector follows the shape after updateConnectorsForMoved', async () => {
    const { slide, a, b, cxn } = await setup()
    setElementConnection(slide, cxn.id, {
      start: { id: elementSpid(a)!, idx: 3 },
      end: { id: elementSpid(b)!, idx: 0 },
    })
    b.transform.offset = { x: 5000000, y: 4000000, cx: 1000000, cy: 500000 }
    b.dirtyTransform = true
    expect(updateConnectorsForMoved(slide, [b.id])).toBe(1)
    // start = right midpoint of a (1000000, 250000); end = top midpoint of b (5500000, 4000000)
    expect(cxn.transform.offset).toEqual({ x: 1000000, y: 250000, cx: 4500000, cy: 3750000 })
    expect(cxn.transform.flipH).toBe(false)
    expect(cxn.transform.flipV).toBe(false)
    expect(cxn.dirtyTransform).toBe(true)
  })

  it('round-trips through save: stCxn/endCxn reparsed into connection', async () => {
    const { opened, slide, a, b, cxn } = await setup()
    const aId = elementSpid(a)!
    const bId = elementSpid(b)!
    setElementConnection(slide, cxn.id, { start: { id: aId, idx: 3 }, end: { id: bId, idx: 0 } })
    const reopened = await openPptx(Buffer.from(await savePptx(opened)))
    const parsed = reopened.deck.slides[0]!.elements.find(
      (e) => (e as TextElement).presetGeometry === 'straightConnector1',
    )
    expect(parsed?.connection).toEqual({ start: { id: aId, idx: 3 }, end: { id: bId, idx: 0 } })
  })
})

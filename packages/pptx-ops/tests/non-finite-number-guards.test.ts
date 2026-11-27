/**
 * Non-finite number guards: NaN/Infinity pass a bare `typeof v === 'number'`
 * check and would reach integer coercion or generated attributes verbatim —
 * rot="-2147483648" (Math.round(Infinity) clamped by angleAttr) or x="NaN"
 * pasted straight into a:off. Numeric op fields must be rejected at plan time
 * like every other requireFinite site (registry.ts).
 */
import { describe, it, expect, beforeAll } from 'vitest'
import {
  addElement,
  copyElementData,
  createBlankPptx,
  openPptx,
  patchSlideXml,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let shapeId: string
let clipItem: ReturnType<typeof copyElementData>

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const shape = addElement(slide, {
    kind: 'rect',
    offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
  })
  shapeId = shape.id
  clipItem = copyElementData(opened, slide, shape)
})

describe('setTransform rotDeg', () => {
  it.each([Infinity, -Infinity, NaN])('rejects non-finite rotDeg %s', (rotDeg) => {
    const slide = opened.deck.slides[0]!
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setTransform',
          target: { slide: 0, el: shapeId },
          box: { x: 200000, y: 200000, cx: 1000000, cy: 500000 },
          rotDeg,
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.error).toMatch(/rotDeg/)
    expect(r.failures![0]!.error).toMatch(/finite/)
    // the corrupt rot must not exist in the model or the serialized slide
    const el = slide.elements.find((x) => x.id === shapeId)!
    expect(Number.isFinite(el.transform.rot)).toBe(true)
    expect(patchSlideXml(slide)).not.toContain('rot="-2147483648"')
    expect(patchSlideXml(slide)).not.toMatch(/rot="(?:NaN|Infinity|-Infinity)"/)
  })

  it('still applies a finite rotDeg (60000ths-of-a-degree coercion intact)', () => {
    const slide = opened.deck.slides[0]!
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setTransform',
          target: { slide: 0, el: shapeId },
          box: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
          rotDeg: 45,
        },
      ],
    })
    expect(r.applied).toBe(true)
    const el = slide.elements.find((x) => x.id === shapeId)!
    expect(el.transform.rot).toBe(Math.round(45 * 60000))
  })
})

describe('pasteElements dx/dy', () => {
  const pasteOp = (dx: number, dy: number) => ({
    op: 'pasteElements',
    target: { slide: 0 },
    items: [clipItem],
    dx,
    dy,
  })

  it.each([
    ['dx NaN', NaN, 0],
    ['dy NaN', 0, NaN],
    ['dx Infinity', Infinity, 0],
    ['dy -Infinity', 0, -Infinity],
  ])('rejects %s instead of writing it into a:off', (_label, dx, dy) => {
    const slide = opened.deck.slides[0]!
    const countBefore = slide.elements.length
    const r = runTxn(opened, { ops: [pasteOp(dx, dy)] })
    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.error).toMatch(/dx.*dy/)
    expect(r.failures![0]!.error).toMatch(/finite/)
    expect(slide.elements.length).toBe(countBefore)
    expect(patchSlideXml(slide)).not.toMatch(/x="(?:NaN|Infinity|-Infinity)"/)
    expect(patchSlideXml(slide)).not.toMatch(/y="(?:NaN|Infinity|-Infinity)"/)
  })

  it('still pastes with finite offsets', () => {
    const r = runTxn(opened, { ops: [pasteOp(50000, 75000)] })
    expect(r.applied).toBe(true)
    expect(r.records![0]!.created).toHaveLength(1)
    expect(patchSlideXml(opened.deck.slides[0]!)).not.toMatch(/x="(?:NaN|Infinity)"/)
  })
})

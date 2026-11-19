/**
 * Drawing (freehand ink) round-trip verification:
 * one stroke = one transparent PNG picture element (cNvPr name carries the aislides-ink prefix;
 * descr stores the vector-point JSON). Verifies the add-ink IPC's engine path (addPicture with
 * name/descr), descr surviving save-and-reopen, render-node passthrough, and editor-side payload
 * encode/decode plus hit testing.
 */
import { describe, it, expect } from 'vitest'
import { addPicture, createBlankPptx, openPptx, savePptx } from '@genoffice/pptx-engine'
import { buildRenderSlide, type PictureRenderNode } from '@genoffice/pptx-render'
import {
  decodeInkPayload,
  encodeInkPayload,
  inkNodeHitTest,
  inkNodesOf,
  strokeBounds,
  strokeHitTest,
  strokePathD,
  INK_NAME_PREFIX,
  type InkStroke,
} from '../src/renderer/ink'

const FIT = 1280

// 1x1 transparent PNG
const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64',
)

const stroke: InkStroke = {
  tool: 'pen',
  color: 'C00000',
  width: 2,
  points: [
    { x: 0, y: 0 },
    { x: 40, y: 10 },
    { x: 80, y: 0 },
  ],
}

describe('ink payload (cNvPr descr) encode/decode', () => {
  it('round-trips tool/color/width/points', () => {
    const decoded = decodeInkPayload(encodeInkPayload(stroke))!
    expect(decoded.tool).toBe('pen')
    expect(decoded.color).toBe('C00000')
    expect(decoded.width).toBe(2)
    expect(decoded.points).toHaveLength(3)
    expect(decoded.points[1]).toEqual({ x: 40, y: 10 })
  })

  it('rejects foreign/broken payloads', () => {
    expect(decodeInkPayload('not json')).toBeNull()
    expect(decodeInkPayload('{"v":99,"points":[[0,0]]}')).toBeNull()
    expect(decodeInkPayload('{"v":1,"points":[]}')).toBeNull()
  })
})

describe('ink geometry', () => {
  it('strokeBounds pads by half width + 1', () => {
    const b = strokeBounds(stroke)
    expect(b.x).toBe(-2)
    expect(b.y).toBe(-2)
    expect(b.w).toBe(84)
    expect(b.h).toBe(14)
  })

  it('strokeHitTest hits near segments, misses far away', () => {
    expect(strokeHitTest(stroke, { x: 20, y: 5 }, 6)).toBe(true)
    expect(strokeHitTest(stroke, { x: 40, y: 60 }, 6)).toBe(false)
  })

  it('strokePathD renders a dot for single points', () => {
    expect(strokePathD([{ x: 3, y: 4 }])).toContain('M 3 4')
    expect(strokePathD(stroke.points)).toMatch(/^M 0 0 L 40 10/)
  })
})

describe('ink stroke → save → reopen (engine path = slides:add-ink handler)', () => {
  it('addPicture with name/descr survives save + reopen and reaches render nodes', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!

    // The submit side has already converted points relative to the bounding box's top-left
    const bounds = strokeBounds(stroke)
    const rel = {
      ...stroke,
      points: stroke.points.map((p) => ({ x: p.x - bounds.x, y: p.y - bounds.y })),
    }
    const payload = encodeInkPayload(rel)

    const el = addPicture(opened, slide, {
      bytes: new Uint8Array(PNG_1PX),
      ext: 'png',
      offset: { x: 914400, y: 914400, cx: 762000, cy: 127000 },
      name: `${INK_NAME_PREFIX} t1`,
      descr: payload,
    })!
    expect(el.name).toContain(INK_NAME_PREFIX)
    expect(el.descr).toBe(payload)

    // Save -> reopen: descr/name must survive (the eraser can still recognize the stroke after reopening)
    const reopened = await openPptx(await savePptx(opened))
    const pic = reopened.deck.slides[0]!.elements.find((e) => e.type === 'picture')!
    expect(pic.name).toContain(INK_NAME_PREFIX)
    expect(pic.descr).toBe(payload)

    // Render-node passthrough + inkNodesOf recognition
    const rendered = buildRenderSlide(reopened.deck.slides[0]!, reopened.deck.size, {
      fitWidthPx: FIT,
    })
    const entries = inkNodesOf(rendered)
    expect(entries).toHaveLength(1)
    const node = entries[0]!.node as PictureRenderNode
    expect(node.name).toContain(INK_NAME_PREFIX)

    // Page-coordinate hit test: a point on the stroke hits, a far-away point doesn't
    const inside = { x: node.box.x + node.box.w / 2, y: node.box.y + node.box.h / 2 }
    expect(inkNodeHitTest(entries[0]!, inside, 6)).toBe(true)
    expect(inkNodeHitTest(entries[0]!, { x: node.box.x - 100, y: node.box.y - 100 }, 6)).toBe(false)
  })

  it('a regular image (no ink name/descr) is not recognized as an ink stroke', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addPicture(opened, slide, {
      bytes: new Uint8Array(PNG_1PX),
      ext: 'png',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
    })
    const rendered = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: FIT })
    expect(inkNodesOf(rendered)).toHaveLength(0)
  })

  it('rejects hostile descr payloads instead of decoding them', () => {
    const hostile = (over: object) =>
      JSON.stringify({ v: 1, tool: 'pen', color: 'C00000', width: 2, points: [[0, 0]], ...over })
    expect(decodeInkPayload(hostile({ points: [[0, NaN]] }))).toBeNull()
    expect(decodeInkPayload(hostile({ points: [[Infinity, 0]] }))).toBeNull()
    expect(decodeInkPayload(hostile({ points: 'oops' }))).toBeNull()
    expect(
      decodeInkPayload(hostile({ points: Array.from({ length: 20001 }, () => [0, 0]) })),
    ).toBeNull()
    expect(decodeInkPayload(hostile({ width: Infinity }))?.width).toBe(2)
    expect(decodeInkPayload(hostile({ width: 1e9 }))?.width).toBe(200)
    expect(decodeInkPayload(hostile({ color: 'red;evil' }))?.color).toBe('000000')
    // a valid stroke still round-trips
    expect(decodeInkPayload(encodeInkPayload(stroke))).toMatchObject({ color: 'C00000', width: 2 })
  })
})

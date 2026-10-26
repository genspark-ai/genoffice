/**
 * Deck-picture vision injection: when the active endpoint can take images, the
 * send path collects the deck's pictures as multimodal inputs and the outline
 * annotates each one "— attached as image #K"; when it cannot, nothing is
 * collected and the outline stays byte-identical. Budgets truncate with a
 * skipped count; non-whitelisted mimes and failed reads are skipped silently.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
  PlacedBox,
} from '@genoffice/pptx-render'
import {
  collectDeckImages,
  MAX_DECK_IMAGES,
  type DeckImageReader,
} from '../src/renderer/ai/deck-images'
import {
  createSlidesSkill,
  setDeckImageAttachment,
  type DeckAccess,
} from '../src/renderer/ai/slides-skill'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

const png1px =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const gif1px = 'R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=='

const picture = (
  id: string,
  b: PlacedBox,
  extra: Partial<PictureRenderNode> = {},
): PictureRenderNode =>
  ({
    id: `r_${id}`,
    sourceId: id,
    type: 'picture',
    box: b,
    ...extra,
  }) as PictureRenderNode

const shape = (id: string, b: PlacedBox): ShapeRenderNode => ({
  id,
  sourceId: id,
  type: 'shape',
  box: b,
  fill: { kind: 'none' },
})

const slideOf = (nodes: RenderNode[]): RenderSlide => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes,
})

const access = (slides: RenderSlide[]): DeckAccess => ({
  getSlides: () => slides,
  getCurrent: () => 0,
  getSelectedIds: () => [],
  applySlide: () => {},
  applyDeck: () => {},
  fitWidthPx: 1280,
})

/** Bridge stub: answers dataUrl-less pictures by elementId, never touches the network */
const readerOf = (table: Record<string, { base64: string; mime: string }>): DeckImageReader => {
  const spy = vi.fn(async (op: { slideIndex: number; elementId: string }) => {
    const hit = table[op.elementId]
    return hit ? { ok: true, base64: hit.base64, mime: hit.mime } : { ok: false, error: 'missing' }
  })
  return spy as unknown as DeckImageReader
}

beforeEach(() => {
  setDeckImageAttachment(null)
})

describe('collectDeckImages', () => {
  it('collects picture nodes with dataUrl present (no bridge call) and annotates ordinals', async () => {
    const slides = [
      slideOf([
        picture('pic1', box(0, 0, 100, 100), { dataUrl: `data:image/png;base64,${png1px}` }),
        shape('t1', box(0, 120, 300, 60)),
        picture('pic2', box(0, 200, 100, 100), { dataUrl: `data:image/gif;base64,${gif1px}` }),
      ]),
      slideOf([picture('pic3', box(0, 0, 80, 80), { dataUrl: `data:image/png;base64,${png1px}` })]),
    ]
    const reader = vi.fn()
    const plan = await collectDeckImages(slides, reader as unknown as DeckImageReader)
    expect(reader).not.toHaveBeenCalled()
    expect(plan.images).toEqual([
      { base64: png1px, mime: 'image/png' },
      { base64: gif1px, mime: 'image/gif' },
    ])
    // pic3 duplicates pic1's content: not re-attached, but annotated with its ordinal
    expect(plan.ordinalsBySlide.get(0)).toEqual([
      { elementId: 'pic1', ordinal: 1 },
      { elementId: 'pic2', ordinal: 2 },
    ])
    expect(plan.ordinalsBySlide.get(1)).toEqual([{ elementId: 'pic3', ordinal: 1 }])
    expect(plan.skipped).toBe(0)
  })

  it('resolves dataUrl-less pictures through the bridge with the durable id preferred', async () => {
    const slides = [
      slideOf([
        picture('src-pic', box(0, 0, 100, 100), { durableId: 'e_pic' }),
      ]),
    ]
    const seen: Array<{ slideIndex: number; elementId: string }> = []
    const reader: DeckImageReader = async (op) => {
      seen.push(op)
      return { ok: true, base64: png1px, mime: 'image/png' }
    }
    const plan = await collectDeckImages(slides, reader)
    expect(seen).toEqual([{ slideIndex: 0, elementId: 'e_pic' }])
    expect(plan.images).toEqual([{ base64: png1px, mime: 'image/png' }])
    expect(plan.ordinalsBySlide.get(0)).toEqual([{ elementId: 'e_pic', ordinal: 1 }])
  })

  it('offsets ordinals by the attachment images already queued ahead', async () => {
    const slides = [
      slideOf([picture('pic', box(0, 0, 10, 10), { dataUrl: `data:image/png;base64,${png1px}` })]),
    ]
    const plan = await collectDeckImages(slides, readerOf({}), 3)
    expect(plan.ordinalsBySlide.get(0)).toEqual([{ elementId: 'pic', ordinal: 4 }])
  })

  it('skips non-whitelisted mimes (bmp/tiff/svg) and failed reads', async () => {
    const slides = [
      slideOf([
        picture('ok', box(0, 0, 10, 10), { dataUrl: 'data:image/png;base64,QUJD' }),
        picture('bmp', box(0, 20, 10, 10), { dataUrl: 'data:image/bmp;base64,QUJD' }),
        picture('svg', box(0, 40, 10, 10), { dataUrl: 'data:image/svg+xml;base64,QUJD' }),
        picture('tiff', box(0, 60, 10, 10), { dataUrl: 'data:image/tiff;base64,QUJD' }),
        picture('gone', box(0, 80, 10, 10)),
      ]),
    ]
    const plan = await collectDeckImages(slides, readerOf({ gone: { base64: 'QQ', mime: 'image/png' } }))
    // 'gone' resolves via the bridge and is attached; mime failures are skipped
    expect(plan.images.map((i) => i.base64)).toEqual(['QUJD', 'QQ'])
    expect(plan.skipped).toBe(3)
  })

  it('truncates at the 8-picture cap and reports the skipped count', async () => {
    const pics = Array.from({ length: 10 }, (_, i) =>
      picture(`pic${i}`, box(i * 20, 0, 10, 10), {
        dataUrl: `data:image/png;base64,${png1px.slice(0, -4)}${i.toString().padStart(4, 'A')}`,
      }),
    )
    const plan = await collectDeckImages([slideOf(pics)], readerOf({}))
    expect(plan.images.length).toBe(MAX_DECK_IMAGES)
    expect(plan.ordinalsBySlide.get(0)!.length).toBe(MAX_DECK_IMAGES)
    expect(plan.skipped).toBe(2)
  })

  it('truncates when the combined base64 budget would be exceeded', async () => {
    // ~2MB payloads: 8 pictures stay under both the per-image cap and the
    // 8-count cap, but a 9th would cross the 20MB total — so it is skipped
    const big = 'A'.repeat(2 * 1024 * 1024)
    const pics = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i'].map((n, i) =>
      picture(`pic-${n}`, box(i * 20, 0, 10, 10), { dataUrl: `data:image/png;base64,${big}${n}` }),
    )
    const plan = await collectDeckImages([slideOf(pics)], readerOf({}))
    expect(plan.images.length).toBe(8)
    expect(plan.skipped).toBe(1)
  })

  it('rejects a single picture over the 5MB per-image cap', async () => {
    const huge = 'A'.repeat(6 * 1024 * 1024)
    const slides = [
      slideOf([picture('big', box(0, 0, 10, 10), { dataUrl: `data:image/png;base64,${huge}` })]),
    ]
    const plan = await collectDeckImages(slides, readerOf({}))
    expect(plan.images).toEqual([])
    expect(plan.skipped).toBe(1)
  })

  it('never collects pictures withheld with a redact label', async () => {
    const slides = [
      slideOf([
        picture('held', box(0, 0, 10, 10), {
          dataUrl: `data:image/png;base64,${png1px}`,
          redact: 'client photo',
        }),
      ]),
    ]
    const plan = await collectDeckImages(slides, readerOf({}))
    expect(plan.images).toEqual([])
    expect(plan.ordinalsBySlide.size).toBe(0)
    expect(plan.skipped).toBe(0)
  })

  it('walks group children in z-order', async () => {
    const group: GroupRenderNode = {
      id: 'g',
      sourceId: 'g',
      type: 'group',
      box: box(0, 0, 200, 200),
      children: [picture('inner', box(0, 0, 50, 50), { dataUrl: `data:image/gif;base64,${gif1px}` })],
    }
    const slides = [slideOf([group, picture('outer', box(0, 300, 10, 10), { dataUrl: `data:image/png;base64,${png1px}` })])]
    const plan = await collectDeckImages(slides, readerOf({}))
    expect(plan.images.map((i) => i.mime)).toEqual(['image/gif', 'image/png'])
    expect(plan.ordinalsBySlide.get(0)).toEqual([
      { elementId: 'inner', ordinal: 1 },
      { elementId: 'outer', ordinal: 2 },
    ])
  })
})

describe('deck outline picture annotation', () => {
  const picSlide = slideOf([
    picture('pic1', box(0, 0, 100, 100)),
    shape('t1', box(0, 120, 300, 60)),
    picture('pic2', box(0, 200, 100, 100)),
  ])

  it('appends "— attached as image #K" to picture lines when a plan is set', () => {
    setDeckImageAttachment({
      ordinals: new Map([[0, [{ elementId: 'pic1', ordinal: 2 }, { elementId: 'pic2', ordinal: 3 }]]]),
      skipped: 0,
    })
    const outline = createSlidesSkill(access([picSlide])).buildContext!()
    expect(outline).toContain('- pic1 | picture — attached as image #2')
    expect(outline).toContain('- pic2 | picture — attached as image #3')
    expect(outline).not.toContain('- t1 | shape —')
  })

  it('stays byte-identical when no plan is set (vision off / nothing collected)', () => {
    setDeckImageAttachment(null)
    const withSkill = createSlidesSkill(access([picSlide])).buildContext!()
    expect(withSkill).not.toContain('attached as image')
    expect(withSkill).not.toContain('pictures not attached')
    expect(withSkill).toContain('- pic1 | picture')
    expect(withSkill).toContain('- t1 | shape')
    expect(withSkill).toContain('- pic2 | picture')
  })

  it('a plan with an empty map changes nothing', () => {
    setDeckImageAttachment({ ordinals: new Map(), skipped: 0 })
    const outline = createSlidesSkill(access([picSlide])).buildContext!()
    expect(outline).not.toContain('attached as image')
    expect(outline).not.toContain('pictures not attached')
  })

  it('reports skipped pictures after the per-slide lines when the budget truncated', () => {
    setDeckImageAttachment({
      ordinals: new Map([[0, [{ elementId: 'pic1', ordinal: 1 }]]]),
      skipped: 5,
    })
    const outline = createSlidesSkill(access([picSlide])).buildContext!()
    expect(outline).toContain('(5 more pictures not attached)')
    expect(outline.indexOf('(5 more pictures not attached)')).toBeLessThan(
      outline.indexOf('(Use read_slide'),
    )
  })
})

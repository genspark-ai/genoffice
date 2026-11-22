/**
 * Executor rejection reasons: unknown op, target not found, out of range.
 * The executor validates the whole batch first (plan), then applies atomically
 * by default; every failure is a guided error that states what failed and how
 * to fix it, with the op usage appended for known ops.
 */
import { describe, it, expect, beforeAll } from 'vitest'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let textId: string

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'Hello' }] }],
  })
  textId = el.id
})

const textOp = (target: unknown) => ({
  op: 'setText',
  target,
  paragraphs: [{ runs: [{ text: 'Updated' }] }],
})

describe('executor rejections', () => {
  it('rejects non-finite and negative rect dimensions for insert operations', () => {
    const invalidOffsets = [
      { x: 0, y: 0, cx: NaN, cy: 100 },
      { x: 0, y: 0, cx: Infinity, cy: 100 },
      { x: 0, y: 0, cx: -1, cy: 100 },
      { x: 0, y: 0, cx: 100, cy: -1 },
    ]

    for (const offset of invalidOffsets) {
      const r = runTxn(opened, {
        ops: [{ op: 'addElement', target: { slide: 0 }, offset, kind: 'textbox' }],
      })
      expect(r.applied).toBe(false)
      expect(r.failures).toHaveLength(1)
      expect(r.failures![0]!.error).toMatch(/offset\.(cx|cy)/)
      expect(r.failures![0]!.error).toMatch(/finite number|>= 0/)
    }
  })

  it('accepts finite, non-negative rect dimensions for insert operations', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { slide: 0 },
          offset: { x: 0, y: 0, cx: 100, cy: 100 },
          kind: 'textbox',
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(1)
  })

  it('rejects an unknown op with the supported-ops vocabulary', () => {
    const r = runTxn(opened, { ops: [{ op: 'sparkle' }] })
    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.index).toBe(0)
    expect(r.failures![0]!.error).toContain('unknown op "sparkle"')
    expect(r.failures![0]!.error).toContain('Supported ops:')
    expect(r.failures![0]!.error).not.toContain('Usage:')
  })

  it('rejects a missing element with the available ids', () => {
    const r = runTxn(opened, {
      ops: [textOp({ slide: 0, el: 'e_DOES_NOT_EXIST' })],
    })
    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.error).toContain('no element "e_DOES_NOT_EXIST" on slide 0')
    expect(r.failures![0]!.error).toContain('Available:')
    expect(r.failures![0]!.error).toContain('Nothing was applied (atomic)')
    expect(r.failures![0]!.error).toContain('Usage: setText')
  })

  it('rejects a missing slide id with the available slide ids', () => {
    const r = runTxn(opened, {
      ops: [textOp({ slide: 's_999', el: textId })],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no slide "s_999"')
    expect(r.failures![0]!.error).toContain('Available:')
  })

  it('rejects an out-of-range slide index with the valid range', () => {
    const r = runTxn(opened, {
      ops: [textOp({ slide: 99, el: textId })],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('slide index 99 is out of range (0-0)')
    expect(r.failures![0]!.error).toContain('Nothing was applied (atomic)')
  })

  it('rejects an empty batch without touching the deck', () => {
    const before = opened.deck.slides.length
    const r = runTxn(opened, { ops: [] })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('ops must be a non-empty array')
    expect(opened.deck.slides.length).toBe(before)
  })

  it('atomic batches apply nothing when any op fails', async () => {
    const fresh = await openPptx(await createBlankPptx())
    const slide = fresh.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Before' }] }],
    })
    const r = runTxn(fresh, {
      ops: [textOp({ slide: 0, el: el.id }), { op: 'sparkle' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)
    expect(fresh.deck.slides[0]!.elements.find((e) => e.id === el.id)).toBeTruthy()
  })

  it('per_op batches apply successes and collect failures by index', async () => {
    const fresh = await openPptx(await createBlankPptx())
    const slide = fresh.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Before' }] }],
    })
    const r = runTxn(fresh, {
      isolation: 'per_op',
      ops: [textOp({ slide: 0, el: el.id }), { op: 'sparkle' }],
    })
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(1)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.index).toBe(1)
  })

  it('dry runs validate without mutating', async () => {
    const fresh = await openPptx(await createBlankPptx())
    const slide = fresh.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Before' }] }],
    })
    const r = runTxn(fresh, {
      dryRun: true,
      ops: [textOp({ slide: 0, el: el.id })],
    })
    expect(r.applied).toBe(false)
    expect(r.dryRun).toBe(true)
    expect(r.plan).toHaveLength(1)
    expect(r.failures ?? []).toEqual([])
  })

  it('rejects non-integer, out-of-range and non-finite addTable dims', () => {
    const offset = { x: 0, y: 0, cx: 1828800, cy: 914400 }
    const bad: Array<[unknown, unknown]> = [
      [NaN, 2],
      [2, Infinity],
      [0, 2],
      [2, 76],
      [1.5, 2],
      ['2', 2],
    ]
    for (const [rows, cols] of bad) {
      const r = runTxn(opened, {
        ops: [{ op: 'addTable', target: { slide: 0 }, offset, rows, cols }],
      })
      expect(r.applied).toBe(false)
      expect(r.failures).toHaveLength(1)
      expect(r.failures![0]!.error).toMatch(/"rows" and "cols" \(1\.\.75\)/)
    }
  })

  it('accepts addTable dims at the 75x75 limit', async () => {
    const fresh = await openPptx(await createBlankPptx())
    const r = runTxn(fresh, {
      dryRun: true,
      ops: [
        {
          op: 'addTable',
          target: { slide: 0 },
          offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
          rows: 75,
          cols: 75,
        },
      ],
    })
    expect(r.failures ?? []).toEqual([])
  })

  it('element-level ops in per_op mode only snapshot the targeted slide', async () => {
    const fresh = await openPptx(await createBlankPptx())
    for (let i = 0; i < 5; i++) duplicateSlide(fresh, 0)
    expect(fresh.deck.slides).toHaveLength(6)

    const elSlide2 = addElement(fresh.deck.slides[2]!, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Slide 2 initial' }] }],
    })

    const slide0Before = fresh.deck.slides[0]!
    const slide4Before = fresh.deck.slides[4]!

    const r = runTxn(fresh, {
      isolation: 'per_op',
      ops: [
        {
          op: 'setText',
          target: { slide: 2, el: elSlide2.id },
          paragraphs: [
            {
              runs: [{ text: 'Rejected' }],
              bullet: 'blip',
              bulletImage: { base64: 'AA==', ext: 'not-real' },
            },
          ],
        },
      ],
    })

    expect(r.applied).toBe(false)
    expect(r.failures).toHaveLength(1)

    // Slide 0 and slide 4 were NOT cloned or replaced — they are the exact same object references!
    expect(fresh.deck.slides[0]).toBe(slide0Before)
    expect(fresh.deck.slides[4]).toBe(slide4Before)

    // Slide 2 was restored to pre-op state
    const restoredEl = fresh.deck.slides[2]!.elements.find((e) => e.id === elSlide2.id)
    expect(restoredEl?.text?.paragraphs[0]?.runs[0]?.text).toBe('Slide 2 initial')
  })

  it('per_op batch of 40 setText ops does not pay ops × deck size snapshot overhead', async () => {
    const fresh = await openPptx(await createBlankPptx())
    for (let i = 0; i < 59; i++) duplicateSlide(fresh, 0)
    expect(fresh.deck.slides).toHaveLength(60)

    const textEls: Array<{ slide: number; el: string }> = []
    for (let i = 0; i < 40; i++) {
      const slide = fresh.deck.slides[i]!
      const el = addElement(slide, {
        kind: 'textbox',
        offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
        paragraphs: [{ runs: [{ text: `Init ${i}` }] }],
      })
      textEls.push({ slide: i, el: el.id })
    }

    const start = performance.now()
    const r = runTxn(fresh, {
      isolation: 'per_op',
      ops: textEls.map(({ slide, el }, idx) => ({
        op: 'setText',
        target: { slide, el },
        paragraphs: [{ runs: [{ text: `Updated ${idx}` }] }],
      })),
    })
    const elapsed = performance.now() - start
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(40)
    expect(elapsed).toBeLessThan(2000)
  })
})

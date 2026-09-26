/**
 * Executor rejection reasons: unknown op, target not found, out of range.
 * The executor validates the whole batch first (plan), then applies atomically
 * by default; every failure is a guided error that states what failed and how
 * to fix it, with the op usage appended for known ops.
 */
import { describe, it, expect, beforeAll } from 'vitest'
import { addElement, createBlankPptx, openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
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
})

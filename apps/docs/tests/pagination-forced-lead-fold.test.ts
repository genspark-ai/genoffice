import { describe, expect, it } from 'vitest'
import { foldForcedLeadSpaceBefore } from '../src/renderer/pagination-measure'
import type { BlockBox } from '../src/renderer/pagination-types'

const box = (top: number, height: number, extra: Partial<BlockBox> = {}): BlockBox => ({
  top,
  height,
  ...extra,
})

describe('foldForcedLeadSpaceBefore', () => {
  it('moves the kept part of the collapsed margin from the previous space after to the lead', () => {
    // after=200 -> before=400 collapse to 400/20pt = 26.67px; the lead keeps 10pt = 13.33px
    const prev = box(0, 100 + 26.67, { spaceAfterPx: 26.67 })
    const lead = box(126.67, 50, {
      breakBefore: true,
      marginBeforePx: 26.67,
      forcedLeadSpacePx: 13.33,
    })
    foldForcedLeadSpaceBefore([prev, lead])
    expect(prev.spaceAfterPx).toBeCloseTo(13.34, 2)
    expect(prev.height).toBeCloseTo(113.34, 2)
    expect(lead.top).toBeCloseTo(113.34, 2)
    expect(lead.height).toBeCloseTo(63.33, 2)
    expect(lead.spaceBeforePx).toBeCloseTo(13.33, 2)
    expect(lead.leadFoldPx).toBeCloseTo(13.33, 2)
    expect(lead.leadMtPx).toBeCloseTo(13.33, 2)
  })

  it('leaves blocks after a break-only paragraph, without a plan amount, or with padded spacing alone', () => {
    const cases: BlockBox[][] = [
      [
        box(0, 120, { spaceAfterPx: 20, breakAfter: true }),
        box(120, 50, { breakBefore: true, marginBeforePx: 20, forcedLeadSpacePx: 20 }),
      ],
      [box(0, 120, { spaceAfterPx: 20 }), box(120, 50, { marginBeforePx: 20 })],
      [
        box(0, 120, { spaceAfterPx: 20 }),
        box(120, 50, { breakBefore: true, padBeforePx: 20, forcedLeadSpacePx: 20 }),
      ],
    ]
    for (const blocks of cases) {
      foldForcedLeadSpaceBefore(blocks)
      expect(blocks[1].top).toBe(120)
      expect(blocks[1].leadMtPx).toBeUndefined()
      expect(blocks[0].spaceAfterPx).toBe(20)
    }
  })
})

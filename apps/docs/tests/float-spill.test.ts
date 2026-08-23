import { describe, expect, it } from 'vitest'
import {
  appendFloatSpillBlock,
  computeSectionedSlicesF2,
  type BlockBox,
  type FloatBox,
} from '../src/renderer/pagination'

const el = null as unknown as HTMLElement

describe('appendFloatSpillBlock', () => {
  it('appends a virtual block spanning to the lowest float bottom', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100 }]
    const floats: FloatBox[] = [{ el, top: 950, height: 400 }]
    const total = appendFloatSpillBlock(blocks, 100, floats)
    expect(total).toBe(1350)
    const spill = blocks[1]
    expect(spill.isFloatSpill).toBe(true)
    expect(spill.top).toBe(100)
    expect(spill.height).toBe(1250)
    expect(spill.lineBoxes!.length).toBeGreaterThan(1)
  })

  it('returns null when floats stay within the flow', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 500 }]
    expect(appendFloatSpillBlock(blocks, 500, [{ el, top: 100, height: 200 }])).toBeNull()
    expect(blocks).toHaveLength(1)
  })

  it('materializes trailing pages for overflowing floats', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100 }]
    const floats: FloatBox[] = [{ el, top: 1500, height: 400 }]
    const total = appendFloatSpillBlock(blocks, 100, floats)!
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 900, forceBreak: false }],
      total,
    )
    expect(slices.length).toBeGreaterThanOrEqual(2)
    expect(slices[slices.length - 1].end).toBeGreaterThanOrEqual(1900)
  })
})

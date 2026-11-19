import { describe, expect, it } from 'vitest'
import {
  clickSelection,
  currentAfterHistory,
  movedBlockPositions,
  normalizeSelection,
  planSlideDuplicates,
  planSlideMoves,
  rangeSelection,
  toggleSelection,
} from '../src/shared/slide-selection'

const sel = (selected: number[], current: number) => ({ selected, current })

describe('thumbnail click selection', () => {
  it('a plain click selects the clicked slide alone and moves the anchor', () => {
    expect(clickSelection(sel([0, 1, 2], 1), 4, {})).toEqual(sel([4], 4))
  })

  it('Shift-click selects the range from the anchor, either direction, keeping the anchor', () => {
    expect(clickSelection(sel([2], 2), 5, { shift: true })).toEqual(sel([2, 3, 4, 5], 2))
    expect(clickSelection(sel([2, 3, 4, 5], 2), 0, { shift: true })).toEqual(sel([0, 1, 2], 2))
    expect(rangeSelection(3, 3)).toEqual([3])
  })

  it('Ctrl-click toggles one slide without touching the others', () => {
    expect(clickSelection(sel([1, 3], 1), 5, { toggle: true })).toEqual(sel([1, 3, 5], 1))
    expect(clickSelection(sel([1, 3, 5], 1), 3, { toggle: true })).toEqual(sel([1, 5], 1))
  })

  it('Ctrl-click on the anchor hands the anchor to the nearest remaining slide', () => {
    expect(toggleSelection(sel([1, 4, 5], 4), 4)).toEqual(sel([1, 5], 5))
    expect(toggleSelection(sel([0, 4, 5], 5), 5)).toEqual(sel([0, 4], 4))
  })

  it('never empties the selection', () => {
    expect(toggleSelection(sel([2], 2), 2)).toEqual(sel([2], 2))
  })

  it('normalize keeps a valid selection by reference and collapses a stale one', () => {
    const valid = [1, 2]
    expect(normalizeSelection(valid, 2, 3)).toBe(valid)
    expect(normalizeSelection([1, 2], 0, 3)).toEqual([0])
    expect(normalizeSelection([1, 5], 1, 3)).toEqual([1])
    expect(normalizeSelection([], 1, 3)).toEqual([1])
  })
})

describe('history selection', () => {
  const before = ['A', 'B', 'C', 'D'].map((partPath) => ({ partPath }))
  const moved = [before[0]!, before[2]!, before[3]!, before[1]!]

  it('keeps the same slide selected when undo restores its earlier position', () => {
    expect(currentAfterHistory(before, 3, moved[3]!.partPath)).toBe(1)
  })

  it('keeps the same slide selected when redo reapplies the move', () => {
    expect(currentAfterHistory(moved, 1, before[1]!.partPath)).toBe(3)
  })

  it('clamps the old index when the selected slide no longer exists', () => {
    expect(currentAfterHistory(before.slice(0, 3), 3, 'D')).toBe(2)
    expect(currentAfterHistory(before, 2)).toBe(2)
  })
})

/** Applies moveSlide steps the way the engine does: remove, then insert at `to`. */
function applyMoves(count: number, steps: Array<{ from: number; to: number }>): number[] {
  const order = Array.from({ length: count }, (_, i) => i)
  for (const { from, to } of steps) {
    expect(from).not.toBe(to)
    const [id] = order.splice(from, 1)
    order.splice(to, 0, id!)
  }
  return order
}

describe('drag-reorder plan', () => {
  it('moves a scattered selection as one block to the drop gap', () => {
    const steps = planSlideMoves(6, [1, 3], 5)
    expect(applyMoves(6, steps)).toEqual([0, 2, 4, 1, 3, 5])
    expect(movedBlockPositions([1, 3], 5)).toEqual([3, 4])
  })

  it('handles drops above the block and at both ends', () => {
    expect(applyMoves(5, planSlideMoves(5, [3, 4], 1))).toEqual([0, 3, 4, 1, 2])
    expect(applyMoves(5, planSlideMoves(5, [1, 2], 0))).toEqual([1, 2, 0, 3, 4])
    expect(applyMoves(5, planSlideMoves(5, [0, 2], 5))).toEqual([1, 3, 4, 0, 2])
    expect(movedBlockPositions([0, 2], 5)).toEqual([3, 4])
  })

  it('is empty when the order would not change', () => {
    expect(planSlideMoves(5, [1, 2], 1)).toEqual([])
    expect(planSlideMoves(5, [1, 2], 3)).toEqual([])
    expect(planSlideMoves(5, [2], 2)).toEqual([])
  })
})

/** Applies duplicate/move steps on a deck of labels; a copy of slide n is labelled `${n}'`. */
function applyDuplicates(
  count: number,
  steps: ReturnType<typeof planSlideDuplicates>,
): { order: string[]; validTargets: boolean } {
  const order = Array.from({ length: count }, (_, i) => String(i))
  let validTargets = true
  for (const st of steps) {
    if ('duplicate' in st) {
      if (st.duplicate >= count) validTargets = false
      order.splice(st.duplicate + 1, 0, `${order[st.duplicate]}'`)
    } else {
      if (st.from >= count) validTargets = false
      const [id] = order.splice(st.from, 1)
      order.splice(st.to, 0, id!)
    }
  }
  return { order, validTargets }
}

describe('duplicate plan', () => {
  it('lands the copies in order after the last selected slide', () => {
    expect(applyDuplicates(4, planSlideDuplicates([1, 2])).order).toEqual([
      '0',
      '1',
      '2',
      "1'",
      "2'",
      '3',
    ])
    expect(applyDuplicates(4, planSlideDuplicates([0, 3])).order).toEqual([
      '0',
      '1',
      '2',
      '3',
      "0'",
      "3'",
    ])
  })

  it('only targets slides that exist before the transaction, even when the last slide is selected', () => {
    const r = applyDuplicates(5, planSlideDuplicates([0, 2, 4]))
    expect(r.validTargets).toBe(true)
    expect(r.order).toEqual(['0', '1', '2', '3', '4', "0'", "2'", "4'"])
  })

  it('a single slide duplicates in place', () => {
    expect(planSlideDuplicates([2])).toEqual([{ duplicate: 2 }])
    expect(planSlideDuplicates([])).toEqual([])
  })
})

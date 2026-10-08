import { describe, expect, it } from 'vitest'

import {
  groupLevelOps,
  groupTogglePlan,
  levelButtonPlan,
  moveOutlineEntries,
  outlineGroups,
  outlineGutterThickness,
  outlineMaxLevel,
  shiftOutlineEntries,
  type OutlineEntry,
} from '../src/renderer/outline-model'

function entries(levels: Record<number, number>): Map<number, OutlineEntry> {
  const map = new Map<number, OutlineEntry>()
  for (const [index, level] of Object.entries(levels)) {
    map.set(Number(index), { level, collapsed: false })
  }
  return map
}

const never = (): boolean => false

describe('outlineGroups', () => {
  it('turns nested levels into one group per level run, summary below by default', () => {
    // Rows 1-6 at level 1, rows 2-3 and 5-6 at level 2 (0-based).
    const levels = entries({ 1: 1, 2: 2, 3: 2, 4: 1, 5: 2, 6: 2 })
    expect(outlineMaxLevel(levels)).toBe(2)
    expect(outlineGroups(levels, true, never)).toEqual([
      { level: 1, start: 1, end: 6, summary: 7, collapsed: false, folded: false },
      { level: 2, start: 2, end: 3, summary: 4, collapsed: false, folded: false },
      { level: 2, start: 5, end: 6, summary: 7, collapsed: false, folded: false },
    ])
  })

  it('places the summary above when summaryBelow is off and flags fully hidden groups', () => {
    const levels = entries({ 0: 1, 1: 1, 3: 1 })
    const groups = outlineGroups(levels, false, (index) => index === 3)
    expect(groups).toEqual([
      { level: 1, start: 0, end: 1, summary: -1, collapsed: false, folded: false },
      { level: 1, start: 3, end: 3, summary: 2, collapsed: true, folded: false },
    ])
  })

  it('sizes the gutter by level buttons (1..max+1) and hides it without an outline', () => {
    expect(outlineGutterThickness(0)).toBe(0)
    expect(outlineGutterThickness(1)).toBeGreaterThan(outlineGutterThickness(0))
    expect(outlineGutterThickness(3)).toBeGreaterThan(outlineGutterThickness(1))
  })
})

describe('level buttons', () => {
  const levels = entries({ 1: 1, 2: 2, 3: 2, 4: 1, 5: 2, 6: 2 })
  const groups = outlineGroups(levels, true, never)

  it('level 1 folds everything and marks every summary collapsed', () => {
    const plan = levelButtonPlan(levels, groups, 1)
    expect(plan.hide).toEqual([{ start: 1, end: 6 }])
    expect(plan.show).toEqual([])
    expect(plan.summaries).toEqual(
      expect.arrayContaining([
        { index: 7, collapsed: true },
        { index: 4, collapsed: true },
      ]),
    )
  })

  it('level 2 keeps level-1 detail and folds only level-2 runs', () => {
    const plan = levelButtonPlan(levels, groups, 2)
    expect(plan.hide).toEqual([
      { start: 2, end: 3 },
      { start: 5, end: 6 },
    ])
    expect(plan.show).toEqual([
      { start: 1, end: 1 },
      { start: 4, end: 4 },
    ])
    // Row 7 is summary of both the level-1 group (open) and a level-2 group
    // (folded); the shallower group wins so the flag reads "expanded".
    expect(plan.summaries).toEqual(
      expect.arrayContaining([
        { index: 7, collapsed: false },
        { index: 4, collapsed: true },
      ]),
    )
  })

  it('the deepest button shows all detail', () => {
    const plan = levelButtonPlan(levels, groups, 3)
    expect(plan.hide).toEqual([])
    expect(plan.show).toEqual([{ start: 1, end: 6 }])
    expect(plan.summaries.every((summary) => !summary.collapsed)).toBe(true)
  })
})

describe('groupTogglePlan', () => {
  it('[-] hides the span and [+] restores it while keeping folded nested groups hidden', () => {
    const levels = entries({ 1: 1, 2: 2, 3: 2, 4: 1 })
    levels.set(4, { level: 1, collapsed: true })
    const groups = outlineGroups(levels, true, (index) => index === 2 || index === 3)
    const outer = groups.find((group) => group.level === 1)!
    const inner = groups.find((group) => group.level === 2)!
    expect(inner.collapsed).toBe(true)
    expect(inner.folded).toBe(true)
    expect(groupTogglePlan(groups, outer, true)).toEqual({
      hide: [{ start: 1, end: 4 }],
      show: [],
      summaries: [{ index: 5, collapsed: true }],
    })
    expect(groupTogglePlan(groups, outer, false)).toEqual({
      hide: [{ start: 2, end: 3 }],
      show: [{ start: 1, end: 4 }],
      summaries: [{ index: 5, collapsed: false }],
    })
  })
})

describe('groupTogglePlan nested state', () => {
  it('[+] on the outer group re-shows inner groups hidden only by the outer fold', () => {
    const levels = entries({ 1: 1, 2: 2, 3: 2, 4: 1 })
    const groups = outlineGroups(levels, true, () => true)
    const outer = groups.find((group) => group.level === 1)!
    expect(groupTogglePlan(groups, outer, false).hide).toEqual([])
  })

  it('an inner group sharing the outer summary line unfolds with it', () => {
    // Rows 1-4 level 1, rows 3-4 level 2: both groups summarize on row 5.
    const levels = entries({ 1: 1, 2: 1, 3: 2, 4: 2 })
    levels.set(5, { level: 0, collapsed: true })
    const groups = outlineGroups(levels, true, () => true)
    const outer = groups.find((group) => group.level === 1)!
    expect(groups.find((group) => group.level === 2)!.folded).toBe(true)
    expect(groupTogglePlan(groups, outer, false).hide).toEqual([])
  })
})

describe('shiftOutlineEntries', () => {
  it('moves levels down on insert and drops removed lines', () => {
    const levels = entries({ 2: 1, 4: 1, 6: 2 })
    expect(shiftOutlineEntries(levels, 3, 2)).toBe(0)
    expect([...levels.keys()].sort((a, b) => a - b)).toEqual([2, 6, 8])
    expect(shiftOutlineEntries(levels, 5, -2)).toBe(0)
    expect([...levels.keys()].sort((a, b) => a - b)).toEqual([2, 6])
  })

  it('lines inserted inside a group inherit the level above so the run stays whole', () => {
    const levels = entries({ 1: 1, 2: 2, 3: 1 })
    expect(shiftOutlineEntries(levels, 2, 2)).toBe(1)
    expect(
      [...levels.entries()].sort((a, b) => a[0] - b[0]).map(([at, e]) => [at, e.level]),
    ).toEqual([
      [1, 1],
      [2, 1],
      [3, 1],
      [4, 2],
      [5, 1],
    ])
  })
})

describe('moveOutlineEntries', () => {
  it('swaps the moved block with the rows it jumps over, in both directions', () => {
    const keys = (map: Map<number, OutlineEntry>): number[] => [...map.keys()].sort((a, b) => a - b)
    const down = entries({ 1: 1, 2: 1, 5: 2 })
    moveOutlineEntries(down, 1, 2, 6)
    expect(keys(down)).toEqual([3, 4, 5])
    const up = entries({ 1: 1, 5: 2, 6: 2 })
    moveOutlineEntries(up, 5, 2, 1)
    expect(keys(up)).toEqual([1, 2, 3])
  })
})

describe('groupLevelOps', () => {
  it('shifts each equal-level run and clamps at the 7-level ceiling', () => {
    const levels = entries({ 1: 1, 2: 7, 3: 7, 4: 0 })
    expect(groupLevelOps(levels, 0, 4, 1)).toEqual([
      { start: 0, end: 0, level: 1 },
      { start: 1, end: 1, level: 2 },
      { start: 4, end: 4, level: 1 },
    ])
    expect(groupLevelOps(levels, 0, 4, -1)).toEqual([
      { start: 1, end: 1, level: 0 },
      { start: 2, end: 3, level: 6 },
    ])
  })
})

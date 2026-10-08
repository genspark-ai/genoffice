import { describe, expect, it } from 'vitest'

import {
  applyReorder,
  finalOrder,
  insertIndex,
  moveTargets,
  planReorder,
  tabRange,
  unhideCandidates,
  visibleAfterRemoving,
  type SheetTabEntry,
} from '../src/renderer/sheet-tab-menu'

const sheet = (id: string, extra: Partial<SheetTabEntry> = {}): SheetTabEntry => ({
  id,
  name: id.toUpperCase(),
  hidden: false,
  veryHidden: false,
  ...extra,
})

const SHEETS = [
  sheet('a'),
  sheet('b', { hidden: true }),
  sheet('c', { hidden: true, veryHidden: true }),
  sheet('d'),
  sheet('e', { hidden: true }),
]

describe('unhide list', () => {
  it('lists hidden sheets but never veryHidden ones', () => {
    expect(unhideCandidates(SHEETS).map((s) => s.id)).toEqual(['b', 'e'])
  })

  it('offers every non-veryHidden sheet as a move target', () => {
    expect(moveTargets(SHEETS).map((s) => s.id)).toEqual(['a', 'b', 'd', 'e'])
  })

  it('counts the visible sheets that would remain', () => {
    expect(visibleAfterRemoving(SHEETS, new Set(['a']))).toBe(1)
    expect(visibleAfterRemoving(SHEETS, new Set(['a', 'd']))).toBe(0)
    expect(visibleAfterRemoving(SHEETS, new Set(['b']))).toBe(2)
  })
})

describe('insert index', () => {
  it('inserts before the active sheet, appending when it is unknown', () => {
    expect(insertIndex(['a', 'b', 'c'], 'b')).toBe(1)
    expect(insertIndex(['a', 'b', 'c'], 'zz')).toBe(3)
  })
})

describe('move or copy order math', () => {
  const order = ['a', 'b', 'c', 'd', 'e']

  it('moves a sheet in front of a later sheet', () => {
    expect(finalOrder(order, new Set(['a']), 'd')).toEqual(['b', 'c', 'a', 'd', 'e'])
  })

  it('moves a sheet in front of an earlier sheet', () => {
    expect(finalOrder(order, new Set(['d']), 'b')).toEqual(['a', 'd', 'b', 'c', 'e'])
  })

  it('moves to the end when no anchor is given', () => {
    expect(finalOrder(order, new Set(['b']), null)).toEqual(['a', 'c', 'd', 'e', 'b'])
  })

  it('keeps a group in its relative order and resolves an anchor inside the group', () => {
    expect(finalOrder(order, new Set(['d', 'b']), 'e')).toEqual(['a', 'c', 'b', 'd', 'e'])
    expect(finalOrder(order, new Set(['b', 'c']), 'c')).toEqual(order)
    expect(finalOrder(order, new Set(['a', 'b']), 'a')).toEqual(order)
    expect(finalOrder(order, new Set(['d', 'e']), 'e')).toEqual(order)
  })

  it('plans single-splice steps that reproduce the target through the mutation semantics', () => {
    const cases: [ReadonlySet<string>, string | null][] = [
      [new Set(['a']), 'd'],
      [new Set(['d']), 'b'],
      [new Set(['b']), null],
      [new Set(['d', 'b']), 'e'],
      [new Set(['e', 'a']), 'c'],
      [new Set(['a', 'b', 'c', 'd', 'e']), null],
    ]
    for (const [moving, before] of cases) {
      const target = finalOrder(order, moving, before)
      const steps = planReorder(order, target, moving)
      expect(applyReorder(order, steps)).toEqual(target)
      expect(steps.length).toBeLessThanOrEqual(moving.size)
    }
  })

  it('reproduces every subset/anchor combination of a six-sheet book', () => {
    const six = ['a', 'b', 'c', 'd', 'e', 'f']
    for (let mask = 1; mask < 1 << six.length; mask += 1) {
      const moving = new Set(six.filter((_, i) => mask & (1 << i)))
      for (const before of [...six, null]) {
        const target = finalOrder(six, moving, before)
        const steps = planReorder(six, target, moving)
        expect(applyReorder(six, steps)).toEqual(target)
        expect(steps.length).toBeLessThanOrEqual(moving.size)
      }
    }
  })

  it('plans nothing when the order already matches', () => {
    expect(planReorder(order, order, new Set(['c']))).toEqual([])
  })

  it('uses the post-removal index Univer expects', () => {
    expect(planReorder(order, finalOrder(order, new Set(['a']), 'd'), new Set(['a']))).toEqual([
      { id: 'a', to: 2 },
    ])
  })
})

describe('shift-click tab range', () => {
  it('selects the visible tabs between two tabs in either direction', () => {
    expect(tabRange(SHEETS, 'a', 'd')).toEqual(['a', 'd'])
    expect(tabRange(SHEETS, 'd', 'a')).toEqual(['a', 'd'])
    expect(tabRange(SHEETS, 'zz', 'd')).toEqual(['d'])
  })
})

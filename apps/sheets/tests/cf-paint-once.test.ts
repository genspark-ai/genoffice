import { describe, expect, it } from 'vitest'

import { dropShadowedPaintOnceRules } from '../src/renderer/univer-sync'

const area = (startRow: number, startColumn: number, endRow: number, endColumn: number) => ({
  startRow,
  startColumn,
  endRow,
  endColumn,
})

const rule = (
  ruleType: string,
  priority: number,
  ranges: ReturnType<typeof area>[],
): Parameters<typeof dropShadowedPaintOnceRules>[0][number] =>
  ({ ruleType, priority, ranges, formulas: [], percent: false, bottom: false, cfvos: [] }) as never

describe('dropShadowedPaintOnceRules', () => {
  it('drops a color scale fully covered by a higher-precedence one', () => {
    // tdf105272: the real scale (B3:D4, priority 12) plus three stacked
    // duplicates on B4:D4 — Excel paints only the top one.
    const rules = [
      rule('colorScale', 12, [area(2, 1, 3, 3)]),
      rule('colorScale', 13, [area(3, 1, 3, 3)]),
      rule('colorScale', 14, [area(3, 1, 3, 3)]),
      rule('dataBar', 11, [area(3, 4, 3, 4)]),
      rule('dataBar', 16, [area(3, 4, 3, 4)]),
    ]
    expect(dropShadowedPaintOnceRules(rules).map((kept) => kept.priority)).toEqual([12, 11])
  })

  it('keeps partial overlaps and non-paint-once rules', () => {
    const rules = [
      rule('colorScale', 1, [area(0, 0, 5, 5)]),
      rule('colorScale', 2, [area(3, 3, 9, 9)]),
      rule('cellIs', 3, [area(0, 0, 5, 5)]),
      rule('cellIs', 4, [area(0, 0, 5, 5)]),
    ]
    expect(dropShadowedPaintOnceRules(rules)).toHaveLength(4)
  })
})

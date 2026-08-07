import { describe, expect, it } from 'vitest'

import {
  monthKeyOf,
  monthKeyParts,
  timelineDomainOf,
  timelineSelection,
} from '../src/domain/pivot-timeline'

const item = (x: number | null, hidden = false) => ({ x, hidden })

describe('monthKeyOf', () => {
  it('parses Excel serials and ISO-ish strings to the same month key', () => {
    // 45292 = 2024-01-01
    expect(monthKeyOf(45292)).toBe(2024 * 12)
    expect(monthKeyOf('2024-01-15')).toBe(2024 * 12)
    expect(monthKeyOf('2024/12/31')).toBe(2024 * 12 + 11)
    expect(monthKeyOf('not a date')).toBeNull()
  })

  it('round-trips through monthKeyParts', () => {
    expect(monthKeyParts(2024 * 12 + 11)).toEqual({ year: 2024, month: 12 })
  })
})

describe('timelineDomainOf', () => {
  it('maps date members to month keys and computes the domain', () => {
    const domain = timelineDomainOf(
      ['2023-11-02', '2024-01-15', '2024-03-01'],
      [item(0), item(1), item(2), item(null)],
    )
    expect(domain).not.toBeNull()
    expect(domain?.members).toEqual([
      { member: 0, monthKey: 2023 * 12 + 10 },
      { member: 1, monthKey: 2024 * 12 },
      { member: 2, monthKey: 2024 * 12 + 2 },
    ])
    expect(domain?.minMonth).toBe(2023 * 12 + 10)
    expect(domain?.maxMonth).toBe(2024 * 12 + 2)
  })

  it('skips blank members but rejects fields with non-date values', () => {
    expect(timelineDomainOf(['2024-01-01', ''], [item(0), item(1)])).not.toBeNull()
    expect(timelineDomainOf(['2024-01-01', 'West'], [item(0), item(1)])).toBeNull()
    expect(timelineDomainOf([true], [item(0)])).toBeNull()
    expect(timelineDomainOf([''], [item(0)])).toBeNull()
  })
})

describe('timelineSelection', () => {
  const domain = timelineDomainOf(
    ['2023-11-02', '2024-01-15', '2024-03-01'],
    [item(0), item(1), item(2)],
  )!

  it('returns null (no filter) for a null range', () => {
    expect(timelineSelection(domain, null)).toBeNull()
  })

  it('selects only members inside the inclusive month range', () => {
    expect(timelineSelection(domain, { start: 2024 * 12, end: 2024 * 12 + 2 })).toEqual([1, 2])
    expect(timelineSelection(domain, { start: 2024 * 12, end: 2024 * 12 })).toEqual([1])
    expect(timelineSelection(domain, { start: 2025 * 12, end: 2025 * 12 })).toEqual([])
  })
})

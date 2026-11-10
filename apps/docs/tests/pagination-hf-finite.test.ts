import { describe, expect, it } from 'vitest'
import { formatPageNumber } from '../src/renderer/pagination-hf'

describe('formatPageNumber finite guard', () => {
  it('keeps normal numbering unchanged', () => {
    expect(formatPageNumber(4, 'upperRoman')).toBe('IV')
    expect(formatPageNumber(27, 'upperLetter')).toBe('AA')
    expect(formatPageNumber(7, 'decimal')).toBe('7')
  })

  it('never hangs or OOMs on non-finite input', () => {
    const start = Date.now()
    for (const n of [NaN, Infinity, -Infinity]) {
      for (const fmt of ['upperRoman', 'lowerRoman', 'upperLetter', 'upperGreek', 'decimal']) {
        const out = formatPageNumber(n, fmt)
        expect(typeof out).toBe('string')
        expect(out.length).toBeLessThan(100)
      }
    }
    expect(Date.now() - start).toBeLessThan(5000)
  })

  it('caps huge values instead of building megabyte strings', () => {
    const out = formatPageNumber(1e12, 'upperRoman')
    expect(out.length).toBeLessThan(100_000)
    expect(formatPageNumber(1e12, 'upperGreek').length).toBeLessThan(100_000)
  })
})

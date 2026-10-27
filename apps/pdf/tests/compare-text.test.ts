import { describe, expect, it } from 'vitest'
import {
  compareDocuments,
  diffLines,
  normalizePageText,
  type PageText,
} from '../src/renderer/compare-text'

describe('normalizePageText', () => {
  it('collapses whitespace, trims and drops blank lines', () => {
    expect(normalizePageText('  a   b  \n\n\tc\td\n   \n')).toEqual(['a b', 'c d'])
  })
})

describe('diffLines', () => {
  it('reports an insertion with its index in the new page', () => {
    const { added, removed } = diffLines(['a', 'c'], ['a', 'b', 'c'])
    expect(removed).toEqual([])
    expect(added).toEqual([{ line: 'b', index: 1 }])
  })

  it('reports a deletion with its index in the old page', () => {
    const { added, removed } = diffLines(['a', 'b', 'c'], ['a', 'c'])
    expect(added).toEqual([])
    expect(removed).toEqual([{ line: 'b', index: 1 }])
  })

  it('reads a swapped line as one removal plus one addition', () => {
    const { added, removed } = diffLines(['old'], ['new'])
    expect(removed).toEqual([{ line: 'old', index: 0 }])
    expect(added).toEqual([{ line: 'new', index: 0 }])
  })

  it('returns nothing for identical pages', () => {
    const { added, removed } = diffLines(['a', 'b'], ['a', 'b'])
    expect(added).toEqual([])
    expect(removed).toEqual([])
  })
})

describe('compareDocuments', () => {
  const page = (text: string): PageText => ({ pageIndex: 0, text })

  it('calls identical documents equal', () => {
    const result = compareDocuments([page('a\nb')], [page('a\nb')])
    expect(result.equal).toBe(true)
    expect(result.pages).toHaveLength(1)
  })

  it('adds the extra pages of the longer document as wholes', () => {
    const result = compareDocuments([page('a')], [page('a'), page('extra\nlines')])
    expect(result.equal).toBe(false)
    expect(result.pages[1]!.addedCount).toBe(2)
    expect(result.pages[1]!.removedCount).toBe(0)
    expect(result.pages[1]!.added.map((l) => l.line)).toEqual(['extra', 'lines'])
  })

  it('counts per-page changes independently', () => {
    const result = compareDocuments(
      [page('same\nold'), page('untouched')],
      [page('same\nnew'), page('untouched')],
    )
    expect(result.pages[0]!.removedCount).toBe(1)
    expect(result.pages[0]!.addedCount).toBe(1)
    expect(result.pages[1]!.addedCount).toBe(0)
  })

  it('flags past-cap pages as truncated instead of diffing', () => {
    const big = Array.from({ length: 4001 }, (_, i) => `line${i}`).join('\n')
    const small = 'one line'
    const result = compareDocuments([page(big)], [page(small)])
    expect(result.pages[0]!.truncated).toBe(true)
    expect(result.pages[0]!.removedCount).toBe(4001)
    expect(result.pages[0]!.addedCount).toBe(1)
    expect(result.pages[0]!.removed).toEqual([])
  })
})

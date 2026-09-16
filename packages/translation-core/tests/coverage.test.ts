/**
 * Coverage math.
 *
 * The rule has to match the Python handlers byte for byte, because a wrong
 * answer here means the UI tells the user a document is fully translated when
 * it is not (or the reverse, and they re-run a pass for nothing).
 */
import { describe, expect, it } from 'vitest'

import { assessCoverage, isSegmentCovered, mergeDictionary } from '../src/coverage'

describe('isSegmentCovered', () => {
  const dict = { 克重: 'GSM', '面料克重 220 g/m²': 'fabric weight 220 g/m²' }

  it('matches an exact key', () => {
    expect(isSegmentCovered('克重', dict)).toBe(true)
  })

  it('matches a key appearing inside a longer segment', () => {
    // The handlers fall back to longest-key-first substring replacement.
    expect(isSegmentCovered('本批克重检验合格', dict)).toBe(true)
  })

  it('does not match when no key appears', () => {
    expect(isSegmentCovered('本批订单三周完成', dict)).toBe(false)
  })

  it('does not treat an empty target as coverage', () => {
    // A gap-fill template has empty values; the handler leaves those strings
    // alone, and a blank target would otherwise delete the text.
    expect(isSegmentCovered('克重', { 克重: '' })).toBe(false)
    expect(isSegmentCovered('克重', { 克重: '   ' })).toBe(false)
    expect(isSegmentCovered('克重', { 克重: undefined })).toBe(false)
  })

  it('ignores empty sources', () => {
    expect(isSegmentCovered('anything', { '': 'x' })).toBe(false)
  })

  it('handles an empty segment', () => {
    expect(isSegmentCovered('', dict)).toBe(false)
  })
})

describe('assessCoverage', () => {
  it('reports covered/uncovered and preserves order', () => {
    const segments = ['克重', '订单编号', '色牢度', '批次']
    const report = assessCoverage(segments, { 克重: 'GSM', 色牢度: 'color fastness' })
    expect(report.total).toBe(4)
    expect(report.covered).toBe(2)
    expect(report.uncovered).toEqual(['订单编号', '批次'])
    expect(report.ratio).toBe(0.5)
    expect(report.exact).toBe(2)
    expect(report.partial).toEqual([])
  })

  it('separates exact hits from partial (mixed-language) ones', () => {
    // A key that merely appears inside a longer segment leaves the rest of that
    // segment in the source language, which is worth flagging on its own.
    const report = assessCoverage(['克重', '产品验收报告已提交。'], { 克重: 'GSM', 产品验收报告: 'Product Acceptance Report' })
    expect(report.exact).toBe(1)
    expect(report.partial).toEqual(['产品验收报告已提交。'])
    expect(report.covered).toBe(2)
    expect(report.uncovered).toEqual([])
  })

  it('is 100% when there is nothing to translate', () => {
    expect(assessCoverage([], {}).ratio).toBe(1)
    expect(assessCoverage([], {}).uncovered).toEqual([])
  })

  it('is 0% with an empty dictionary', () => {
    const report = assessCoverage(['a', 'b'], {})
    expect(report.covered).toBe(0)
    expect(report.ratio).toBe(0)
  })
})

describe('mergeDictionary', () => {
  it('keeps existing values and sorts keys', () => {
    const merged = mergeDictionary({ b: 'B', a: 'A' }, { a: 'override', c: 'C' })
    expect(merged).toEqual({ a: 'A', b: 'B', c: 'C' })
    expect(Object.keys(merged)).toEqual(['a', 'b', 'c'])
  })

  it('drops entries without a usable target', () => {
    const merged = mergeDictionary({ a: 'A', bad: undefined }, { c: undefined, d: 'D' })
    expect(merged).toEqual({ a: 'A', d: 'D' })
  })
})

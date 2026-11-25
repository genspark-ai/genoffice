import { describe, expect, it } from 'vitest'
import {
  computeListNumberSites,
  paragraphNumberText,
  type ListItemRef,
  type NumberingDef,
  type NumberingLevel,
} from '../src/index'

const level = (numFmt: string, lvlText: string): NumberingLevel => ({ numFmt, lvlText, start: 1 })

const def = (numId: string, levels: NumberingLevel[], abstractNumId = numId): NumberingDef => ({
  numId,
  abstractNumId,
  levels: Object.fromEntries(levels.map((l, i) => [i, l])),
  startOverrides: {},
})

/** 1. / (a) / (i): each level shows its own number only */
const CLAUSES = def('1', [
  level('decimal', '%1.'),
  level('lowerLetter', '(%2)'),
  level('lowerRoman', '(%3)'),
])
/** 1. / 1.1 / 1.1.1: each level carries the ones above it */
const LEGAL = def('2', [
  level('decimal', '%1.'),
  level('decimal', '%1.%2'),
  level('decimal', '%1.%2.%3'),
])
/** 1. / 1.1 / (a): the third level stands under a number that carries the first */
const MIXED = def('3', [
  level('decimal', '%1.'),
  level('decimal', '%1.%2'),
  level('lowerLetter', '(%3)'),
])
const DEFS = new Map([CLAUSES, LEGAL, MIXED].map((d) => [d.numId, d]))

const items = (numId: string, ...ilvls: number[]): ListItemRef[] =>
  ilvls.map((ilvl) => ({ numId, ilvl }))

describe('the number a REF field shows for a numbered paragraph', () => {
  it('drops one final period and nothing else', () => {
    const one = (lvlText: string) => {
      const d = def('9', [level('decimal', lvlText)])
      const sites = computeListNumberSites(items('9', 0), new Map([['9', d]]))
      return (['relative', 'none', 'full'] as const).map((c) => paragraphNumberText(sites, 0, c))
    }
    // the results Word saved for these level texts
    expect(one('%1.')).toEqual(['1', '1', '1'])
    expect(one('%1..')).toEqual(['1.', '1.', '1.'])
    expect(one('%1.).')).toEqual(['1.)', '1.)', '1.)'])
    expect(one('%1)')).toEqual(['1)', '1)', '1)'])
    expect(one('[%1]')).toEqual(['[1]', '[1]', '[1]'])
    expect(one('Part %1')).toEqual(['Part 1', 'Part 1', 'Part 1'])
  })

  it('no context is the paragraph number alone', () => {
    // 1. (a) (b) (i)
    const sites = computeListNumberSites(items('1', 0, 1, 1, 2), DEFS)
    expect(paragraphNumberText(sites, 3, 'none')).toBe('(i)')
    expect(paragraphNumberText(sites, 2, 'none')).toBe('(b)')
  })

  it('full context puts the levels above in front where the number does not carry them', () => {
    const clauses = computeListNumberSites(items('1', 0, 1, 1, 2), DEFS)
    expect(paragraphNumberText(clauses, 3, 'full')).toBe('1.(b)(i)')
    const legal = computeListNumberSites(items('2', 0, 1, 2, 2), DEFS)
    expect(paragraphNumberText(legal, 3, 'full')).toBe('1.1.2')
    // 3. 3.1 3.2 (a): Word's result for this shape is "3.2(a)"
    const mixed = computeListNumberSites(items('3', 0, 0, 0, 1, 1, 2), DEFS)
    expect(paragraphNumberText(mixed, 5, 'full')).toBe('3.2(a)')
  })

  it('relative context leaves out the levels the field shares with the paragraph', () => {
    // 1. (a) (b) (c) 2. (a)
    const sites = computeListNumberSites(items('1', 0, 1, 1, 1, 0, 1), DEFS)
    expect(paragraphNumberText(sites, 2, 'relative', { from: 3 })).toBe('(b)')
    expect(paragraphNumberText(sites, 2, 'relative', { from: 0 })).toBe('(b)')
    expect(paragraphNumberText(sites, 2, 'relative', { from: 5 })).toBe('1.(b)')
    expect(paragraphNumberText(sites, 2, 'relative')).toBe('1.(b)')
    // a field under the paragraph it points at
    expect(paragraphNumberText(sites, 0, 'relative', { from: 2 })).toBe('1')
    // 4. 4.6 (c): from under 4.6, Word shows "4.6"
    const mixed = computeListNumberSites(items('3', 0, 1, 2), DEFS)
    expect(paragraphNumberText(mixed, 1, 'relative', { from: 2 })).toBe('1.1')
  })

  it('without the words: the numbers of one level, nothing where two levels are joined', () => {
    const worded = def('8', [level('decimal', 'Article %1.'), level('decimal', 'Section %1.%2')])
    const sites = computeListNumberSites(items('8', 0, 1), new Map([['8', worded]]))
    expect(paragraphNumberText(sites, 1, 'full')).toBe('Section 1.1')
    expect(paragraphNumberText(sites, 1, 'full', { bare: true })).toBe('1.1')
    expect(paragraphNumberText(sites, 0, 'relative', { bare: true })).toBe('1')
    // 1. (a): what Word keeps between the two levels is not known
    const clauses = computeListNumberSites(items('1', 0, 1), DEFS)
    expect(paragraphNumberText(clauses, 1, 'full', { bare: true })).toBeNull()
    expect(paragraphNumberText(clauses, 1, 'relative', { from: 0, bare: true })).toBe('a')
  })

  it('a separator stands between the levels in place of their own period', () => {
    const sites = computeListNumberSites(items('1', 0, 1, 2), DEFS)
    expect(paragraphNumberText(sites, 2, 'full', { separator: '-' })).toBe('1-(a)-(i)')
    const legal = computeListNumberSites(items('2', 0, 1), DEFS)
    expect(paragraphNumberText(legal, 1, 'full', { separator: '-' })).toBe('1.1')
  })

  it('lists that share no definition do not lend each other levels, and a bullet has no number', () => {
    const bullets = def('7', [level('bullet', '•')])
    const defs = new Map([...DEFS, ['7', bullets]])
    const sites = computeListNumberSites(
      [...items('1', 0), ...items('2', 1), ...items('7', 0)],
      defs,
    )
    expect(sites[1]?.parent).toBe(-1)
    expect(paragraphNumberText(sites, 2, 'full')).toBeNull()
    expect(paragraphNumberText(sites, 5, 'full')).toBeNull()
  })
})

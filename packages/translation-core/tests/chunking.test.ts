import { describe, expect, it } from 'vitest'

import { chunkDocument, makeUnitId } from '../src/chunking'

describe('chunkDocument', () => {
  it('returns an empty list for blank input', () => {
    expect(chunkDocument('', 'doc')).toEqual([])
    expect(chunkDocument('   \n  \t  ', 'doc')).toEqual([])
  })

  it('splits on blank lines into paragraph units with stable ids', () => {
    const units = chunkDocument('First paragraph.\n\nSecond paragraph here.', 'doc')
    expect(units).toHaveLength(2)
    expect(units[0].unitId).toBe('doc-0')
    expect(units[0].order).toBe(0)
    expect(units[0].sourceText).toBe('First paragraph.')
    expect(units[1].unitId).toBe('doc-1')
    expect(units[1].sourceText).toBe('Second paragraph here.')
  })

  it('chunks long paragraphs at sentence boundaries', () => {
    const long = Array.from({ length: 8 }, (_, i) => `Sentence number ${i + 1} is here.`).join(' ')
    const units = chunkDocument(long, 'doc', 60)
    expect(units.length).toBeGreaterThan(1)
    for (const u of units) {
      expect(u.sourceText.length).toBeLessThanOrEqual(60)
    }
  })

  it('hard-splits a single sentence that is longer than the cap', () => {
    const long = 'word '.repeat(200).trim() // 999 chars, no sentence boundary
    const units = chunkDocument(long, 'doc', 100)
    expect(units.length).toBeGreaterThan(1)
    for (const u of units) {
      expect(u.sourceText.length).toBeLessThanOrEqual(100)
    }
  })
})

describe('makeUnitId', () => {
  it('combines parent and offset into a stable id', () => {
    expect(makeUnitId('document', 3)).toBe('document-3')
    expect(makeUnitId('document', '0')).toBe('document-0')
  })
})

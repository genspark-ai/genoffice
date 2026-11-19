import { beforeAll, describe, expect, it } from 'vitest'
import {
  breakPoints,
  hyphenWords,
  hyphenationConfigOf,
  hyphenationDictFor,
  hyphenationSettingsOf,
  hyphenationZoneApplies,
  loadHyphenation,
} from '../src/renderer/editor/hyphenation'

const split = (dict: string, word: string) => {
  const points = breakPoints(dict, word)
  return points.reduce<string[]>(
    (parts, p, i) => [...parts, word.slice(p, points[i + 1] ?? word.length)],
    [word.slice(0, points[0] ?? word.length)],
  )
}

describe('hyphenation patterns', () => {
  beforeAll(async () => {
    expect(await loadHyphenation('en-us')).toBe(true)
    expect(await loadHyphenation('de')).toBe(true)
  })

  it('maps a document language to a pattern set', () => {
    expect(hyphenationDictFor('en-US')).toBe('en-us')
    expect(hyphenationDictFor('en-AU')).toBe('en-gb')
    expect(hyphenationDictFor('de-DE')).toBe('de')
    expect(hyphenationDictFor('pt-BR')).toBe('pt')
    expect(hyphenationDictFor(undefined)).toBe('en-us')
    expect(hyphenationDictFor('ru-RU')).toBeNull()
    expect(hyphenationDictFor('zh-CN')).toBeNull()
  })

  it('has no break points before the patterns load', () => {
    expect(breakPoints('fr', 'internationalisation')).toEqual([])
  })

  it('breaks words the way the patterns say, whatever the case', () => {
    expect(split('en-us', 'hyphenation')).toEqual(['hy', 'phen', 'ation'])
    expect(split('en-us', 'HYPHENATION')).toEqual(['HY', 'PHEN', 'ATION'])
    expect(split('en-us', 'Hyphenation')).toEqual(['Hy', 'phen', 'ation'])
    expect(split('de', 'Silbentrennung')).toEqual(['Sil', 'ben', 'tren', 'nung'])
  })

  it('leaves short words whole', () => {
    expect(breakPoints('en-us', 'into')).toEqual([])
    expect(hyphenWords('a into the', 'en-us', true)).toEqual([])
  })

  it('finds the words of a line with their offsets', () => {
    const words = hyphenWords('An international table', 'en-us', true)
    expect(words.map((w) => [w.from, w.to, w.text])).toEqual([
      [3, 16, 'international'],
      [17, 22, 'table'],
    ])
    expect(words[0].points).toEqual([2, 5, 7, 11])
  })

  it('skips words in capitals only when told to', () => {
    const text = 'INTERNATIONAL International international'
    expect(hyphenWords(text, 'en-us', true).map((w) => w.text)).toEqual([
      'INTERNATIONAL',
      'International',
      'international',
    ])
    expect(hyphenWords(text, 'en-us', false).map((w) => w.text)).toEqual([
      'International',
      'international',
    ])
  })

  it('keeps addresses, codes and words with an optional hyphen whole', () => {
    expect(hyphenWords('https://example.com/international', 'en-us', true)).toEqual([])
    expect(hyphenWords('someone@international.example', 'en-us', true)).toEqual([])
    expect(hyphenWords('interna\u00adtional', 'en-us', true)).toEqual([])
    expect(hyphenWords('model2international', 'en-us', true)).toEqual([])
  })

  it('breaks the parts of a compound on their own', () => {
    const words = hyphenWords('well-international', 'en-us', true)
    expect(words.map((w) => [w.from, w.text])).toEqual([[5, 'international']])
  })
})

describe('hyphenation settings of a document', () => {
  it('reads Word defaults from a file that says nothing', () => {
    expect(hyphenationSettingsOf({})).toEqual({ caps: true, limit: 0, zoneTwips: 360 })
  })

  it('reads the file values and lets the session edit win', () => {
    const parsed = {
      doNotHyphenateCaps: true,
      consecutiveHyphenLimit: 2,
      hyphenationZoneTwips: 425,
    }
    expect(hyphenationSettingsOf(parsed)).toEqual({ caps: false, limit: 2, zoneTwips: 425 })
    expect(hyphenationSettingsOf(parsed, { caps: true, limit: 0 })).toEqual({
      caps: true,
      limit: 0,
      zoneTwips: 425,
    })
  })

  it('switches our hyphenation off when the document does not hyphenate', () => {
    expect(hyphenationConfigOf({ docDefaults: { lang: 'en-US' } }, false).dict).toBeNull()
    expect(hyphenationConfigOf({ docDefaults: { lang: 'en-US' } }, true).dict).toBe('en-us')
    expect(hyphenationConfigOf({ docDefaults: { lang: 'ru-RU' } }, true).dict).toBeNull()
  })

  it('applies the zone only in the layout modes that have one', () => {
    const parsed = { hyphenationZoneTwips: 300, docDefaults: { lang: 'en-US' } }
    expect(hyphenationZoneApplies(parsed)).toBe(true)
    expect(hyphenationConfigOf(parsed, true).zonePx).toBe(20)
    expect(hyphenationZoneApplies({ ...parsed, compatibilityMode: 15 })).toBe(false)
    expect(hyphenationConfigOf({ ...parsed, compatibilityMode: 15 }, true).zonePx).toBe(0)
  })
})

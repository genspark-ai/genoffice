import { describe, expect, it } from 'vitest'
import { parseOutline, PLACEHOLDER } from '../src/slides/outline'

const HARBOR = String.fromCharCode(0x6e2f, 0x53e3)

const page = (over: Record<string, unknown> = {}) => ({
  title: 'Why logistics margins fell 40% in two years',
  type: 'content',
  layout: 'three_column_cards',
  brief:
    'Three cards: fuel cost up 31% (2024 vs 2022), driver wages up 18%, spot rates down 22%; source line under each card.',
  image_queries: [],
  ...over,
})

const deck = (pages: unknown[], over: Record<string, unknown> = {}) =>
  JSON.stringify({ core_hook: 'Margins fell 40% while volume grew', pages, ...over })

describe('parseOutline', () => {
  it('accepts a well-formed outline with no findings', () => {
    const r = parseOutline(
      deck([
        page({ type: 'cover', layout: 'cover_typography_hero', title: 'The 40% squeeze' }),
        page(),
        page({ layout: 'hero_big_number' }),
        page({ type: 'data', layout: 'kpi_cards_row' }),
        page({ type: 'closing', layout: 'closing_cta' }),
      ]),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.issues).toEqual([])
    expect(r.outline.pages).toHaveLength(5)
    expect(r.outline.core_hook).toBe('Margins fell 40% while volume grew')
  })

  it('flags unknown layouts, placeholders, repeats and a missing core hook', () => {
    const r = parseOutline(
      JSON.stringify({
        pages: [
          page({ type: 'cover', layout: 'cover_typography_hero' }),
          page({
            layout: 'hero_big_number',
            brief: 'Revenue grew XX% year over year across all regions, per the deck brief.',
          }),
          page({ layout: 'hero_big_number' }),
          page({ type: 'data', layout: 'three_column_cards' }),
          page({ type: 'closing', layout: 'closing_cta' }),
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    const errors = r.issues
      .filter((i) => i.level === 'error')
      .map((i) => `${i.page ?? '-'}: ${i.message}`)
    expect(errors).toEqual(
      expect.arrayContaining([
        expect.stringContaining('1: placeholder "XX%"'),
        expect.stringContaining('2: layout hero_big_number repeats'),
        expect.stringContaining('3: "layout" three_column_cards is not a data variant'),
        expect.stringContaining('-: "core_hook" is missing'),
      ]),
    )
  })

  it('accepts apparel sizes such as XXL while still catching XX%', () => {
    expect(PLACEHOLDER.test('Available in XS, S, M, L, XL and XXL.')).toBe(false)
    expect(PLACEHOLDER.test('XXXL hoodie, size xxl')).toBe(false)
    expect(PLACEHOLDER.test('Revenue grew XX% year over year')).toBe(true)
    expect(PLACEHOLDER.test('Serving XX customers')).toBe(true)
  })

  it('warns about thin briefs, CJK image queries, little variety and missing cover/closing', () => {
    const r = parseOutline(
      deck([
        page({ brief: 'Three points.', image_queries: [HARBOR] }),
        page({ layout: 'hero_big_number' }),
        page(),
        page({ layout: 'hero_big_number' }),
      ]),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.issues.filter((i) => i.level === 'error')).toEqual([])
    const warnings = r.issues.map((i) => i.message)
    expect(warnings).toEqual(
      expect.arrayContaining([
        expect.stringContaining('"brief" is thin'),
        expect.stringContaining('should be an English phrase'),
        expect.stringContaining('only 2 layout variant(s)'),
        'the first page is not a cover',
        'the last page is not a closing page',
      ]),
    )
    expect(r.outline.pages[0]!.image_queries).toEqual([HARBOR])
  })

  it('rejects documents that are not an outline', () => {
    expect(parseOutline('nope')).toMatchObject({
      ok: false,
      error: expect.stringContaining('invalid JSON'),
    })
    expect(parseOutline('[]')).toMatchObject({
      ok: false,
      error: expect.stringContaining('expected {'),
    })
    expect(parseOutline('{"pages":[]}')).toMatchObject({
      ok: false,
      error: expect.stringContaining('empty'),
    })
    expect(parseOutline(deck(new Array(61).fill(page())))).toMatchObject({
      ok: false,
      error: expect.stringContaining('too many pages'),
    })
  })

  it('rejects oversized raw outlines and caps image queries per page', () => {
    expect(parseOutline('x'.repeat(600_000))).toMatchObject({
      ok: false,
      error: expect.stringContaining('too large'),
    })
    const many = { ...page(), image_queries: Array.from({ length: 30 }, (_, i) => `scene ${i}`) }
    const r = parseOutline(deck([many]))
    expect(r.ok).toBe(true)
    if (r.ok) {
      expect(r.outline.pages[0]!.image_queries).toHaveLength(8)
      expect(r.issues.some((i) => i.message.includes('capped at 8'))).toBe(true)
    }
  })
})

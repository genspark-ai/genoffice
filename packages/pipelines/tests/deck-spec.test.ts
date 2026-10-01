import { describe, expect, it } from 'vitest'
import { parseDeckSpec } from '../src/slides/deck-spec'
import { parsePageSpec } from '../src/slides/page-spec'
import type { OutlinePage } from '../src/slides/outline'
import {
  checkPageAgainstOutline,
  offPaletteColors,
  pageColors,
  stylePalette,
} from '../src/slides/stage-check'

const textEl = (text: string, over: Record<string, unknown> = {}) => ({
  type: 'text',
  x: 80,
  y: 60,
  w: 800,
  h: 90,
  paragraphs: [{ runs: [{ text, sizePt: 24, color: '#F5F5F5' }] }],
  ...over,
})

const validPage = (text = 'Quarterly wins across regions') => ({
  background: '#0E1A2B',
  elements: [textEl(text)],
})

const entry: OutlinePage = {
  title: 'Where the margin went',
  type: 'content',
  layout: 'three_column_cards',
  brief: 'Three cards: fuel up 31 percent, wages up 18 percent, spot rates down 22 percent.',
  image_queries: ['container port cranes at dawn'],
}

function stagedPage(extra: Record<string, unknown>, texts: string[], withImage = false) {
  const raw = {
    ...extra,
    background: '#0E1A2B',
    elements: [
      ...texts.map((t, i) => textEl(t, { y: 80 + i * 120 })),
      { type: 'shape', shape: 'rect', x: 80, y: 600, w: 200, h: 40, fill: '#1F3A5F' },
      ...(withImage
        ? [{ type: 'image', url: 'https://example.com/photo.png', x: 900, y: 80, w: 300, h: 300 }]
        : []),
    ],
  }
  const parsed = parsePageSpec(JSON.stringify(raw), undefined, undefined, { localImages: true })
  if (!parsed.ok) throw new Error(parsed.error)
  return { raw, spec: parsed.spec }
}

describe('parseDeckSpec refusals', () => {
  it('rejects invalid JSON', () => {
    expect(parseDeckSpec('not json')).toMatchObject({
      ok: false,
      error: expect.stringContaining('invalid JSON'),
    })
  })

  it('rejects a document without a pages array', () => {
    expect(parseDeckSpec('{"background":"#0E1A2B"}')).toMatchObject({
      ok: false,
      error: expect.stringContaining('expected { "pages"'),
    })
  })

  it('rejects an empty pages array', () => {
    expect(parseDeckSpec('{"pages":[]}')).toMatchObject({
      ok: false,
      error: expect.stringContaining('empty'),
    })
  })

  it('rejects more pages than the deck limit', () => {
    const raw = JSON.stringify({ pages: new Array(61).fill(validPage()) })
    expect(parseDeckSpec(raw)).toMatchObject({
      ok: false,
      error: expect.stringContaining('too many pages'),
    })
  })

  it('rejects an over-large raw payload before parsing it', () => {
    const filler = 'x'.repeat(20_000)
    const raw = JSON.stringify({ pages: new Array(60).fill(validPage(filler)) })
    expect(raw.length).toBeGreaterThan(1_000_000)
    expect(parseDeckSpec(raw)).toMatchObject({
      ok: false,
      error: expect.stringContaining('too large'),
    })
  })

  it('accepts a spec over the cap when localImages is set (the CLI data: path)', () => {
    // ~2 MB of base64 in one data: URL — a documented, legitimate CLI input that the
    // shared cap would otherwise reject outright.
    const big = 'A'.repeat(2_000_000)
    const raw = JSON.stringify({
      pages: [
        {
          background: '#0E1A2B',
          elements: [
            { type: 'image', url: `data:image/png;base64,${big}`, x: 80, y: 60, w: 800, h: 450 },
          ],
        },
      ],
    })
    expect(raw.length).toBeGreaterThan(1_000_000)
    expect(parseDeckSpec(raw, { localImages: true })).toMatchObject({ ok: true })
    // the default http(s)-only path still refuses it
    expect(parseDeckSpec(raw)).toMatchObject({
      ok: false,
      error: expect.stringContaining('too large'),
    })
  })

  it('refuses when no page survives validation', () => {
    const raw = JSON.stringify({
      pages: [{ background: '#0E1A2B', elements: [] }, { background: '#0E1A2B' }],
    })
    const r = parseDeckSpec(raw)
    expect(r.ok).toBe(false)
    if (r.ok) return
    expect(r.error).toContain('no valid pages')
  })

  it('keeps valid pages and reports the dropped ones', () => {
    const raw = JSON.stringify({
      pages: [validPage('First page wins'), { background: '#0E1A2B', elements: [] }],
    })
    const r = parseDeckSpec(raw)
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.pages).toHaveLength(1)
    expect(r.issues).toHaveLength(1)
    expect(r.issues[0]).toMatchObject({ page: 1, error: expect.stringContaining('elements') })
  })

  it('accepts a bare array and a single page object', () => {
    const bare = parseDeckSpec(JSON.stringify([validPage('Bare page wins')]))
    expect(bare.ok).toBe(true)
    if (!bare.ok) return
    expect(bare.spec.pages).toHaveLength(1)

    const single = parseDeckSpec(JSON.stringify(validPage('Single page wins')))
    expect(single.ok).toBe(true)
    if (!single.ok) return
    expect(single.spec.pages).toHaveLength(1)
  })
})

describe('staged-deck QA gate: outline agreement', () => {
  it('passes a page that echoes its outline entry', () => {
    const p = stagedPage(
      { title: entry.title, type: 'content', layout: 'three_column_cards' },
      ['Where the margin went', 'Fuel up 31 percent'],
      true,
    )
    expect(checkPageAgainstOutline(p.raw, p.spec, entry, 1)).toEqual([])
  })

  it('errors when echoed type or layout disagrees with the outline', () => {
    const p = stagedPage(
      { title: entry.title, type: 'data', layout: 'hero_big_number' },
      ['Where the margin went', 'Fuel up 31 percent'],
      true,
    )
    const issues = checkPageAgainstOutline(p.raw, p.spec, entry, 1)
    expect(issues.map((i) => i.level)).toEqual(['error', 'error'])
    expect(issues[0]!.message).toContain('type "data" is not outline pages[1]')
    expect(issues[1]!.message).toContain('layout "hero_big_number"')
  })

  it('errors on placeholder copy in the staged text', () => {
    const p = stagedPage(
      { title: entry.title },
      ['Where the margin went', 'Growth of TBD here'],
      true,
    )
    const issues = checkPageAgainstOutline(p.raw, p.spec, entry, 1)
    expect(issues.some((i) => i.level === 'error' && i.message.includes('placeholder'))).toBe(true)
  })

  it('warns when the outline title is missing and when a planned photo is absent', () => {
    const p = stagedPage({}, ['Costs went up across every lane'])
    expect(checkPageAgainstOutline(p.raw, p.spec, entry, 1)).toEqual([
      { page: 1, level: 'warning', message: expect.stringContaining('does not appear') },
      { page: 1, level: 'warning', message: expect.stringContaining('plans 1 photo(s)') },
    ])
  })
})

describe('staged-deck QA gate: palette agreement', () => {
  it('collects palette colors and passes an on-palette page', () => {
    const palette = stylePalette(
      '- Backgrounds: main #0e1a2b, cards #1F3A5F\n- Text: #F5F5F5\nplain words',
    )
    expect([...palette]).toEqual(['#0E1A2B', '#1F3A5F', '#F5F5F5'])
    const p = stagedPage({}, ['On palette page'])
    expect(pageColors(p.spec)).toEqual(['#0E1A2B', '#F5F5F5', '#1F3A5F'])
    expect(offPaletteColors(p.spec, palette)).toEqual([])
  })

  it('flags colors the style sheet never names', () => {
    const p = stagedPage({}, ['Off palette page'])
    p.spec.elements.push({
      type: 'shape',
      shape: 'rect',
      x: 0,
      y: 0,
      w: 10,
      h: 10,
      fill: '#FF0000',
    })
    expect(offPaletteColors(p.spec, new Set(['#0E1A2B', '#F5F5F5', '#1F3A5F']))).toEqual([
      '#FF0000',
    ])
  })

  it('always allows white and black and skips the check without a palette', () => {
    const p = stagedPage({}, ['Neutral page'])
    p.spec.elements.push({
      type: 'shape',
      shape: 'rect',
      x: 0,
      y: 0,
      w: 10,
      h: 10,
      fill: '#FFFFFF',
    })
    expect(offPaletteColors(p.spec, new Set(['#0E1A2B']))).toEqual(
      expect.arrayContaining(['#F5F5F5', '#1F3A5F']),
    )
    expect(offPaletteColors(p.spec, new Set(['#0E1A2B']))).not.toContain('#FFFFFF')
    expect(offPaletteColors(p.spec, new Set())).toEqual([])
  })
})

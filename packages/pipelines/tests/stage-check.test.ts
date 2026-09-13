import { describe, expect, it } from 'vitest'
import { parsePageSpec } from '../src/slides/page-spec'
import type { OutlinePage } from '../src/slides/outline'
import {
  checkPageAgainstOutline,
  offPaletteColors,
  pageColors,
  stylePalette,
} from '../src/slides/stage-check'

const entry: OutlinePage = {
  title: 'Where the margin went',
  type: 'content',
  layout: 'three_column_cards',
  brief: 'Three cards: fuel +31%, wages +18%, spot rates -22%.',
  image_queries: ['container port cranes at dawn'],
}

function page(extra: Record<string, unknown>, texts: string[], withImage = false) {
  const raw = {
    ...extra,
    background: '#0E1A2B',
    elements: [
      ...texts.map((t, i) => ({
        type: 'text',
        x: 80,
        y: 80 + i * 120,
        w: 800,
        h: 80,
        paragraphs: [{ runs: [{ text: t, sizePt: 24, color: '#F5F5F5' }] }],
      })),
      { type: 'shape', shape: 'rect', x: 80, y: 600, w: 200, h: 40, fill: '#1F3A5F' },
      ...(withImage ? [{ type: 'image', url: 'x.png', x: 900, y: 80, w: 300, h: 300 }] : []),
    ],
  }
  const parsed = parsePageSpec(JSON.stringify(raw), undefined, undefined, { localImages: true })
  if (!parsed.ok) throw new Error(parsed.error)
  return { raw, spec: parsed.spec }
}

describe('checkPageAgainstOutline', () => {
  it('passes a page that echoes its entry, carries the title and has the planned photo', () => {
    const p = page(
      { title: entry.title, type: 'content', layout: 'three_column_cards' },
      ['Where the margin went', 'Fuel +31%'],
      true,
    )
    expect(checkPageAgainstOutline(p.raw, p.spec, entry, 1)).toEqual([])
  })

  it('errors on an echoed field that disagrees and on placeholder copy', () => {
    const p = page(
      { layout: 'hero_big_number', type: 'data' },
      ['Where the margin went', 'XX% up'],
      true,
    )
    const issues = checkPageAgainstOutline(p.raw, p.spec, entry, 1)
    expect(issues.map((i) => i.level)).toEqual(['error', 'error', 'error'])
    expect(issues[0]!.message).toContain('type "data" is not outline pages[1]\'s "content"')
    expect(issues[1]!.message).toContain('layout "hero_big_number"')
    expect(issues[2]!.message).toContain('placeholder "XX%"')
  })

  it('warns when the title is missing from the text or a planned photo has no image element', () => {
    const p = page({}, ['Costs went up'])
    const issues = checkPageAgainstOutline(p.raw, p.spec, entry, 1)
    expect(issues).toEqual([
      {
        page: 1,
        level: 'warning',
        message: expect.stringContaining('does not appear in the page text'),
      },
      { page: 1, level: 'warning', message: expect.stringContaining('plans 1 photo(s)') },
    ])
  })

  it('matches the title loosely (case, punctuation, spacing)', () => {
    const p = page({}, ['WHERE the margin — went!'], true)
    expect(checkPageAgainstOutline(p.raw, p.spec, entry, 1)).toEqual([])
  })
})

describe('style palette', () => {
  it('collects #RRGGBB values from the style sheet and flags colors a page adds', () => {
    const palette = stylePalette(
      '- Backgrounds: main #0e1a2b, cards #1F3A5F80\n- Text: #F5F5F5\nno color here',
    )
    expect([...palette]).toEqual(['#0E1A2B', '#1F3A5F', '#F5F5F5'])
    const p = page({}, ['x'])
    expect(pageColors(p.spec)).toEqual(['#0E1A2B', '#F5F5F5', '#1F3A5F'])
    expect(offPaletteColors(p.spec, palette)).toEqual([])
    expect(offPaletteColors(p.spec, new Set(['#0E1A2B']))).toEqual(['#F5F5F5', '#1F3A5F'])
  })

  it('never flags white or black and skips the check without a palette', () => {
    const p = page({}, ['x'])
    p.spec.elements.push({
      type: 'shape',
      shape: 'rect',
      x: 0,
      y: 0,
      w: 10,
      h: 10,
      fill: '#FFFFFF',
    })
    expect(offPaletteColors(p.spec, new Set(['#0E1A2B', '#F5F5F5', '#1F3A5F']))).toEqual([])
    expect(offPaletteColors(p.spec, new Set())).toEqual([])
  })
})

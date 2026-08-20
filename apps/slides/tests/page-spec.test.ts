/**
 * Local page generation: LLM spec JSON → parse/validate → build a one-slide
 * pptx directly with pptx-engine primitives. The build tests reopen the bytes
 * with openPptx and assert on the parsed model (true roundtrip, no mocks).
 */
import { describe, it, expect } from 'vitest'
import { openPptx, type TextElement, type PictureElement } from '@genoffice/pptx-engine'
import { parsePageSpec, buildPagePptx, type PageSpec } from '../src/main/page-spec'

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
)

const textSpec = (text: string, extra: Record<string, unknown> = {}) => ({
  type: 'text',
  x: 80,
  y: 60,
  w: 800,
  h: 90,
  paragraphs: [{ runs: [{ text, sizePt: 32, bold: true, color: '#112233' }] }],
  ...extra,
})

describe('parsePageSpec', () => {
  it('accepts fenced JSON with junk around it', () => {
    const raw =
      'Here is the design:\n```json\n{"background":"#0E1A2B","elements":[' +
      JSON.stringify(textSpec('Hello')) +
      ']}\n```\nDone.'
    const r = parsePageSpec(raw)
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.background).toBe('#0E1A2B')
    expect(r.spec.elements).toHaveLength(1)
  })

  it('rejects output without a usable JSON object', () => {
    expect(parsePageSpec('sorry, I cannot').ok).toBe(false)
    expect(parsePageSpec('{"elements":[]}').ok).toBe(false)
    const bad = parsePageSpec('{"elements":[{"type":"text","x":0,"y":0,"w":100,"h":40}]}')
    expect(bad.ok).toBe(false) // text without any runs → all elements dropped
  })

  it('clamps out-of-canvas boxes and drops vanishing ones with warnings', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          { ...textSpec('kept'), x: 1200, w: 400 }, // clamped to 80px wide
          { ...textSpec('gone'), x: 5000, y: 5000 },
          { type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 100, fill: '#FFF' },
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.elements).toHaveLength(2)
    const kept = r.spec.elements[0]!
    expect(kept.x + kept.w).toBeLessThanOrEqual(1280)
    expect(r.warnings.some((w) => w.includes('outside'))).toBe(true)
    // #FFF expands to #FFFFFF
    expect((r.spec.elements[1] as { fill?: string }).fill).toBe('#FFFFFF')
  })

  it('falls back to rect for unknown shapes and drops invisible ones', () => {
    const r = parsePageSpec(
      JSON.stringify({
        elements: [
          { type: 'shape', shape: 'wavyMagicBlob', x: 0, y: 0, w: 10, h: 10, fill: '#123456' },
          { type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10 }, // no fill/stroke → dropped
          { type: 'image', url: 'ftp://nope', x: 0, y: 0, w: 10, h: 10 }, // bad scheme → dropped
          textSpec('t'),
        ],
      }),
    )
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.spec.elements).toHaveLength(2)
    expect((r.spec.elements[0] as { shape: string }).shape).toBe('rect')
  })
})

describe('buildPagePptx', () => {
  const noImages = { fetchImage: async () => null }

  it('builds text/shape/background into a reopenable one-slide pptx', async () => {
    const spec: PageSpec = {
      background: '#0E1A2B',
      elements: [
        {
          type: 'shape',
          shape: 'roundRect',
          x: 80,
          y: 200,
          w: 400,
          h: 200,
          fill: '#FFFFFF14',
          stroke: { color: '#3B82F6', widthPt: 1 },
        },
        {
          type: 'text',
          x: 80,
          y: 60,
          w: 800,
          h: 90,
          paragraphs: [
            {
              align: 'left',
              lineSpacingPct: 110,
              runs: [{ text: 'Quarterly Wins', sizePt: 36, bold: true, color: '#FFFFFF' }],
            },
          ],
        },
      ],
    }
    const { bytes, imageFailures } = await buildPagePptx(spec, noImages)
    expect(imageFailures).toEqual([])
    const opened = await openPptx(bytes)
    expect(opened.deck.slides).toHaveLength(1)
    const slide = opened.deck.slides[0]!
    // The full-bleed background rect is promoted to the slide background at build time
    const texts = slide.elements.filter((e): e is TextElement => e.type === 'text')
    const shapes = slide.elements.filter((e) => e.type === 'shape')
    expect(shapes.length).toBeGreaterThanOrEqual(1)
    expect(texts).toHaveLength(1)
    const run = texts[0]!.text!.paragraphs[0]!.runs[0]!
    expect(run.text).toBe('Quarterly Wins')
    expect(run.bold).toBe(true)
    expect(run.fontSize).toBe(36)
    // 80px at the deck's 1280px-wide canvas = 80 * 9525 EMU
    expect(texts[0]!.transform.offset.x).toBe(80 * 9525)
  })

  it('adds images with cover-crop and reports failed downloads without failing the page', async () => {
    const spec: PageSpec = {
      elements: [
        { type: 'image', url: 'https://ok.example/a.png', x: 0, y: 0, w: 640, h: 720 },
        { type: 'image', url: 'https://dead.example/b.png', x: 640, y: 0, w: 640, h: 720 },
        { type: 'text', x: 100, y: 100, w: 400, h: 60, paragraphs: [{ runs: [{ text: 'cap' }] }] },
      ],
    }
    const { bytes, imageFailures } = await buildPagePptx(spec, {
      fetchImage: async (url) =>
        url.includes('ok.example') ? { bytes: PNG_1PX, ext: 'png' } : null,
      imageDims: () => ({ width: 200, height: 100 }),
    })
    expect(imageFailures).toEqual(['https://dead.example/b.png'])
    const opened = await openPptx(bytes)
    const pics = opened.deck.slides[0]!.elements.filter(
      (e): e is PictureElement => e.type === 'picture',
    )
    expect(pics).toHaveLength(1)
    // 200x100 source into a 640x720 portrait frame → horizontal crop applied
    expect(pics[0]!.srcRect?.l ?? 0).toBeGreaterThan(0)
  })
})

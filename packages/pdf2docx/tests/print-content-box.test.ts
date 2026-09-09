/**
 * P35: browser prints (Chromium printToPDF with margins) — the body wash
 * covers the print content box instead of the paper, and CSS gradients arrive
 * as pattern-filled paths whose color PDFium cannot voice.
 */
import { describe, expect, it } from 'vitest'
import {
  extractBackgroundPanels,
  extractPageBackground,
  extractTextBackdrops,
  normalizeShapes,
} from '../src/analyze'
import { classifyFloatImages } from '../src/analyze/floats'
import { extractPage, withPdfDocument } from '../src/extract'
import { coversBox, printContentBox } from '../src/geometry'
import type { RawPath, RawSubpath } from '../src/ir'
import { buildGradientCardPdf, buildPrintMarginWashPdf } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

const rect = (x0: number, y0: number, x1: number, y1: number): RawSubpath => ({
  points: [
    { x: x0, y: y0 },
    { x: x1, y: y0 },
    { x: x1, y: y1 },
    { x: x0, y: y1 },
  ],
  closed: true,
  hasCurves: false,
})
const fill = (x0: number, y0: number, x1: number, y1: number, fillColor: string): RawPath => ({
  subpaths: [rect(x0, y0, x1, y1)],
  filled: true,
  stroked: false,
  fillColor,
  strokeColor: '000000',
  strokeWidth: 1,
})

describe('printContentBox', () => {
  const W = 595
  const H = 842
  const box = { x0: 43, y0: 43, x1: 552, y1: 799 }
  const ink = [box, { x0: 70, y0: 600, x1: 400, y1: 720 }, { x0: 80, y0: 420, x1: 380, y1: 540 }]

  it('accepts a centered inset fill that frames the page ink', () => {
    expect(printContentBox(box, ink, W, H)).toEqual(box)
  })

  it('tolerates a stray decoration outside the box', () => {
    const many = [
      ...ink,
      ...Array.from({ length: 40 }, (_, i) => ({
        x0: 60,
        y0: 100 + i * 10,
        x1: 300,
        y1: 105 + i * 10,
      })),
    ]
    expect(printContentBox(box, [...many, { x0: 0, y0: 0, x1: 20, y1: 20 }], W, H)).toEqual(box)
  })

  it('rejects unequal margins, full-bleed fills and boxes that leave ink outside', () => {
    expect(printContentBox({ x0: 43, y0: 43, x1: 473, y1: 799 }, ink, W, H)).toBeNull()
    expect(printContentBox({ x0: 0, y0: 0, x1: W, y1: H }, ink, W, H)).toBeNull()
    expect(printContentBox(box, [...ink, { x0: 10, y0: 10, x1: 590, y1: 30 }], W, H)).toBeNull()
  })

  it('coversBox measures against the given box', () => {
    expect(coversBox({ x0: 43, y0: 43, x1: 552, y1: 799 }, box, 0.94)).toBe(true)
    expect(
      coversBox({ x0: 43, y0: 43, x1: 552, y1: 799 }, { x0: 0, y0: 0, x1: W, y1: H }, 0.94),
    ).toBe(false)
  })
})

describe('analyze: content-box washes and panels', () => {
  const W = 595
  const H = 842
  const content = { x0: 43, y0: 43, x1: 552, y1: 799 }

  it('extractPageBackground pulls a content-box wash out of the pool', () => {
    const shapes = normalizeShapes([
      fill(43, 43, 552, 799, '0A0A0A'),
      fill(88, 43, 508, 677, 'E8E0C8'),
    ])
    expect(extractPageBackground(shapes.fills, W, H)).toBeUndefined()
    expect(shapes.fills).toHaveLength(2)
    expect(extractPageBackground(shapes.fills, W, H, content)).toBe('0A0A0A')
    expect(shapes.fills.map((f) => f.color)).toEqual(['E8E0C8'])
  })

  it('extractBackgroundPanels takes a spine flush with the content-box edge', () => {
    const shapes = normalizeShapes([fill(43, 43, 221, 799, '4A5578')])
    const chars = Array.from({ length: 6 }, (_, i) => ({
      x0: 60,
      y0: 700 - i * 20,
      x1: 200,
      y1: 712 - i * 20,
    }))
    const kept = normalizeShapes([fill(43, 43, 221, 799, '4A5578')])
    expect(extractBackgroundPanels(kept.fills, chars, W, H)).toHaveLength(0)
    const panels = extractBackgroundPanels(shapes.fills, chars, W, H, content)
    expect(panels).toHaveLength(1)
    expect(shapes.fills).toHaveLength(0)
  })
})

describe('analyze: plates hosting lighter cards (P35)', () => {
  const W = 595
  const H = 842
  const chars = (x0: number, y0: number, n: number, color: string) =>
    Array.from({ length: n }, (_, i) => ({
      box: { x0: x0 + i * 6, y0, x1: x0 + i * 6 + 5, y1: y0 + 8 },
      code: 0x41,
      color,
    }))

  it('keeps a dark plate whose only text sits on a lighter sub-card', () => {
    const shapes = normalizeShapes([
      fill(64, 448, 533, 800, '161312'),
      fill(322, 507, 486, 718, 'F7F3EC'),
    ])
    // dark text lives on the light card: plate vs text contrast is nil
    const dark = chars(340, 600, 9, '161312')
    const backdrops = extractTextBackdrops(shapes.fills, dark, [], W, H)
    expect(backdrops.map((f) => f.color).sort()).toEqual(['161312', 'F7F3EC'])
    expect(shapes.fills).toHaveLength(0)
  })

  it('a dark plate with nothing lighter on it and dark text stays out', () => {
    const shapes = normalizeShapes([fill(64, 448, 533, 800, '161312')])
    const backdrops = extractTextBackdrops(shapes.fills, chars(100, 600, 9, '161312'), [], W, H)
    expect(backdrops).toHaveLength(0)
    expect(shapes.fills).toHaveLength(1)
  })
})

describe('analyze: synthetic pattern images float behind (P35)', () => {
  it('a title-only gradient card never flows inline', () => {
    const card = {
      kind: 'image' as const,
      box: { x0: 100, y0: 400, x1: 400, y1: 600 },
      data: new Uint8Array(0),
      mime: 'image/png' as const,
      pixelWidth: 300,
      pixelHeight: 200,
      synthetic: true as const,
    }
    const { floats, inline } = classifyFloatImages([card], [], 612 * 792, {
      x0: 0,
      y0: 0,
      x1: 612,
      y1: 792,
    })
    expect(inline).toHaveLength(0)
    expect(floats).toHaveLength(1)
    expect(floats[0]!.float?.wrap).toBe('behind')
  })
})

describe('extract: print-margin pages (P35)', () => {
  it('a dark content-box wash becomes the page background render', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildPrintMarginWashPdf(), (doc) =>
      extractPage(m, doc, 0),
    )
    expect(page.scanned).toBe(false)
    expect(page.degraded).toBe(false)
    expect(page.contentBox).toBeDefined()
    expect(page.contentBox!.x0).toBeCloseTo(43, 0)
    expect(page.contentBox!.x1).toBeCloseTo(552, 0)
    expect(page.bgRender).toBeDefined()
    // the wash rides the background bitmap, the light card stays a fill
    const big = page.paths.filter((p) => p.filled && p.fillColor === '0A0A0A')
    expect(big).toHaveLength(0)
    expect(page.paths.some((p) => p.filled && p.fillColor === 'F0E6D1')).toBe(true)
    expect(page.chars.length).toBeGreaterThan(20)
  })

  it('an off-center inset fill is a plate, not the page wash', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildPrintMarginWashPdf(true), (doc) =>
      extractPage(m, doc, 0),
    )
    expect(page.contentBox).toBeUndefined()
    expect(page.bgRender).toBeUndefined()
    expect(page.paths.some((p) => p.filled && p.fillColor === '0A0A0A')).toBe(true)
  })

  it('a pattern-filled path rasterizes into the image stream; a real white card stays a path', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildGradientCardPdf(), (doc) => extractPage(m, doc, 0))
    expect(page.images).toHaveLength(1)
    const img = page.images[0]!
    expect(img.box.x0).toBeCloseTo(100, 0)
    expect(img.box.x1).toBeCloseTo(400, 0)
    expect(img.box.y0).toBeCloseTo(400, 0)
    expect(img.box.y1).toBeCloseTo(600, 0)
    expect(img.mime).toBe('image/png')
    expect(img.z).toBeDefined()
    const whites = page.paths.filter((p) => p.filled && p.fillColor === 'FFFFFF')
    expect(whites).toHaveLength(1)
    const ys = whites[0]!.subpaths[0]!.points.map((pt) => pt.y)
    expect(Math.max(...ys)).toBeCloseTo(250, 0)
  })
})

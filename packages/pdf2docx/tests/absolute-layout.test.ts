/**
 * Absolute-layout mode (pptx exporter): flow-only signals never cost a slide
 * its editability — overlapping blocks do not lower page confidence, weak
 * borderless tables dissolve into positioned text, and a graphics-lost page
 * keeps its text over a text-free render instead of becoming one bitmap.
 */
import { describe, expect, it } from 'vitest'
import { openPptx } from '../../pptx-engine/src/index'
import { convertPdfToDocx, convertPdfToPptx } from '../src'
import { analyzePage } from '../src/analyze'
import type { ExtractedPage } from '../src/extract'
import type { PdfChar } from '../src/ir'
import { extractIrDocument } from '../src/pipeline'
import { mkText } from './helpers/chars'
import { buildSlideDecorPdf } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

function slidePage(chars: PdfChar[]): ExtractedPage {
  return {
    index: 0,
    widthPt: 720,
    heightPt: 405,
    rotation: 0,
    chars,
    images: [],
    paths: [],
    degraded: false,
    scanned: false,
    hasStructTree: false,
    vectorRegions: [],
    badUnicodeRatio: 0,
  }
}

/** headings whose boxes dip into the body line below them (negative gaps) */
function overlappingChars(): PdfChar[] {
  return [
    ...mkText('Big Title Here', 60, { y: 300, fontSize: 30 }).chars,
    ...mkText('Body line overlapping the title box above it', 60, { y: 292 }).chars,
    ...mkText('Second body line', 60, { y: 278 }).chars,
    ...mkText('Another Heading', 60, { y: 200, fontSize: 24 }).chars,
    ...mkText('More body overlapping again', 60, { y: 194 }).chars,
  ]
}

/** 3×2 aligned label grid with no rulings or banding (confidence 0.6) */
function weakGridChars(): PdfChar[] {
  const rows: Array<[string, string, number]> = [
    ['Alpha', 'One', 300],
    ['Beta', 'Two', 280],
    ['Gamma', 'Three', 260],
  ]
  return rows.flatMap(([a, b, y]) => [
    ...mkText(a, 100, { y }).chars,
    ...mkText(b, 250, { y }).chars,
  ])
}

describe('analyzePage absoluteLayout', () => {
  it('overlapping-block warnings stay reported but no longer lower confidence', () => {
    const flow = analyzePage(slidePage(overlappingChars()))
    const abs = analyzePage(slidePage(overlappingChars()), { absoluteLayout: true })
    const overlaps = (w: string[] | undefined) =>
      (w ?? []).filter((x) => x.startsWith('overlapping blocks')).length
    expect(overlaps(flow.warnings)).toBeGreaterThan(0)
    expect(overlaps(abs.warnings)).toBe(overlaps(flow.warnings))
    expect(flow.confidence).toBeLessThan(1)
    expect(abs.confidence).toBe(1)
  })

  it('an evidence-free borderless table dissolves into positioned text blocks', () => {
    const flow = analyzePage(slidePage(weakGridChars()))
    const table = flow.blocks.find((b) => b.kind === 'table')
    expect(table?.kind === 'table' ? table.confidence : undefined).toBeLessThanOrEqual(0.7)

    const abs = analyzePage(slidePage(weakGridChars()), { absoluteLayout: true })
    expect(abs.blocks.some((b) => b.kind === 'table')).toBe(false)
    const texts = abs.blocks.flatMap((b) =>
      b.kind === 'text' ? b.lines.flatMap((l) => l.spans.map((s) => s.text)) : [],
    )
    for (const word of ['Alpha', 'One', 'Beta', 'Two', 'Gamma', 'Three']) {
      expect(texts.some((t) => t.includes(word))).toBe(true)
    }
    expect(abs.confidence).toBe(1)
  })
})

describe('graphics-lost page in absolute layout', () => {
  it('docx still ships the bitmap; pptx keeps editable text over a text-free render', async () => {
    const pdfium = await loadPdfium()
    const pdf = await buildSlideDecorPdf()

    const docx = await convertPdfToDocx(pdf, { pdfium })
    expect(docx.pageResults[0]).toMatchObject({ status: 'degraded', reason: 'graphics-lost' })

    const { irPages, pageResults, warnings } = extractIrDocument(pdf, {
      pdfium,
      absoluteLayout: true,
    })
    expect(pageResults[0]).toMatchObject({ page: 1, status: 'ok' })
    const page = irPages[0]!
    expect(page.bgRender).toBeDefined()
    expect(page.bgColor).toBeUndefined()
    expect(page.blocks.length).toBeGreaterThan(0)
    expect(page.blocks.every((b) => b.kind === 'text')).toBe(true)
    expect(warnings.some((w) => /painted as one image behind the text/.test(w))).toBe(true)

    const result = await convertPdfToPptx(pdf, { pdfium })
    expect(result.pageResults[0]!.status).toBe('ok')
    const opened = await openPptx(result.pptx)
    const elements = opened.deck.slides[0]!.elements
    expect(elements.filter((e) => e.type === 'picture').length).toBe(1)
    expect(elements.filter((e) => e.type === 'text').length).toBeGreaterThanOrEqual(2)
  })
})

describe('absolute layout keeps same-row units apart', () => {
  /** a prose line on the left and a diagram label far to its right, one baseline */
  function proseAndLabel(): PdfChar[] {
    return [
      ...mkText('Body text of the left column runs here', 40, { y: 300 }).chars,
      ...mkText('Body text continues on a second line', 40, { y: 286 }).chars,
      ...mkText('Label', 600, { y: 300 }).chars,
    ]
  }

  it('flow mode joins the row, absolute mode gives the label its own block', () => {
    const flow = analyzePage(slidePage(proseAndLabel()))
    const flowTexts = flow.blocks.filter((b) => b.kind === 'text')
    const joined = flowTexts.some(
      (b) =>
        b.kind === 'text' &&
        b.lines.some((l) =>
          l.spans
            .map((s) => s.text)
            .join('')
            .includes('Label'),
        ) &&
        b.box.x0 < 100,
    )
    expect(joined).toBe(true)

    const abs = analyzePage(slidePage(proseAndLabel()), { absoluteLayout: true })
    const label = abs.blocks.find(
      (b) =>
        b.kind === 'text' && b.lines.some((l) => l.spans.map((s) => s.text).join('') === 'Label'),
    )
    expect(label).toBeDefined()
    expect(label!.box.x0).toBeGreaterThan(500)
    const prose = abs.blocks.find(
      (b) =>
        b.kind === 'text' &&
        b.lines.some((l) =>
          l.spans
            .map((s) => s.text)
            .join('')
            .startsWith('Body text of'),
        ),
    )
    expect(prose?.kind === 'text' ? prose.lines.length : 0).toBe(2)
  })

  it('a full-width title above prose and label does not pull them together', () => {
    const chars = [
      ...mkText('A title that spans the whole slide width from the prose to the label', 40, {
        y: 330,
        fontSize: 16,
      }).chars,
      ...proseAndLabel(),
    ]
    const abs = analyzePage(slidePage(chars), { absoluteLayout: true })
    const label = abs.blocks.find(
      (b) =>
        b.kind === 'text' && b.lines.some((l) => l.spans.map((s) => s.text).join('') === 'Label'),
    )
    expect(label).toBeDefined()
    expect(label!.box.x0).toBeGreaterThan(500)
    expect(label?.kind === 'text' ? label.lines.length : 0).toBe(1)
  })

  it('a bullet glyph a word away from its text still shares the line', () => {
    const chars = [
      ...mkText('•', 40, { y: 300 }).chars,
      ...mkText('First point', 52, { y: 300 }).chars,
    ]
    const abs = analyzePage(slidePage(chars), { absoluteLayout: true })
    const texts = abs.blocks.filter((b) => b.kind === 'text')
    expect(texts).toHaveLength(1)
  })
})

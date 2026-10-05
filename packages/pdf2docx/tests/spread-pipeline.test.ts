/**
 * Page-scale spread regressions (PR #1433 review round 2): Math.min(...xs)
 * passes every element as a function argument and throws RangeError past V8's
 * argument limit, and a 125k-line page (CAD/map exports reach that) is enough.
 *
 * The first fix covered only bodyContextOf, but the page flows through
 * EARLIER spread sites first: classifyFloatImages (analyze/floats.ts) reduces
 * all page units before bodyContextOf ever runs, and the section/column pass
 * (analyze/columns.ts) reduces page-sized element arrays on the way down.
 *
 * A full analyzePage run at 125k+ baselines is impractical in CI for a reason
 * UNRELATED to spreads — the pre-existing O(slabs × elements) cost of the
 * column pass (slabsOf) measures ~30s+ at 130k units — so per the review
 * fallback these tests drive each converted page-level site at/above the
 * argument limit through the real exported entries, plus the
 * earliest-crashing stage (classifyFloatImages) directly.
 */
import { describe, expect, it } from 'vitest'
import { analyzePage, bodyContextOf, groupIntoBlocks } from '../src/analyze'
import { detectSections, type SectionElement } from '../src/analyze/columns'
import { analyzeChars } from '../src/analyze/chars'
import { classifyFloatImages, suppressTextShadowImages } from '../src/analyze/floats'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { splitIntoUnits } from '../src/analyze/units'
import type { ExtractedPage } from '../src/extract'
import { maxOf, minOf, type Rect } from '../src/geometry'
import type { ImageBlock, PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

/** page tall enough to hold every line in-bounds (CAD/map poster pages) */
const PAGE_W = 612
const LINE_GAP = 14
const pageH = (lines: number): number => lines * LINE_GAP + 200

/** one visible char per line, top to bottom, all inside the page box */
function densePageChars(lines: number, text = 'x'): PdfChar[] {
  const topY = pageH(lines) - 100
  const chars: PdfChar[] = new Array(lines * text.length)
  let at = 0
  for (let i = 0; i < lines; i++) {
    for (const c of text) chars[at++] = mkChar(c, 72, { y: topY - i * LINE_GAP })
  }
  return chars
}

function unitsOf(chars: PdfChar[]): ReturnType<typeof splitIntoUnits> {
  return splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))
}

function extractedPage(chars: PdfChar[], heightPt: number): ExtractedPage {
  return {
    index: 0,
    widthPt: PAGE_W,
    heightPt,
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

const pageBoxOf = (lines: number): Rect => ({ x0: 0, y0: 0, x1: PAGE_W, y1: pageH(lines) })

describe('minOf/maxOf: Math.min/max semantics without the argument limit', () => {
  it('matches Math.min/Math.max on finite input, including the empty case', () => {
    const samples = [
      [3, 1, 2],
      [-5, -1, -3],
      [7, 7, 7],
      [0.5, -0.5, 0],
      [1e308, -1e308, 42],
    ]
    for (const xs of samples) {
      expect(minOf(xs)).toBe(Math.min(...xs))
      expect(maxOf(xs)).toBe(Math.max(...xs))
    }
    expect(minOf([])).toBe(Infinity)
    expect(maxOf([])).toBe(-Infinity)
  })

  it('reduces a 200k-element array that Math.min(...xs) cannot', () => {
    const xs = Array.from({ length: 200_000 }, (_, i) => (i % 2 === 0 ? i : -i))
    expect(minOf(xs)).toBe(-199_999)
    expect(maxOf(xs)).toBe(199_998)
  })
})

describe('earliest-crashing stage: classifyFloatImages page-unit bounds', () => {
  it('reduces 200k page units with loop helpers, not spreads', () => {
    const units = unitsOf(densePageChars(200_000))
    expect(units).toHaveLength(200_000)
    // even with no images at all, the body-extent reduction runs over every
    // unit — Math.min(...units.map(...)) threw past the argument limit here
    // before bodyContextOf was ever reached
    const pageBox = pageBoxOf(200_000)
    expect(() => classifyFloatImages([], units, PAGE_W * pageBox.y1, pageBox)).not.toThrow()
  })

  it('suppressTextShadowImages unions 200k chars sitting inside one image', () => {
    const units = unitsOf(densePageChars(200_000))
    const img: ImageBlock = {
      kind: 'image',
      box: pageBoxOf(200_000),
      data: new Uint8Array(0),
      mime: 'image/png',
      pixelWidth: 1,
      pixelHeight: 1,
    }
    let kept: ImageBlock[] = []
    expect(() => {
      kept = suppressTextShadowImages([img], units)
    }).not.toThrow()
    expect(kept).toHaveLength(1)
  })

  it('empty unit list keeps the 0-edge behavior', () => {
    const pageBox = pageBoxOf(1)
    expect(() => classifyFloatImages([], [], PAGE_W * 792, pageBox)).not.toThrow()
  })
})

describe('column/section pass: page-sized element bounds', () => {
  it('detectSections reduces 250k elements (header/footer peel, column extents)', () => {
    // two side-by-side column populations sharing ONE baseline band: the peel
    // pass reduces the whole page body, then the column solve reduces each
    // 125k-element column — all Math.min/max spreads before the fix
    const elements: SectionElement[] = new Array(250_000)
    for (let i = 0; i < 250_000; i++) {
      const left = i < 125_000
      elements[i] = {
        box: {
          x0: left ? 0 : 200,
          x1: left ? 100 : 300,
          y0: 400,
          y1: 600,
        },
      }
    }
    const sections = detectSections(elements, 1000)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(2)
    expect(sections[0]!.columns[0]!.elements).toHaveLength(125_000)
    expect(sections[0]!.columns[1]!.elements).toHaveLength(125_000)
  })
})

describe('block pass: page-sized group/run bounds', () => {
  it('groupIntoBlocks marks breaks across a 200k-line page (groupLeft, verse maxRight)', () => {
    // uniform 14pt gaps → one paragraph group of 200k lines; every line ends
    // in EOL punctuation so the verse scan reduces the full 200k run
    const lines = analyzeChars(densePageChars(200_000, 'x.'))
    expect(lines).toHaveLength(200_000)
    let blocks: ReturnType<typeof groupIntoBlocks> = []
    expect(() => {
      blocks = groupIntoBlocks(lines)
    }).not.toThrow()
    const blockLines = blocks.flatMap((b) => b.lines)
    expect(blockLines.length).toBeGreaterThan(190_000)
  })

  it('bodyContextOf reduces a 200k-line page', () => {
    const lines = analyzeChars(densePageChars(200_000))
    const ctx = bodyContextOf(lines)
    expect(ctx.bodyLeft).toBe(72)
    expect(ctx.bodyRight).toBeGreaterThan(72)
  })
})

describe('analyzePage: extracted page → analyze sanity', () => {
  it('keeps small-page results identical', () => {
    const chars = [
      ...mkText('The quick brown fox', 72, { y: 700 }).chars,
      ...mkText('jumps over the lazy dog', 72, { y: 686 }).chars,
      ...mkText('and keeps flowing on.', 72, { y: 672 }).chars,
    ]
    const page = analyzePage(extractedPage(chars, 792))
    const text = page.blocks
      .flatMap((b) => (b.kind === 'text' ? b.lines : []))
      .flatMap((l) => l.spans.map((s) => s.text))
      .join(' ')
    expect(text).toContain('quick brown fox')
    expect(text).toContain('keeps flowing')
  })
})

import { describe, expect, it } from 'vitest'
import {
  buildPrintDocumentHtml,
  currentRangeIndices,
  normalizePrintRatio,
  parsePrintRange,
  printPageSizeIn,
  printPageCount,
} from '../src/shared/print-html'

const srcs = (n: number) => Array.from({ length: n }, (_x, i) => `blob:img-${i}`)

describe('parsePrintRange', () => {
  it('parses singles, ranges and mixed separators, deduped and sorted', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
    expect(parsePrintRange('3，1、2 5;4', 5)).toEqual([0, 1, 2, 3, 4])
    expect(parsePrintRange('2-4, 3', 5)).toEqual([1, 2, 3])
  })

  it('rejects empty, malformed, out-of-bounds and inverted input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('a-b', 5)).toBeNull()
    expect(parsePrintRange('0-2', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
  })

  it('accepts spaces around dashes', () => {
    expect(parsePrintRange('1 - 3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 – 3, 5', 5)).toEqual([0, 1, 2, 4])
  })

  it('accepts em-dash, tilde, fullwidth tilde and wave dash (IME variants)', () => {
    expect(parsePrintRange('1—3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1~3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1～3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1〜3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 — 3, 5～6', 6)).toEqual([0, 1, 2, 4, 5])
  })

  it('accepts spaces around dashes', () => {
    expect(parsePrintRange('1 - 3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 – 3, 5', 5)).toEqual([0, 1, 2, 4])
  })
})

describe('printPageCount', () => {
  it('follows the per-page thumbnail count of each layout', () => {
    expect(printPageCount(7, 'full')).toBe(7)
    expect(printPageCount(7, 'notes')).toBe(7)
    expect(printPageCount(7, 'handout2')).toBe(4)
    expect(printPageCount(7, 'handout3')).toBe(3)
    expect(printPageCount(7, 'handout4')).toBe(2)
    expect(printPageCount(7, 'handout6')).toBe(2)
    expect(printPageCount(10, 'handout9')).toBe(2)
  })
})

describe('currentRangeIndices', () => {
  it('yields [] for stale indices instead of crashing', () => {
    expect(currentRangeIndices(2, 5)).toEqual([2])
    expect(currentRangeIndices(7, 5)).toEqual([])
    expect(currentRangeIndices(-1, 5)).toEqual([])
    expect(currentRangeIndices(1.5, 5)).toEqual([])
    expect(currentRangeIndices(NaN, 5)).toEqual([])
  })
})

describe('buildPrintDocumentHtml', () => {
  it('full layout sizes pages by the slide ratio and can frame slides', () => {
    const html = buildPrintDocumentHtml({ srcs: srcs(2), ratio: 16 / 9, layout: 'full' })
    expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(html.match(/class="page"/g)).toHaveLength(2)
    expect(html).not.toContain('.page > img { border')

    const framed = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'full',
      frame: true,
    })
    expect(framed).toContain('.page > img { border: 1px solid #bbb; }')
  })

  it('handouts use A4 and swap dimensions in landscape', () => {
    const portrait = buildPrintDocumentHtml({ srcs: srcs(5), ratio: 16 / 9, layout: 'handout6' })
    expect(portrait).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
    expect(portrait).toContain('grid-template-columns: 1fr 1fr;')

    const landscape = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout6',
      orientation: 'landscape',
    })
    expect(landscape).toContain('@page { size: 11.69in 8.27in; margin: 0; }')
    expect(landscape).toContain('grid-template-columns: 1fr 1fr 1fr;')
  })

  it('handout4 lays a 2×2 grid and handout9 a 3×3 grid', () => {
    const four = buildPrintDocumentHtml({ srcs: srcs(5), ratio: 16 / 9, layout: 'handout4' })
    expect(four).toContain(
      '.page.handout.h4 { display: grid; grid-auto-flow: row; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; }',
    )
    expect(four.match(/class="page handout h4"/g)).toHaveLength(2)
    expect(four.match(/class="cell"/g)).toHaveLength(5)

    const nine = buildPrintDocumentHtml({ srcs: srcs(10), ratio: 16 / 9, layout: 'handout9' })
    expect(nine).toContain(
      '.page.handout.h9 { display: grid; grid-auto-flow: row; grid-template-columns: 1fr 1fr 1fr; grid-template-rows: 1fr 1fr 1fr; }',
    )
    expect(nine.match(/class="page handout h9"/g)).toHaveLength(2)
    expect(nine.match(/class="cell"/g)).toHaveLength(10)
  })

  it('vertical order flows the grid down each column and stacks the landscape pair', () => {
    const sixVertical = buildPrintDocumentHtml({
      srcs: srcs(6),
      ratio: 16 / 9,
      layout: 'handout6',
      order: 'vertical',
    })
    expect(sixVertical).toContain('grid-auto-flow: column;')
    expect(sixVertical).toContain('grid-template-columns: 1fr 1fr;')

    const pairHorizontal = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 16 / 9,
      layout: 'handout2',
      orientation: 'landscape',
    })
    expect(pairHorizontal).toContain('grid-template-columns: 1fr 1fr; grid-template-rows: 1fr; }')
    const pairVertical = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 16 / 9,
      layout: 'handout2',
      orientation: 'landscape',
      order: 'vertical',
    })
    expect(pairVertical).toContain('grid-template-columns: 1fr; grid-template-rows: 1fr 1fr; }')
  })

  it('handout3 keeps its flex column of slide + ruled lines in both orientations', () => {
    for (const orientation of [undefined, 'landscape'] as const) {
      const html = buildPrintDocumentHtml({
        srcs: srcs(4),
        ratio: 16 / 9,
        layout: 'handout3',
        ...(orientation ? { orientation } : {}),
      })
      expect(html).toContain('.page.handout.h3 { display: flex; flex-direction: column; }')
      expect(html.match(/class="rules"/g)).toHaveLength(4)
      expect(html).not.toContain('grid-auto-flow')
    }
  })

  it('notes layout escapes note text and pairs it with each slide', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 4 / 3,
      layout: 'notes',
      notes: ['a < b\nnext', ''],
    })
    expect(html.match(/class="page notes"/g)).toHaveLength(2)
    expect(html).toContain('a &lt; b<br>next')
  })

  it('escapes image sources for attribute context in every layout', () => {
    const full = buildPrintDocumentHtml({
      srcs: ['x" onerror="alert(1)'],
      ratio: 16 / 9,
      layout: 'full',
    })
    expect(full).toContain('src="x&quot; onerror=&quot;alert(1)"')
    expect(full).not.toContain('src="x" onerror=')
    const handout = buildPrintDocumentHtml({
      srcs: ["c'd", 'e&f<g>'],
      ratio: 16 / 9,
      layout: 'handout2',
    })
    expect(handout).toContain('src="c&#39;d"')
    expect(handout).toContain('src="e&amp;f&lt;g&gt;"')
  })

  it('preview mode adds page badges with the total page count', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout2',
      preview: true,
    })
    expect(html).toContain("content: counter(pg) ' / 3';")
    expect(html).toContain('counter-reset: pg;')
  })
})

describe('normalizePrintRatio', () => {
  it('keeps real slide formats exact', () => {
    expect(normalizePrintRatio(16 / 9)).toBe(16 / 9)
    expect(normalizePrintRatio(4 / 3)).toBe(4 / 3)
    expect(normalizePrintRatio(16 / 10)).toBe(16 / 10)
  })

  it('falls back to 16:9 for non-finite and non-positive ratios', () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, 0, -1.5]) {
      expect(normalizePrintRatio(bad)).toBe(16 / 9)
      const html = buildPrintDocumentHtml({ srcs: srcs(1), ratio: bad, layout: 'full' })
      expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
      expect(html).not.toContain('NaNin')
      expect(html).not.toContain('Infinityin')
    }
  })

  it('clamps absurd ratios so @page stays printable', () => {
    const tiny = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 0.01, layout: 'full' })
    expect(tiny).toContain('@page { size: 1.5in 7.5in; margin: 0; }')
    const huge = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 100, layout: 'full' })
    expect(huge).toContain('@page { size: 37.5in 7.5in; margin: 0; }')
  })
})

describe('printPageSizeIn', () => {
  it('the full layout follows the slide ratio at 7.5in tall', () => {
    expect(printPageSizeIn('full', 'portrait', 16 / 9)).toEqual({ width: 13.333, height: 7.5 })
    expect(printPageSizeIn('full', 'landscape', 16 / 9)).toEqual({ width: 13.333, height: 7.5 })
  })

  it('handout and notes sheets are A4 in the chosen orientation', () => {
    for (const layout of [
      'handout2',
      'handout3',
      'handout4',
      'handout6',
      'handout9',
      'notes',
    ] as const) {
      expect(printPageSizeIn(layout, 'portrait', 16 / 9)).toEqual({ width: 8.27, height: 11.69 })
      expect(printPageSizeIn(layout, 'landscape', 16 / 9)).toEqual({ width: 11.69, height: 8.27 })
    }
  })

  it('a corrupt ratio clamps like the page rule does (same source of truth)', () => {
    expect(printPageSizeIn('full', 'portrait', NaN)).toEqual({ width: 13.333, height: 7.5 })
    expect(printPageSizeIn('full', 'portrait', 100)).toEqual({ width: 37.5, height: 7.5 })
    // and the HTML builder agrees with the helper byte for byte
    const html = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 16 / 9, layout: 'full' })
    expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
  })
})

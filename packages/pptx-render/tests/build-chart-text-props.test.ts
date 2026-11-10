import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@genoffice/pptx-engine'
import { buildChartNode, truncateToWidth, rotatedLabelOrigin, applyCaps } from '../src/build-chart'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = {
  x: 0,
  y: 0,
  w: 600,
  h: 400,
  centerX: 300,
  centerY: 200,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
}
const bars = (extra: Partial<ChartModel> = {}): ChartModel => ({
  kind: 'bar',
  barDir: 'col',
  grouping: 'clustered',
  categories: ['T1', 'T2', 'Total year'],
  series: [{ name: 'Year 2025', values: [17714, 16524, 173021], plotKind: 'bar' }],
  dataLabels: true,
  ...extra,
})
const build = (m: ChartModel) => buildChartNode('r', 'e', m, box, vp, metrics)!

describe('data labels: rotation, caps, letter spacing, General format', () => {
  it('prints General values without a thousands separator', () => {
    const texts = build(bars()).labels.map((l) => l.text)
    expect(texts).toContain('173021')
    expect(texts).not.toContain('173,021')
  })

  it('bodyPr rot=-90 draws the label bottom-to-top centered on the bar', () => {
    const flat = build(bars()).labels.find((l) => l.text === '173021')!
    const node = build(bars({ dataLabelRotDeg: -90 }))
    const bar = node.bars[2]!
    const lbl = node.labels.find((l) => l.text === '173021')!
    expect(lbl.rotationDeg).toBe(-90)
    expect(flat.rotationDeg).toBeUndefined()
    // glyph box (1.1em to the right of the origin) straddles the bar center
    expect(lbl.x).toBeLessThan(bar.x + bar.w / 2)
    expect(lbl.x + lbl.fontSizePx * 1.1).toBeGreaterThan(bar.x + bar.w / 2)
    // the text starts just above the bar top and runs upward
    expect(lbl.y).toBeLessThan(bar.y)
    expect(lbl.y).toBeGreaterThan(bar.y - lbl.fontSizePx)
  })

  it('inside labels rotate about the segment center', () => {
    const o = rotatedLabelOrigin(100, 50, 40, 10, true)
    expect(o.x).toBeCloseTo(94.5, 5)
    // horizontal center line = 50 + 5.5; the text's bottom end starts half its width below it
    expect(o.y).toBeCloseTo(75.5, 5)
    expect(rotatedLabelOrigin(100, 50, 40, 10, false).y).toBeCloseTo(60, 5)
  })

  it('cap="all" upper-cases and spc spaces the data labels and categories', () => {
    const m = bars({
      categories: ['Total year', 'Px hn', 'Px hcm'],
      dataLabelCatName: true,
      dataLabelNoValue: true,
      dataLabelCaps: true,
      dataLabelSpcPt: 1.2,
      catAxis: { labelCaps: true, labelSpcPt: 1.2 },
    })
    const node = build(m)
    const dl = node.labels.filter((l) => l.text === 'TOTAL YEAR')
    expect(dl.length).toBe(2)
    expect(dl.every((l) => (l.letterSpacingPx ?? 0) > 0)).toBe(true)
    expect(node.labels.some((l) => l.text === 'Total year')).toBe(false)
    expect(applyCaps('a\u00e0', true)).toBe('A\u00c0')
  })

  it('a filled label box spans the letter-spaced text width', () => {
    const boxed = (spc?: number) =>
      build(
        bars({
          series: [
            {
              name: 'S',
              values: [17714, 16524, 173021],
              plotKind: 'bar',
              dataLabelFill: '#FFFFFF',
            },
          ],
          ...(spc ? { dataLabelSpcPt: spc } : {}),
        }),
      ).labels.find((l) => l.text === '17714')!
    const plain = boxed()
    const spaced = boxed(1.2)
    expect(plain.w).toBeGreaterThan(0)
    expect(spaced.letterSpacingPx).toBeGreaterThan(0)
    expect(spaced.w! - plain.w!).toBeCloseTo(spaced.letterSpacingPx! * 5, 5)
  })

  it('cap/spc on the title widen it into more wrapped lines', () => {
    const title = 'Production hours per workshop september 2026 overview'
    const plain = build(bars({ title, titlePt: 21.6 }))
    const caps = build(bars({ title, titlePt: 21.6, titleCaps: true, titleSpcPt: 1.2 }))
    const titleLines = (n: typeof plain) => n.labels.filter((l) => l.fontSizePx > 25)
    expect(titleLines(caps).length).toBeGreaterThan(titleLines(plain).length)
    expect(titleLines(caps)[0]!.text).toBe(titleLines(caps)[0]!.text.toUpperCase())
    expect(titleLines(caps)[0]!.letterSpacingPx).toBeGreaterThan(0)
  })
})

describe('rotated category labels: band cap and ellipsis', () => {
  const longCats = [
    `\u7121\u6708\u7d4c`,
    `\u4e0d\u6b63\u6027\u5668\u51fa\u8840\u306e\u7cbe\u67fb`,
    `\u8996\u5e8a\u4e0b\u90e8\u6027\u7121\u6708\u7d4c\uff08\u30a2\u30b9\u30ea\u30fc\u30c8\uff09\u8ffd\u52a0\u306e\u9577\u3044\u8aac\u660e\u6587`,
    'PCOS',
    'OHSS',
    'FSH',
    `\u6392\u5375\u8a98\u767a`,
    `\u5185\u819c\u30dd\u30ea\u30fc\u30d7`,
    `\u6a5f\u80fd\u6027\u51fa\u8840`,
    `\u6708\u7d4c\u56f0\u96e3\u75c7`,
    `\u9ad8PRL`,
    `\u4f53\u91cd\u6e1b\u5c11\u6027\u7121\u6708\u7d4c`,
  ]
  it('truncateToWidth keeps a prefix plus an ellipsis', () => {
    const measure = (t: string) => Array.from(t).length * 10
    expect(truncateToWidth('abcdefgh', 1000, measure)).toBe('abcdefgh')
    expect(truncateToWidth('abcdefgh', 50, measure)).toBe('abcd\u2026')
    expect(truncateToWidth('abcdefgh', Infinity, measure)).toBe('abcdefgh')
  })

  it('caps the rotated band at 40% of the frame and truncates the overlong label', () => {
    const model = bars({
      categories: longCats,
      series: [{ name: 's', values: longCats.map((_, i) => i + 1), plotKind: 'bar' }],
      dataLabels: false,
      catAxis: { labelSizePt: 9 },
    })
    const short = { ...box, h: 300 }
    const node = buildChartNode('r', 'e', model, short, vp, metrics)!
    const rotated = node.labels.filter((l) => l.rotationDeg === -45)
    expect(rotated.length).toBe(longCats.length)
    const longest = rotated.find((l) => l.text.endsWith('\u2026'))!
    expect(longest).toBeDefined()
    expect(longest.text.startsWith(`\u8996\u5e8a\u4e0b\u90e8`)).toBe(true)
    expect(rotated.filter((l) => l.text.endsWith('\u2026')).length).toBe(1)
    // the plot keeps at least ~55% of the frame
    const axisY = Math.max(...node.axisLines.map((a) => a.y1))
    expect(axisY).toBeGreaterThan(short.h * 0.5)
    expect(axisY).toBeLessThan(short.h * 0.66)
  })
})

describe('data table header pitch', () => {
  it('wrapped header lines stack at 1.6em with ~1em padding', () => {
    const cats = [
      `\u30bd\u30d5\u30a3\u306f\u3060\u304a\u3082\u3044\u30aa\u30fc\u30ac\u30cb\u30c3\u30af\u30b3\u30c3\u30c8\u30f3`,
      `\u30ed\u30ea\u30a8\u30b9\u30ea\u30e0\u30ac\u30fc\u30c9`,
      `\u671d\u307e\u3067\u5b89\u5fc3`,
      `\u30bd\u30d5\u30a3SPORTS`,
      `\u30a8\u30ea\u30b9\u30b7\u30e7\u30fc\u30c4`,
      `\u30ed\u30ea\u30a8\u808c\u304d\u308c\u3044\u30ac\u30fc\u30c9`,
      `\u30a8\u30ea\u30b9\u65b0\u7d20\u808c\u611f`,
      `\u30bd\u30d5\u30a3\u8d85\u719f\u7761`,
    ]
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      categories: cats,
      series: [{ name: 'A', values: cats.map((_, i) => 10 + i), plotKind: 'bar' }],
      dataTable: { showKeys: true, horzBorder: true, vertBorder: true, outline: true, sizePt: 12 },
    }
    const node = buildChartNode('r', 'e', model, { ...box, w: 1186, h: 523 }, vp, metrics)!
    const size = node.labels.find((l) => l.text === 'A')!.fontSizePx
    const first = node.labels
      .filter((l) => cats[0]!.includes(l.text) && l.text !== cats[0])
      .sort((a, b) => a.y - b.y)
    expect(first.length).toBeGreaterThanOrEqual(3)
    expect(first[1]!.y - first[0]!.y).toBeCloseTo(size * 1.6, 1)
    const axisY = Math.max(...node.axisLines.map((a) => a.y1))
    const tableTop = Math.min(
      ...node.gridLines.filter((g) => g.y1 === g.y2 && g.y1 > axisY + 1).map((g) => g.y1),
    )
    expect(tableTop - axisY).toBeCloseTo(size * (1.6 * first.length + 1.05), 0)
  })
})

describe('title paragraphs and parked legends', () => {
  it('keeps each title paragraph on its own line', () => {
    const node = build(bars({ title: 'Output hours\nPer workshop', titlePt: 18 }))
    const lines = node.labels.filter((l) => l.fontSizePx > 20).map((l) => l.text)
    expect(lines).toEqual(['Output hours', 'Per workshop'])
  })

  it('a bottom legend parked at the top edge reserves no bottom row', () => {
    const base = bars({ legendPos: 'b' })
    const parked = bars({
      legendPos: 'b',
      legendLayout: { x: 0.73, y: 0.045, xMode: 'edge', yMode: 'edge' },
    })
    const axisY = (m: ChartModel) => Math.max(...build(m).axisLines.map((a) => a.y1))
    const none = axisY(bars({}))
    expect(axisY(base)).toBeLessThan(none - 10)
    expect(axisY(parked)).toBeCloseTo(none, 3)
  })
})

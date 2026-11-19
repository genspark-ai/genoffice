import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@genoffice/pptx-engine'
import { buildChartNode } from '../src/build-chart'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = {
  x: 0,
  y: 0,
  w: 520,
  h: 500,
  centerX: 260,
  centerY: 250,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
}
const texts = (node: { labels: Array<{ text: string }> }) => node.labels.map((l) => l.text)

// Synthetic Japanese product names (katakana + kanji), 9-15 characters, no spaces
const JA_CATS = [
  `\u30c6\u30b9\u30c8\u5546\u54c1\u30a2\u30eb\u30d5\u30a1`,
  `\u30b5\u30f3\u30d7\u30eb\u30d6\u30e9\u30f3\u30c9\u30d9\u30fc\u30bf`,
  `\u9577\u3044\u540d\u524d\u306e\u30ab\u30c6\u30b4\u30ea\u30fc\u9805\u76ee`,
  `\u30b3\u30f3\u30d1\u30af\u30c8\u30bf\u30a4\u30d7\uff08\u8584\u578b\uff09\uff11\uff0f\uff12`,
  `\u671d\u307e\u3067\u5b89\u5fc3\u30b7\u30ea\u30fc\u30ba`,
  `\u3057\u3063\u3068\u308a\u30ac\u30fc\u30c9\u30d7\u30e9\u30b9`,
]

function combo(extra: Partial<ChartModel> = {}): ChartModel {
  return {
    kind: 'bar',
    barDir: 'col',
    grouping: 'clustered',
    categories: JA_CATS,
    series: [
      {
        name: 'Before',
        values: [0.52, 0.5, 0.43, 0.3, 0.33, 0.41],
        plotKind: 'bar',
        valFmt: '0.0%',
      },
      {
        name: 'After',
        values: [0.59, 0.57, 0.5, 0.33, 0.36, 0.45],
        plotKind: 'bar',
        valFmt: '0.0%',
      },
      {
        name: 'Delta',
        values: [0.066, 0.068, 0.072, 0.034, 0.03, 0.037],
        plotKind: 'line',
        secondaryAxis: true,
        valFmt: '0.0%',
      },
    ],
    valAxis: { numFmt: '0.0%', labelSizePt: 9 },
    valAxis2: { numFmt: '0.0%', labelSizePt: 9, crosses: 'max' },
    catAxis: { labelSizePt: 9 },
    ...extra,
  }
}

describe('combo chart: secondary axis format, CJK label wrapping, data table', () => {
  it('formats the secondary value axis ticks with its own numFmt', () => {
    const node = buildChartNode('r', 'e', combo(), box, vp, metrics)!
    const ticks = texts(node).filter((t) => /^-?\d+(\.\d+)?%$/.test(t))
    expect(ticks).toContain('8.0%')
    expect(texts(node).some((t) => /^0\.0[2468]$/.test(t))).toBe(false)
  })

  it('wraps long CJK category labels per character instead of rotating them', () => {
    const node = buildChartNode('r', 'e', combo(), box, vp, metrics)!
    const rotated = node.labels.filter((l) => l.rotationDeg)
    expect(rotated).toEqual([])
    const first = JA_CATS[0]!
    // the first category's column: its pieces share the slot center of its opening line
    const head = node.labels.find((l) => first.startsWith(l.text) && l.text.length < first.length)!
    const cx =
      head.x +
      metrics.measure(head.text, {
        fontFamily: 'Calibri',
        fontSizePx: head.fontSizePx,
        bold: false,
        italic: false,
      }) /
        2
    const pieces = node.labels
      .filter(
        (l) =>
          first.includes(l.text) &&
          l.text.length < first.length &&
          Math.abs(l.x + (l.text.length * head.fontSizePx) / 2 - cx) < head.fontSizePx,
      )
      .sort((p, q) => p.y - q.y)
    expect(pieces.length).toBeGreaterThan(1)
    expect(pieces.map((p) => p.text).join('')).toBe(first)
  })

  it('draws the data table rows below the category labels and shrinks the plot', () => {
    const plain = buildChartNode('r', 'e', combo(), box, vp, metrics)!
    const model = combo({
      dataTable: { showKeys: true, horzBorder: true, vertBorder: true, outline: true, sizePt: 12 },
    })
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const axisY = (n: typeof node) => Math.max(...n.axisLines.map((a) => a.y1))
    expect(axisY(node)).toBeLessThan(axisY(plain) - 60)
    const cell = node.labels.find((l) => l.text === '52.0%')!
    const name = node.labels.find((l) => l.text === 'Before')!
    expect(cell).toBeDefined()
    expect(cell.y).toBeGreaterThan(axisY(node))
    expect(name.y).toBeCloseTo(cell.y, 5)
    expect(name.x).toBeLessThan(cell.x)
    // bar series keep a swatch key; the line series shows a line sample like the legend
    expect(node.swatches.length).toBeGreaterThanOrEqual(2)
    const delta = node.labels.find((l) => l.text === 'Delta')!
    const lineKey = node.polylines.find(
      (p) =>
        p.points.length === 4 && Math.abs(p.points[1]! - (delta.y + delta.fontSizePx * 0.6)) < 8,
    )
    expect(lineKey).toBeDefined()
    expect(lineKey!.points[2]! - lineKey!.points[0]!).toBeCloseTo(delta.fontSizePx * 1.7, 1)
    const rows = new Set(
      node.labels.filter((l) => /%$/.test(l.text) && l.y > axisY(node)).map((l) => l.y),
    )
    expect(rows.size).toBe(3)
    // header labels take the table text size, not the 9pt axis size
    const header = node.labels.find((l) => l.text === JA_CATS[0]!.slice(0, 3))
    expect(header?.fontSizePx).toBeCloseTo(cell.fontSizePx, 5)
  })
})

describe('CJK category wrapping: kinsoku and katakana words', () => {
  it('keeps katakana runs together and never opens a line with a small kana or closing bracket', () => {
    // katakana words + full-width brackets + small kana
    const label = `\u5438\u6c34\u30b7\u30e7\u30fc\u30c4\uff08\u65b0\u578b\u30b7\u30e7\u30fc\u30c4\uff09\u3084\u30ab\u30c3\u30d7\u306a\u3069\u3001\u65b0\u3057\u3044\u30bf\u30a4\u30d7`
    const model = combo({ categories: JA_CATS.map((c, i) => (i === 1 ? label : c)) })
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const lines = node.labels
      .filter((l) => label.includes(l.text) && l.text.length < label.length)
      .sort((a, b) => a.y - b.y)
      .map((l) => l.text)
    expect(lines.join('')).toBe(label)
    for (const l of lines) {
      expect(`\u30e7\u30c3\uff09\u3001`.includes(l[0]!)).toBe(false)
      expect(l.endsWith(`\uff08`)).toBe(false)
    }
    expect(lines.some((l) => l.includes(`\u30b7\u30e7\u30fc\u30c4`))).toBe(true)
  })
})

const KATAKANA_NO_START = `ーッャュョァィゥェォ）、`
const lineW = (l: { text: string; fontSizePx: number }) =>
  metrics.measure(l.text, {
    fontFamily: 'Calibri',
    fontSizePx: l.fontSizePx,
    bold: false,
    italic: false,
  })
// the label's own column: pieces whose center sits near the opening line's center
const linesOf = (node: ReturnType<typeof buildChartNode>, label: string) => {
  const pieces = node!.labels.filter((l) => label.includes(l.text) && l.text.length < label.length)
  const head = pieces.find((l) => label.startsWith(l.text))!
  const cx = head.x + lineW(head) / 2
  return pieces
    .filter((l) => Math.abs(l.x + lineW(l) / 2 - cx) < head.fontSizePx)
    .sort((a, b) => a.y - b.y)
}

describe('CJK category wrapping: overlong katakana runs', () => {
  // a katakana run wider than the slot splits per character, still under kinsoku
  const runs = [
    `スーパーウルトラマーケットチェーンストアーズ`,
    `アーカイブショップセールチャンネルマネージャー`,
    `プレミアムコーヒーメーカーキャンペーンセット`,
  ]
  it('never opens a line with a prolonged sound mark or small kana', () => {
    for (const run of runs)
      for (const nCats of [4, 6, 8]) {
        const cats = JA_CATS.slice(0, nCats - 1).concat(run)
        const node = buildChartNode('r', 'e', combo({ categories: cats }), box, vp, metrics)!
        const lines = linesOf(node, run)
        expect(lines.length, run).toBeGreaterThan(1)
        expect(lines.map((l) => l.text).join('')).toBe(run)
        for (const l of lines) expect(KATAKANA_NO_START.includes(l.text[0]!), l.text).toBe(false)
      }
  })

  it('re-checks the width after an opening bracket is carried to the next line', () => {
    const label = `\u65b0\u578b\uff08` + runs[0] + `\uff09`
    const cats = JA_CATS.slice(0, 5).concat(label)
    const node = buildChartNode('r', 'e', combo({ categories: cats }), box, vp, metrics)!
    expect(node.labels.filter((l) => l.rotationDeg)).toEqual([])
    const lines = linesOf(node, label)
    expect(lines.map((l) => l.text).join('')).toBe(label)
    const slotW = (node.gridLines.length ? (node.box?.w ?? box.w) : box.w) / cats.length
    for (const l of lines) expect(lineW(l), l.text).toBeLessThan(slotW)
    for (const l of lines) expect(l.text.endsWith(`（`), l.text).toBe(false)
  })
})

describe('data table layout edge cases', () => {
  const table = { showKeys: true, horzBorder: true, vertBorder: true, outline: true, sizePt: 12 }
  const barsOnly = (extra: Partial<ChartModel> = {}) => {
    const m = combo({ dataTable: table, ...extra })
    m.series = m.series.filter((s) => !s.secondaryAxis)
    delete m.valAxis2
    return m
  }

  it('still draws the table when the category tick labels are hidden', () => {
    const model = combo({ dataTable: table, catAxis: { labelSizePt: 9, tickLblHidden: true } })
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(texts(node)).toContain('52.0%')
    expect(texts(node)).toContain('Before')
    expect(node.swatches.length).toBeGreaterThanOrEqual(2)
    expect(texts(node).some((t) => JA_CATS[0]!.startsWith(t) && t !== JA_CATS[0])).toBe(false)
  })

  it('reserves the key column when the value axis sits on the right', () => {
    const node = buildChartNode(
      'r',
      'e',
      barsOnly({ catAxis: { labelSizePt: 9, reversed: true } }),
      box,
      vp,
      metrics,
    )!
    const name = node.labels.find((l) => l.text === 'Before')!
    const verticals = node.gridLines.filter((g) => g.x1 === g.x2).map((g) => g.x1)
    const left = Math.min(...verticals)
    const plotX = Math.min(...verticals.filter((x) => x > left + 1))
    expect(name.x).toBeGreaterThan(left)
    expect(name.x + lineW(name)).toBeLessThanOrEqual(plotX)
    const swatch = node.swatches.reduce((a, b) => (a.x < b.x ? a : b))
    expect(swatch.x).toBeGreaterThan(left)
  })

  it('places the table below the multi-level group label row', () => {
    const model = barsOnly({
      categoryGroups: [
        { label: 'Group A', start: 0 },
        { label: 'Group B', start: 3 },
      ],
    })
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const group = node.labels.find((l) => l.text === 'Group A')!
    const cell = node.labels.find((l) => l.text === '52.0%')!
    const name = node.labels.find((l) => l.text === 'Before')!
    expect(group).toBeDefined()
    expect(cell.y).toBeGreaterThanOrEqual(group.y + group.fontSizePx * 1.2)
    expect(name.y).toBeCloseTo(cell.y, 5)
    const topRule = node.gridLines
      .filter((g) => g.y1 === g.y2 && g.x1 < g.x2)
      .map((g) => g.y1)
      .sort((a, b) => a - b)
    expect(
      topRule.some(
        (y) => Math.abs(y - (cell.y - (cell.fontSizePx * 2 - cell.fontSizePx * 1.2) / 2)) < 1,
      ),
    ).toBe(true)
  })
})

describe('horizontal bar chart data table', () => {
  const hbar = (extra: Partial<ChartModel> = {}): ChartModel => ({
    kind: 'bar',
    barDir: 'bar',
    grouping: 'percentStacked',
    categories: ['Medicare', 'Commercial', 'Total'],
    series: [
      { name: 'Low', values: [24151816, 53975913, 78127729], valFmt: '#,##0' },
      { name: 'Medium', values: [15952444, 53847915, 69800359], valFmt: '#,##0' },
    ],
    valAxis: { hidden: true, numFmt: '0%' },
    catAxis: { labelSizePt: 12 },
    ...extra,
  })
  const dataTable = {
    showKeys: true,
    horzBorder: true,
    vertBorder: true,
    outline: true,
    sizePt: 12,
  }

  it('draws a category header row and series rows under the plot, in category order', () => {
    const plain = buildChartNode('r', 'e', hbar(), box, vp, metrics)!
    const node = buildChartNode('r', 'e', hbar({ dataTable }), box, vp, metrics)!
    const plotBottom = (n: typeof node) => Math.max(...n.bars.map((b) => b.y + b.h))
    expect(plotBottom(node)).toBeLessThan(plotBottom(plain) - 40)
    const header = node.labels.filter((l) => ['Medicare', 'Commercial', 'Total'].includes(l.text))
    // axis labels on the left plus one header cell each below the plot
    const cells = header.filter((l) => l.y > plotBottom(node))
    expect(cells.map((l) => l.text)).toEqual(['Medicare', 'Commercial', 'Total'])
    expect(cells[0]!.x).toBeLessThan(cells[1]!.x)
    expect(cells[1]!.x).toBeLessThan(cells[2]!.x)
    const low = node.labels.find((l) => l.text === 'Low')!
    const v = node.labels.find((l) => l.text === '24,151,816')!
    expect(low.y).toBeGreaterThan(cells[0]!.y)
    // overlong cells shrink to the slot and shift down a little
    expect(Math.abs(v.y - low.y)).toBeLessThan(10)
    expect(low.x).toBeLessThan(Math.min(...node.bars.map((b) => b.x)))
    expect(node.swatches.length).toBeGreaterThanOrEqual(2)
    expect(node.gridLines.length).toBeGreaterThan(plain.gridLines.length)
  })
})

describe('charts without a category cache', () => {
  it('label the slots 1..N like PowerPoint (axis labels and the data table header)', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      categories: ['', '', ''],
      series: [{ values: [20, 27, 12] }],
      dataTable: { showKeys: true, horzBorder: true, vertBorder: true, outline: true },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const texts = node.labels.map((l) => l.text)
    expect(texts).toEqual(expect.arrayContaining(['1', '2', '3', 'Series1']))
  })
})

/**
 * Phase tests of the table column-width algorithm (apps/docs/src/renderer/table-layout/).
 * Fixtures follow the 2026-11-08 regression cases: 052 uniform placeholder tcW, 009 slash token,
 * 035 intercorrencias, 056 noWrap pins. A fixed-advance TextMeasure makes the twips exact;
 * the tokenizer rules are checked through the heuristic measure.
 */
import { describe, expect, it } from 'vitest'
import type { TableModel } from '@genoffice/docx-engine'
import {
  clampColumnWidths,
  distributeContentAutofit,
  expandAutofitColWidths,
  layoutTableColumns,
  minContentColTwips,
  reclaimSurplus,
  scaleToFit,
  textMeasureOf,
  type TextMeasure,
} from '../src/renderer/table-layout'
import { resolveTableWidth } from '../src/renderer/editor/table-geometry'

/** 10 px per character, tokens split at spaces only: real advances, no slack, px-grid aligned */
const tenPx: TextMeasure = {
  wordPx: (runs) =>
    Math.max(
      0,
      ...runs
        .map((r) => r.text)
        .join('')
        .split(' ')
        .map((w) => w.length * 10),
    ),
  linePx: (runs) => runs.reduce((sum, r) => sum + r.text.length * 10, 0),
  measured: true,
}
const heuristic = textMeasureOf(null)

const cell = (text: string) => ({ paras: [text] })
const LETTER = { avail: 12240 - 1440, fit: 9360 }

/** python-docx placeholder: tblW auto, n equal gridCols, every tcW = gridCol (parse keeps them) */
function uniform(rows: string[][], colTwips: number): TableModel {
  return {
    rows: rows.map((r) => r.map(cell)),
    colWidthsTwips: rows[0].map(() => colTwips),
    colWidthsPct: rows[0].map(() => 100 / rows[0].length),
    autoLayout: true,
    autoFit: 'contents',
  }
}

describe('layoutTableColumns phases', () => {
  it('052: a uniform placeholder grid whose words fit is kept (Word keeps uniform tcW, probe 2026-11-08)', () => {
    const t = uniform([['#', 'Tabela', 'QE', 'R1', 'R2', 'R3']], 1504)
    const laid = layoutTableColumns(t, LETTER, tenPx)
    expect(laid.changed).toBe(false)
    expect(laid.phases).toEqual([])
    expect(laid.widths).toEqual([1504, 1504, 1504, 1504, 1504, 1504])
    expect(expandAutofitColWidths(t, LETTER.avail, LETTER.fit, null)).toBe(t)
  })

  it('009: a column grows to its unbreakable token and the others give the growth back (declared total kept)', () => {
    const token = 'qualificada/simples/nenhuma/risco'
    const t = uniform([['Item', 'Descricao', 'Base', token]], 2256)
    // the python-docx grid already spans the column (A4, 9024 of 9026 twips)
    const laid = layoutTableColumns(t, { avail: 9024 + 1440, fit: 9024 }, tenPx)
    // min-content = token 33 x 10 px + 14.4 px cell margins, aligned up to the px grid (345 px)
    const min = 15 * Math.ceil(token.length * 10 + 216 / 15)
    expect(laid.phases).toEqual(['grow-min', 'reclaim'])
    expect(laid.widths![3]).toBe(min)
    expect(laid.widths!.reduce((a, b) => a + b, 0)).toBe(4 * 2256)
    // the 2919 overflow comes out pro rata of each column's surplus over its own
    // min-content (825 / 1575 / 825 -> surplus 1431 / 681 / 1431 of 3543)
    expect(laid.widths).toEqual([1077, 1695, 1077, min])
  })

  it('009 tokenizer: "/" is not a break opportunity, a space is (probe 2026-11-08)', () => {
    const slash = [{ text: 'qualificada/simples/nenhuma/risco' }]
    const spaced = [{ text: 'qualificada simples nenhuma risco' }]
    expect(heuristic.wordPx(slash)).toBeGreaterThan(heuristic.wordPx(spaced) * 2)
    expect(heuristic.measured).toBe(false)
  })

  it('078: CJK min-content breaks between any two characters except before a closer / after an opener (probe 2026-11-09)', () => {
    const w = (text: string) => heuristic.wordPx([{ text, sizeHalfPoints: 20 }])
    const em = w('\u3042')
    expect(w('\u3042\u3044')).toBe(em)
    expect(w('\u5e83\u544a\u5b85\u914d')).toBe(em)
    expect(w('\u3042\u30fc\u3044')).toBe(em)
    expect(w('\u3042\u3041\u3044')).toBe(em)
    expect(w('\u3042\u3001')).toBe(2 * em)
    expect(w('\uff08\u3042')).toBe(2 * em)
    expect(w('\u3042\u3044\u3046\u3001\u3048\u304a')).toBe(2 * em)
    expect(w('\u3042\u30fb\u3044')).toBe(2 * em)
    expect(w('\u3042SNS\u3044')).toBe(w('SNS'))
    expect(w('SNS\u3001 x')).toBe(w('SNS') + em)
    expect(w('\uff08SNS \u3042')).toBe(em + w('SNS'))
  })

  it('035: the grown column takes the widest word of any row, merged-away cells demand nothing', () => {
    const t: TableModel = {
      rows: [
        [cell('N.'), cell('Descricao'), cell('Data')],
        [cell('1'), cell('intercorrencias/desvios'), cell('2025')],
        [cell('2'), { paras: ['merged'], vMerge: 'continue' }, cell('2025')],
      ],
      colWidthsTwips: [1800, 1800, 1800],
      colWidthsPct: [33.3, 33.3, 33.4],
      autoLayout: true,
    }
    const mins = minContentColTwips(t, 3, tenPx)
    expect(mins).toEqual([
      15 * Math.ceil(2 * 10 + 216 / 15),
      15 * Math.ceil(23 * 10 + 216 / 15),
      15 * Math.ceil(4 * 10 + 216 / 15),
    ])
    const laid = layoutTableColumns(t, LETTER, tenPx)
    // under the column the growth is not reclaimed (Word: 6000 -> 6192, neighbours unchanged, 2026-09-17)
    expect(laid.phases).toEqual(['grow-min'])
    expect(laid.widths).toEqual([1800, mins[1], 1800])
  })

  it('heuristic measure: 8 % slack over the estimate, none for a grid Word laid out (layoutGrid)', () => {
    const t: TableModel = {
      rows: [[cell('Levantamento'), cell('x')]],
      colWidthsTwips: [600, 600],
      colWidthsPct: [50, 50],
      autoLayout: true,
    }
    const px = heuristic.wordPx([{ text: 'Levantamento' }])
    expect(minContentColTwips(t, 2, heuristic)[0]).toBe(Math.ceil(px * 1.08 * 15) + 216)
    expect(minContentColTwips({ ...t, layoutGrid: true }, 2, heuristic, true)[0]).toBe(
      Math.ceil(px * 15) + 216,
    )
  })

  it('one token wider than the column: the minimums scale together instead of spilling', () => {
    const t = uniform([['a', 'x'.repeat(80)]], 1000)
    const laid = layoutTableColumns(t, { avail: 3000, fit: 3000 }, tenPx)
    expect(laid.phases).toEqual(['grow-min', 'reclaim', 'scale'])
    expect(laid.widths!.reduce((a, b) => a + b, 0)).toBe(3000)
    expect(laid.widths!.every((w) => w >= 1)).toBe(true)
  })

  it('008: min-contents over the column keep the side margins and scale the text widths (probe 2026-11-09)', () => {
    // tenPx: 40 m = 400 px -> 15 x ceil(400 + 14.4) = 6225 twips min, 5 m -> 975; sum 9150 > 9026
    const t = uniform(
      [
        ['A', 'B', 'C', 'D'],
        ['m'.repeat(40), 'mmmmm', 'mmmmm', 'mmmmm'],
      ],
      2256,
    )
    const laid = layoutTableColumns(t, { avail: 10466, fit: 9026 }, tenPx)
    expect(laid.phases).toEqual(['grow-min', 'reclaim', 'scale'])
    // s = (9026 - 4 x 216) / (6009 + 3 x 759) = 0.985: 216 + 6009 s = 6135 (-1 rounding drift), 216 + 759 s = 964;
    // a proportional scale would have given 6141 / 962
    expect(laid.widths).toEqual([6134, 964, 964, 964])
  })

  it('090: compat < 15 pct table spans the column plus both cell margins (probe 2026-11-09)', () => {
    const t: TableModel = {
      rows: [[cell('id'), cell('a much longer description of the row')]],
      colWidthsTwips: [2640, 2640],
      colWidthsPct: [50, 50],
      widthPct: 100,
      autoLayout: true,
      contentAutofit: true,
    }
    const sum = (w: number[] | null | undefined) => w!.reduce((a, b) => a + b, 0)
    expect(sum(layoutTableColumns(t, LETTER, tenPx, { legacy: true }).widths)).toBe(9360 + 216)
    // without content-autofit or grow-min the hung widths still commit as absolute columns
    const fixedCells = layoutTableColumns({ ...t, contentAutofit: false }, LETTER, tenPx, {
      legacy: true,
    })
    expect(fixedCells.changed).toBe(true)
    expect(sum(fixedCells.widths)).toBe(9360 + 216)
    expect(layoutTableColumns({ ...t, contentAutofit: false }, LETTER, tenPx).changed).toBe(false)
    // a column grown to its min-content (or the scale) keeps the hung width too
    const wide = { ...t, rows: [[cell('x'.repeat(700)), cell('y'.repeat(100))]] }
    const grown = layoutTableColumns(wide, LETTER, tenPx, { legacy: true })
    expect(grown.phases).toContain('scale')
    expect(sum(grown.widths)).toBe(9360 + 216)
    expect(sum(layoutTableColumns(t, LETTER, tenPx).widths)).toBe(9360)
    expect(sum(layoutTableColumns(t, LETTER, tenPx, { legacy: true, nested: true }).widths)).toBe(
      9360,
    )
    expect(
      sum(layoutTableColumns({ ...t, floatSide: 'left' }, LETTER, tenPx, { legacy: true }).widths),
    ).toBe(9360)
    expect(
      sum(
        layoutTableColumns({ ...t, cellMarTwips: { left: 200, right: 200 } }, LETTER, tenPx, {
          legacy: true,
        }).widths,
      ),
    ).toBe(9360 + 400)
  })

  it('pct table without tcW: pct resolution then CSS-style content distribution', () => {
    const t: TableModel = {
      rows: [[cell('id'), cell('a much longer description of the row')]],
      colWidthsTwips: [2640, 2640],
      colWidthsPct: [50, 50],
      widthPct: 100,
      autoLayout: true,
      contentAutofit: true,
    }
    const laid = layoutTableColumns(t, LETTER, tenPx)
    expect(laid.phases).toEqual(['pct-resolve', 'content-autofit'])
    expect(laid.resolvedPct).toBe(true)
    expect(laid.widths!.reduce((a, b) => a + b, 0)).toBe(9360)
    expect(laid.widths![0]).toBeLessThan(laid.widths![1])
    const model = expandAutofitColWidths(t, LETTER.avail, LETTER.fit, null)
    expect(model.widthPct).toBeUndefined()
  })

  it('pct table whose grid stands still bounds its nested tables by the resolved shares', () => {
    const nested: TableModel = {
      rows: [[cell('n'), cell('m')]],
      colWidthsTwips: [4000, 4000],
      colWidthsPct: [50, 50],
      autoLayout: true,
    }
    const t: TableModel = {
      rows: [[{ paras: [''], nestedTables: [nested] }, cell('b')]],
      colWidthsTwips: [100, 100],
      colWidthsPct: [50, 50],
      widthPct: 50,
      autoLayout: true,
    }
    const laid = layoutTableColumns(t, LETTER, tenPx)
    expect(laid.changed).toBe(false)
    expect(laid.widths).toEqual([2340, 2340])
    const model = expandAutofitColWidths(t, LETTER.avail, LETTER.fit, null)
    const nt = model.rows[0][0].nestedTables![0]
    // the nested grid fits the 2340 - 216 cell, not the 100-twip placeholder nor the column
    expect(nt.colWidthsTwips!.reduce((a, b) => a + b, 0)).toBe(2340 - 216)
  })
})

describe('056 noWrap pins', () => {
  const t: TableModel = {
    rows: [[{ paras: ['02 Jan 2025'], noWrap: true }, cell('text')]],
    colWidthsTwips: [1530, 7000],
    colWidthsPct: [18, 82],
    autoLayout: true,
  }
  const budget = { ...LETTER, paper: 12240 }
  it('pinned sum under the column: the noWrap cell holds its tcW', () => {
    const r = resolveTableWidth({ ...t, noWrapPinnedTwips: 1530 }, budget, { legacyIndent: false })
    expect(r.overWide).toBe(false)
  })
  it('pinned sum over the column: the pinned columns scale and the noWrap cells wrap (probe 2026-11-08)', () => {
    const r = resolveTableWidth({ ...t, noWrapPinnedTwips: 1530 + 12915 }, budget, {
      legacyIndent: false,
    })
    expect(r.overWide).toBe(true)
  })
})

describe('distribution helpers', () => {
  it('reclaimSurplus takes pro rata over the floor and never below it', () => {
    expect(reclaimSurplus([1000, 2000, 500], [500, 500, 500], 1500)).toEqual([625, 875, 500])
    expect(reclaimSurplus([500, 500], [500, 500], 100)).toEqual([500, 500])
  })
  it('scaleToFit compresses proportionally with a 1-twip floor', () => {
    expect(scaleToFit([3000, 1000], 2000)).toEqual([1500, 500])
    expect(scaleToFit([3000, 1], 2000)).toEqual([1999, 1])
    const same = [100, 200]
    expect(scaleToFit(same, 300)).toBe(same)
  })
  it('scaleToFit over floors keeps the floors and scales the rest; floors over the budget scale everything', () => {
    expect(scaleToFit([3000, 1000], 2000, [200, 200])).toEqual([1444, 556])
    expect(scaleToFit([3000, 1000], 300, [200, 200])).toEqual([225, 75])
    expect(scaleToFit([3000, 1000, 1000], 2000, [100, 100, 100])).toEqual([1148, 426, 426])
  })
  it('distributeContentAutofit: over the max sum by max, between by (max - min), under by scale', () => {
    expect(distributeContentAutofit(600, [100, 100], [200, 200])).toEqual([300, 300])
    expect(distributeContentAutofit(300, [100, 100], [200, 300])).toEqual([133, 167])
    expect(distributeContentAutofit(100, [100, 100], [200, 200])).toEqual([50, 50])
  })
  it('clampColumnWidths: null when the grid fits, else the surplus over min-content is cut', () => {
    const t = uniform([['a', 'b']], 5000)
    expect(clampColumnWidths(t, 12240, tenPx)).toBeNull()
    const clamped = clampColumnWidths(t, 6000, tenPx)!
    expect(clamped.widths).toEqual([3000, 3000])
    expect(clamped.budget).toBe(6000)
  })
})

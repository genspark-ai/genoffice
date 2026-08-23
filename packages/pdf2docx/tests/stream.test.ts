/** Stream (borderless) table detection unit tests: hand-built chars, no wasm. */
import { describe, expect, it } from 'vitest'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { detectStreamTables } from '../src/analyze/stream'
import { clusterUnitRows, splitIntoUnits } from '../src/analyze/units'
import type { Fill, PageShapes, PdfChar, Stroke, TableBlock } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

const shapesOf = (strokes: Stroke[] = [], fills: Fill[] = []): PageShapes => ({
  strokes,
  fills,
  ignoredPaths: 0,
})

const hStroke = (x0: number, x1: number, y: number): Stroke => ({
  box: { x0, x1, y0: y - 0.5, y1: y + 0.5 },
  orientation: 'h',
  widthPt: 1,
  color: '000000',
})

const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

const cellText = (table: TableBlock, r: number, c: number): string =>
  table.rows[r]![c]!.blocks.map((b) =>
    b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' '),
  ).join(' ')

/** 3×2 left-aligned grid: rows y 700/680/660, columns x 100 / 250 */
function grid3x2(): PdfChar[] {
  const rows: Array<[string, string, number]> = [
    ['Alpha', 'One', 700],
    ['Beta', 'Two', 680],
    ['Gamma', 'Three', 660],
  ]
  return rows.flatMap(([a, b, y]) => [
    ...mkText(a, 100, { y }).chars,
    ...mkText(b, 250, { y }).chars,
  ])
}

describe('splitIntoUnits', () => {
  it('splits a line at column-scale gaps, keeps word gaps together', () => {
    const chars = [...mkText('Alpha One', 100).chars, ...mkText('Beta', 250).chars]
    const units = unitsOf(chars)
    expect(units).toHaveLength(2)
    expect(units[0]!.chars.map((c) => c.text).join('')).toBe('Alpha One')
    expect(units[1]!.chars.map((c) => c.text).join('')).toBe('Beta')
  })

  it('does not split TOC dot leaders (small gaps between dots)', () => {
    const chars = mkText('Introduction', 72).chars
    for (let x = 140; x <= 374; x += 6) chars.push(mkChar('.', x, { width: 2 }))
    chars.push(...mkText('42', 378).chars)
    expect(unitsOf(chars)).toHaveLength(1)
  })
})

describe('clusterUnitRows', () => {
  it('merges same-baseline units from separately extracted lines into one row', () => {
    // column-major content order → 6 raw lines; rows must re-pair them
    const chars = [
      ...mkText('Alpha', 100, { y: 700 }).chars,
      ...mkText('Beta', 100, { y: 680 }).chars,
      ...mkText('One', 250, { y: 700 }).chars,
      ...mkText('Two', 250, { y: 680 }).chars,
    ]
    const rows = clusterUnitRows(unitsOf(chars))
    expect(rows).toHaveLength(2)
    expect(rows[0]!.units).toHaveLength(2)
    expect(rows[0]!.units[0]!.chars[0]!.text).toBe('A') // x-sorted within the row
  })
})

describe('detectStreamTables: positives', () => {
  it('detects a left-aligned 3×2 grid and routes text into cells', () => {
    const { tables, remainingUnits } = detectStreamTables(unitsOf(grid3x2()), shapesOf())
    expect(tables).toHaveLength(1)
    expect(remainingUnits).toHaveLength(0)
    const table = tables[0]!
    expect(table.rows).toHaveLength(3)
    expect(table.rows[0]).toHaveLength(2)
    expect(table.confidence).toBeGreaterThanOrEqual(0.5)
    expect(cellText(table, 0, 0)).toBe('Alpha')
    expect(cellText(table, 1, 0)).toBe('Beta')
    expect(cellText(table, 2, 1)).toBe('Three')
    expect(table.colWidthsPt).toHaveLength(2)
  })

  it('detects the same grid from column-major content order', () => {
    const chars = [
      ...mkText('Alpha', 100, { y: 700 }).chars,
      ...mkText('Beta', 100, { y: 680 }).chars,
      ...mkText('Gamma', 100, { y: 660 }).chars,
      ...mkText('One', 250, { y: 700 }).chars,
      ...mkText('Two', 250, { y: 680 }).chars,
      ...mkText('Three', 250, { y: 660 }).chars,
    ]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(cellText(tables[0]!, 0, 1)).toBe('One')
  })

  it('accepts a right-aligned column via its right edge', () => {
    const chars = [
      ...mkText('Item', 100, { y: 700 }).chars,
      ...mkText('5', 275, { y: 700 }).chars, // all right edges at 280
      ...mkText('Widget', 100, { y: 680 }).chars,
      ...mkText('120', 265, { y: 680 }).chars,
      ...mkText('Gadget', 100, { y: 660 }).chars,
      ...mkText('37', 270, { y: 660 }).chars,
    ]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(cellText(tables[0]!, 1, 1)).toBe('120')
  })

  it('two-row grids pass only WITH stroke evidence (booktabs three-line table)', () => {
    const chars = [
      ...mkText('Name', 100, { y: 700 }).chars,
      ...mkText('Qty', 250, { y: 700 }).chars,
      ...mkText('Total', 100, { y: 680 }).chars,
      ...mkText('9', 250, { y: 680 }).chars,
    ]
    // bare 2×2: reject (miss rather than misfire)
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
    // with top/mid/bottom rules: accept
    const rules = [hStroke(95, 290, 712), hStroke(95, 290, 692), hStroke(95, 290, 672)]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf(rules))
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(2)
    expect(cellText(tables[0]!, 1, 0)).toBe('Total')
  })

  it('row-banding fills raise confidence', () => {
    const units = unitsOf(grid3x2())
    const plain = detectStreamTables(units, shapesOf()).tables[0]!
    const banded = detectStreamTables(
      unitsOf(grid3x2()),
      shapesOf([], [{ box: { x0: 95, y0: 676, x1: 290, y1: 690 }, color: 'EEEEEE' }]),
    ).tables[0]!
    expect(banded.confidence!).toBeGreaterThan(plain.confidence!)
  })

  it('zebra band fills become cell shading (P16 E)', () => {
    // band behind the middle row (y 680 baseline → band ~676..692)
    const banded = detectStreamTables(
      unitsOf(grid3x2()),
      shapesOf([], [{ box: { x0: 95, y0: 674, x1: 290, y1: 694 }, color: 'E7E7E7' }]),
    ).tables[0]!
    expect(banded.rows[1]!.map((c) => c.fill)).toEqual(['E7E7E7', 'E7E7E7'])
    expect(banded.rows[0]!.every((c) => c.fill === undefined)).toBe(true)
  })
})

describe('detectStreamTables: slide side-by-side regions (P16 E)', () => {
  /** left: financial grid with a wrapped two-line label; right: prose whose
   * lines interleave the grid rows vertically (absolute slide layout) */
  const slideChars = (): PdfChar[] => [
    // grid: label column x40, numeric columns x150/x220/x290
    ...mkText('Revenue', 40, { y: 700 }).chars,
    ...mkText('62', 150, { y: 700 }).chars,
    ...mkText('129', 220, { y: 700 }).chars,
    ...mkText('336', 290, { y: 700 }).chars,
    ...mkText('Customers', 40, { y: 680 }).chars,
    ...mkText('500', 150, { y: 680 }).chars,
    ...mkText('1200', 220, { y: 680 }).chars,
    ...mkText('3000', 290, { y: 680 }).chars,
    ...mkText('Price', 40, { y: 660 }).chars,
    ...mkText('25', 150, { y: 660 }).chars,
    ...mkText('18', 220, { y: 660 }).chars,
    ...mkText('16', 290, { y: 660 }).chars,
    // wrapped label row: label line sits between value baselines
    ...mkText('Expenses', 40, { y: 645 }).chars,
    ...mkText('85', 150, { y: 638 }).chars,
    ...mkText('650', 220, { y: 638 }).chars,
    ...mkText('250', 290, { y: 638 }).chars,
    ...mkText('COGS', 40, { y: 620 }).chars,
    ...mkText('18', 150, { y: 620 }).chars,
    ...mkText('17', 220, { y: 620 }).chars,
    ...mkText('15', 290, { y: 620 }).chars,
    // right prose region (x 500+), lines interleaving the grid rows
    ...mkText('The revenue expenses and net income of', 500, { y: 693 }).chars,
    ...mkText('the venture over the next five years grow', 500, { y: 671 }).chars,
    ...mkText('focus on the three most important drivers', 500, { y: 651 }).chars,
    ...mkText('of your revenue and expenses over time', 500, { y: 628 }).chars,
  ]

  it('interleaved prose breaks the whole-page pass (current bias)', () => {
    const { tables } = detectStreamTables(unitsOf(slideChars()), shapesOf())
    expect(tables).toHaveLength(0)
  })

  it('slideRegions splits at the vast valley and finds the strong grid', () => {
    const { tables, remainingUnits } = detectStreamTables(unitsOf(slideChars()), shapesOf(), [], {
      slideRegions: true,
    })
    expect(tables).toHaveLength(1)
    const t = tables[0]!
    expect(t.rows.length).toBeGreaterThanOrEqual(5)
    expect(t.rows[0]!.length).toBe(4)
    expect(cellText(t, 0, 0)).toBe('Revenue')
    expect(cellText(t, 0, 3)).toBe('336')
    // the wrapped label joined the grid instead of breaking it
    expect(t.rows.some((row) => row[0] && cellText(t, t.rows.indexOf(row), 0) === 'Expenses')).toBe(
      true,
    )
    // prose stays in the flow
    const remainingText = remainingUnits.map((u) => u.chars.map((c) => c.text).join('')).join(' ')
    expect(remainingText).toContain('The revenue expenses')
    expect(remainingText).not.toContain('Customers')
  })

  it('slideRegions never mints a table from two prose regions', () => {
    const chars = [
      ...mkText('A first paragraph line of running text', 40, { y: 700 }).chars,
      ...mkText('and a second line below it that flows', 40, { y: 680 }).chars,
      ...mkText('with a third line closing the thought', 40, { y: 660 }).chars,
      ...mkText('The right column holds more sentences', 500, { y: 693 }).chars,
      ...mkText('that keep running as normal body text', 500, { y: 673 }).chars,
      ...mkText('down the slide without any grid at all', 500, { y: 653 }).chars,
    ]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf(), [], { slideRegions: true })
    expect(tables).toHaveLength(0)
  })
})

describe('detectStreamTables: negative family (never tables)', () => {
  it('plain paragraph lines', () => {
    const chars = [
      ...mkText('The quick brown fox jumps over the lazy', 72, { y: 700 }).chars,
      ...mkText('dog while the cat watches from a warm', 72, { y: 686 }).chars,
      ...mkText('windowsill in the afternoon sun today.', 72, { y: 672 }).chars,
    ]
    const units = unitsOf(chars)
    const { tables, remainingUnits } = detectStreamTables(units, shapesOf())
    expect(tables).toHaveLength(0)
    expect(remainingUnits).toHaveLength(units.length)
  })

  it('poetry short lines', () => {
    const chars = [
      ...mkText('Roses are red', 200, { y: 700 }).chars,
      ...mkText('Violets are blue', 195, { y: 686 }).chars,
      ...mkText('Code has no bugs', 198, { y: 672 }).chars,
      ...mkText('Until you look', 202, { y: 658 }).chars,
    ]
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
  })

  it('side-by-side stanzas become ONE exact-geometry verse row, not a grid (P22 E)', () => {
    const left = [
      'Wer reitet so spat durch Nacht,',
      'Es ist der Vater;',
      'Er hat den Knaben wohl in dem Arm,',
      'Er fasst ihn warm.',
    ]
    const right = [
      'Mein Sohn was birgst du so bang,',
      'Siehst Vater du?',
      'Den Erlenkonig mit Kron und Schweif,',
      'Mein Sohn ein Streif.',
    ]
    const chars = [
      ...left.flatMap((t, i) => mkText(t, 80, { y: 700 - i * 14 }).chars),
      ...right.flatMap((t, i) => mkText(t, 330, { y: 700 - i * 14 }).chars),
      // body prose below keeps the pair under the page-columns coverage gate
      ...Array.from(
        { length: 10 },
        (_, i) =>
          mkText('Body prose line that runs the full measure of the page here', 80, {
            y: 560 - i * 13,
          }).chars,
      ).flat(),
    ]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    const t = tables[0]!
    // one row of two stack cells — NOT a 4×2 grid
    expect(t.rows).toHaveLength(1)
    expect(t.rows[0]).toHaveLength(2)
    const cellLines = (c: number): string[] =>
      t.rows[0]![c]!.blocks.map((b) =>
        b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(''),
      )
    expect(cellLines(0)).toHaveLength(4)
    expect(cellLines(1)[0]).toContain('Mein Sohn was birgst')
  })

  it('absorbs an orphan stanza row the run gate dropped (P22 E)', () => {
    const left = [
      'Wer reitet so spat durch Nacht,',
      'Es ist der Vater;',
      'Er hat den Knaben wohl in dem Arm,',
      'Er fasst ihn warm.',
    ]
    const right = [
      'Mein Sohn was birgst du so bang,',
      'Siehst Vater du?',
      'Den Erlenkonig mit Kron und Schweif,',
      'Mein Sohn ein Streif.',
    ]
    const chars = [
      ...left.flatMap((t, i) => mkText(t, 80, { y: 700 - i * 14 }).chars),
      ...right.flatMap((t, i) => mkText(t, 330, { y: 700 - i * 14 }).chars),
      // 5th right-only line, one pitch below the run
      ...mkText('Und bleibt allein zuruck hier. ', 330, { y: 700 - 4 * 14 }).chars,
      ...Array.from(
        { length: 10 },
        (_, i) =>
          mkText('Body prose line that runs the full measure of the page here', 80, {
            y: 560 - i * 13,
          }).chars,
      ).flat(),
    ]
    const { tables, remainingUnits } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    // the orphan joined the right cell instead of falling into the flow
    // (only the body prose stays out)
    expect(remainingUnits).toHaveLength(10)
    const rightText = tables[0]!.rows[0]![1]!.blocks.map((b) =>
      b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(''),
    ).join('|')
    expect(rightText).toContain('Und bleibt allein')
  })

  it('two-column body text (page columns, not a table)', () => {
    const left = [
      'The quick brown fox jumps ov',
      'lazy dog while a cat watches',
      'from the windowsill and then',
      'sunlight drifts over the roo',
      'pages of an open book turnin',
      'slowly in the afternoon bree',
    ]
    const right = [
      'Meanwhile across the road th',
      'baker sets out fresh loaves,',
      'steam rising from the crusts',
      'as customers queue politely,',
      'coins ready in their pockets',
      'waiting for the door to open',
    ]
    const chars = left.flatMap((t, i) => [
      ...mkText(t, 72, { y: 700 - i * 14 }).chars,
      ...mkText(right[i]!, 320, { y: 700 - i * 14 }).chars,
    ])
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
  })

  it('page-column tails with ragged paragraph-final lines and a footer row (P23)', () => {
    // childAttachments src31: the bottom rows of a real 2-column page plus a
    // short footer row formed a phantom stream table — the ragged lines and
    // footer diluted the MEAN fill just under the prose veto. The majority of
    // entries are still individually full-width sentences.
    const left = [
      'copy of Form 941 or Form 944 and separate',
      'with a revision date showing the year your',
      'quent return is being filed by the office',
      'publication, for various ways to secure it',
      'Publication 15',
    ]
    const right = [
      'Social security wages and social security',
      'Medicare wages and tips on Form W-3 should',
      'clude Forms 941 or Form 944 adjustments on',
      'the current year',
      'Page 31',
    ]
    const chars = left.flatMap((t, i) => [
      ...mkText(t, 72, { y: 700 - i * 14 }).chars,
      ...mkText(right[i]!, 320, { y: 700 - i * 14 }).chars,
    ])
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
  })

  it('TOC lines with dot leaders', () => {
    const chars: PdfChar[] = []
    const entries: Array<[string, string]> = [
      ['Introduction', '1'],
      ['Background', '7'],
      ['Methodology', '19'],
    ]
    for (const [i, [title, page]] of entries.entries()) {
      const y = 700 - i * 16
      const t = mkText(title, 72, { y })
      chars.push(...t.chars)
      for (let x = Math.ceil(t.endX) + 6; x <= 374; x += 6) {
        chars.push(mkChar('.', x, { width: 2, y }))
      }
      chars.push(...mkText(page, 380, { y }).chars)
    }
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
  })

  it('code block with indentation', () => {
    const chars = [
      ...mkText('function total(items) {', 72, { y: 700 }).chars,
      ...mkText('const sum = items.reduce(add)', 90, { y: 686 }).chars,
      ...mkText('return sum * TAX_RATE', 90, { y: 672 }).chars,
      ...mkText('}', 72, { y: 658 }).chars,
    ]
    expect(detectStreamTables(unitsOf(chars), shapesOf()).tables).toHaveLength(0)
  })

  it('candidates overlapping a lattice table are rejected', () => {
    const { tables } = detectStreamTables(unitsOf(grid3x2()), shapesOf(), [
      { x0: 90, y0: 650, x1: 300, y1: 710 },
    ])
    expect(tables).toHaveLength(0)
  })
})

// ── P27: sparse-column alignment + page-columns word gate ──

describe('P27 stream gates', () => {
  it('a sparse header column straddling body columns does not veto the grid', () => {
    // 6 body rows of 3 aligned columns + one header unit straddling cols 1-2:
    // the straddle merges into one wide interval on the header row only —
    // present on 1 of 7 rows, it must stay neutral (mexican_towns class)
    const chars: PdfChar[] = []
    const ys = [700, 685, 670, 655, 640, 625]
    for (const y of ys) {
      chars.push(...mkText('01', 100, { y }).chars)
      chars.push(...mkText('Aguas', 180, { y }).chars)
      chars.push(...mkText('0094', 300, { y }).chars)
    }
    chars.push(...mkText('Clave', 100, { y: 715 }).chars)
    chars.push(...mkText('NombreEntidad', 205, { y: 715 }).chars)
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows.length).toBeGreaterThanOrEqual(6)
  })

  it('a full-page table of short phrase cells is not vetoed as page columns', () => {
    // uniform pitch + full coverage + moderately wide first column, but
    // cells average ~2 words — a data table, not newspaper columns
    const chars: PdfChar[] = []
    for (let i = 0; i < 12; i++) {
      const y = 700 - i * 15
      chars.push(...mkText('Township Council Office', 90, { y }).chars)
      chars.push(...mkText(`${100 + i}`, 320, { y }).chars)
      chars.push(...mkText(`${i} units`, 420, { y }).chars)
    }
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(tables[0]!.colWidthsPt.length).toBe(3)
  })

  it('still vetoes uniform-pitch wordy page columns (running prose)', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 12; i++) {
      const y = 700 - i * 15
      chars.push(...mkText('the quick brown fox jumps over the dog', 60, { y, fontSize: 10 }).chars)
      chars.push(
        ...mkText('while the lazy cat sleeps on a warm sill', 300, { y, fontSize: 10 }).chars,
      )
    }
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(0)
  })
})

describe('P27 sparse-leading runs', () => {
  it('pairs strongly tabular rows across gaps beyond the running-text ratio', () => {
    // 6 short numeric units per row at 25pt pitch (journal-table leading):
    // gap ≈ 15.7pt > 1.6 × 9.3pt row height, so the base rule alone fails
    const ys = [700, 675, 650, 625, 600]
    const chars = ys.flatMap((y) =>
      [100, 160, 220, 280, 340, 400].flatMap(
        (x, i) => mkText(i === 0 ? 'row' : String(10 * i), x, { y }).chars,
      ),
    )
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(5)
    expect(tables[0]!.colWidthsPt).toHaveLength(6)
  })

  it('does not glue 2-unit prose fragments across the same gaps', () => {
    const ys = [700, 675, 650, 625, 600]
    const chars = ys.flatMap((y) => [
      ...mkText('A sentence fragment of running prose text here', 100, { y }).chars,
      ...mkText('another trailing clause', 380, { y }).chars,
    ])
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(0)
  })

  it('absorbs a wrapped-label row between strong value rows', () => {
    // value rows at 20pt pitch with a 1-unit label continuation between them
    const mkRow = (y: number) =>
      [100, 160, 220, 280, 340, 400].flatMap(
        (x, i) => mkText(i === 0 ? 'lbl' : String(i), x, { y }).chars,
      )
    const chars = [
      ...mkRow(700),
      ...mkRow(680),
      ...mkText('(wrapped)', 100, { y: 668 }).chars, // sub-minimum row inside the run
      ...mkRow(656),
      ...mkRow(636),
    ]
    const { tables } = detectStreamTables(unitsOf(chars), shapesOf())
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows.length).toBeGreaterThanOrEqual(4)
  })
})

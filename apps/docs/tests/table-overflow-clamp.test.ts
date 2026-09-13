import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, type Block, type TableModel } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  clampTableColWidths,
  expandAutofitColWidths,
  legacyIndentTable,
  tableModelToPmNode,
} from '../src/renderer/editor/convert'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

type Spec = [string, Record<string, string>, ...unknown[]]

const cell = (text: string) => ({ paras: [text] })

describe('clampTableColWidths', () => {
  it('compresses grids past the cap, real columns floored at min-content', () => {
    // regression sample 08: a 130620472-twips garbage gridCol laid a table out
    // ~8.7M px wide (fixed layout grows past width/max-width to the col sum);
    // the garbage column absorbs the cut, 'label' keeps its word unbroken (766 twips)
    const model: TableModel = {
      rows: [[cell('label'), cell('value')]],
      colWidthsTwips: [6236, 130620472],
      colWidthsPct: [0.0048, 99.9952],
    }
    const clamped = clampTableColWidths(model, 10886)
    expect(clamped.colWidthsTwips).toEqual([766, 10120])
    expect(clamped.colWidthsPct!.map((w) => Math.round(w))).toEqual([7, 93])
    // display-only: the input model keeps the raw document values
    expect(model.colWidthsTwips).toEqual([6236, 130620472])
  })

  it('returns the same model when the grid fits under the cap', () => {
    const model: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [4680, 4680],
      indentTwips: 1450,
    }
    expect(clampTableColWidths(model, 12240)).toBe(model)
    // wider than the text column but within the paper: Word spills, no compression
    const spilling: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [6000, 5800],
    }
    expect(clampTableColWidths(spilling, 12240)).toBe(spilling)
  })

  it('takes a positive left indent out of the budget', () => {
    const model: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [6120, 6120],
      indentTwips: 1450,
    }
    // 1450 over budget, taken proportionally from the two equal columns
    expect(clampTableColWidths(model, 12240).colWidthsTwips).toEqual([5395, 5395])
  })

  it('clamps nested tables to their (clamped) cell content width', () => {
    const model: TableModel = {
      rows: [
        [
          cell('left'),
          {
            paras: [''],
            nestedTables: [
              {
                rows: [[cell('a'), cell('b')]],
                colWidthsTwips: [1871, 130618601],
              },
            ],
          },
        ],
      ],
      colWidthsTwips: [6236, 130620472],
    }
    const clamped = clampTableColWidths(model, 10886)
    const nested = clamped.rows[0][1].nestedTables![0]
    // outer col1 squeezes to min-content('left') = 590, col2 gets the rest (10296);
    // nested budget = 10296 - 2 x 108 cell margins, its 'a' column floors at 360 (24px)
    expect(clamped.colWidthsTwips).toEqual([590, 10296])
    expect(nested.colWidthsTwips).toEqual([360, 10296 - 216 - 360])
    expect(model.rows[0][1].nestedTables![0].colWidthsTwips).toEqual([1871, 130618601])
  })

  it('bounds nested tables by the page budget when the outer grid is unusable', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: [''],
            nestedTables: [{ rows: [[cell('a')]], colWidthsTwips: [130618601] }],
          },
        ],
      ],
    }
    const clamped = clampTableColWidths(model, 10886)
    expect(clamped.rows[0][0].nestedTables![0].colWidthsTwips).toEqual([10886])
  })
})

describe('expandAutofitColWidths', () => {
  // 'الموقف البيئي' at the default 12pt: 6 chars x 0.35em x 16px = 33.6px word
  // -> ceil(33.6 x 1.08 slack x 15) + 2 x 108 default margins = 761 twips min-content
  const ARABIC_HEADER = 'الموقف البيئي'
  const MIN_ARABIC_COL = 761

  it('widens an autofit column to its widest unbreakable word', () => {
    const model: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell('ok')]],
      colWidthsTwips: [567, 4000],
      autoLayout: true,
    }
    const expanded = expandAutofitColWidths(model, 10772, 9638)
    expect(expanded.colWidthsTwips![0]).toBe(MIN_ARABIC_COL)
    expect(expanded.colWidthsTwips![1]).toBe(4000)
    // display-only: the input model keeps the raw document values
    expect(model.colWidthsTwips).toEqual([567, 4000])
  })

  it('takes real advances from an injected metrics provider with rounding slack only', () => {
    const model: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell('ok')]],
      colWidthsTwips: [567, 4000],
      autoLayout: true,
    }
    // a measured 20px word + 2px edge: ceil(22 x 1.02 x 15) + 216 = 553 twips fits the declared 567
    const narrow = { measure: () => 20, metrics: () => ({ ascent: 0, descent: 0, lineHeight: 0 }) }
    expect(expandAutofitColWidths(model, 10772, 9638, narrow)).toBe(model)
    // a measured 40px word needs ceil(42 x 1.02 x 15) + 216 = 859 twips, not the 1.08 heuristic floor
    const wide = { ...narrow, measure: () => 40 }
    expect(expandAutofitColWidths(model, 10772, 9638, wide).colWidthsTwips![0]).toBe(859)
  })

  it('leaves fixed-layout tables and already-wide autofit columns alone', () => {
    const fixed: TableModel = {
      rows: [[cell(ARABIC_HEADER)]],
      colWidthsTwips: [567],
      fixedLayout: true,
    }
    expect(expandAutofitColWidths(fixed, 10772, 9638)).toBe(fixed)

    const wide: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell('ok')]],
      colWidthsTwips: [2000, 4000],
      autoLayout: true,
    }
    expect(expandAutofitColWidths(wide, 10772, 9638)).toBe(wide)
  })

  it('treats a dxa tblW without w:tblLayout fixed as preferred widths (autofit)', () => {
    // regression sample: tblW 9643 dxa, tcW 2835/2835/2835/1137,
    // "${employee.patronymic_name#}" wider than its column. Word grows that
    // column to the word, shrinks the columns with surplus and keeps the total
    const wordPx = 199
    const metrics = {
      measure: (text: string) => (text.length > 20 ? wordPx : 30),
      metrics: () => ({ ascent: 0, descent: 0, lineHeight: 0 }),
    }
    const model: TableModel = {
      rows: [
        [
          cell('${employee.name#}'),
          cell('${employee.patronymic_name#}'),
          cell('${employee.surname#}'),
          cell('${employee.age#}'),
        ],
      ],
      colWidthsTwips: [2835, 2835, 2835, 1137],
      cellMarTwips: { left: 24, right: 55 },
    }
    const expanded = expandAutofitColWidths(model, 10772, 9638, metrics)
    const widths = expanded.colWidthsTwips!
    // ceil((199 + 2) x 1.02 x 15) + 79 = 3155
    expect(widths[1]).toBe(3155)
    expect(widths[0]).toBeLessThan(2835)
    expect(widths[2]).toBeLessThan(2835)
    expect(widths.reduce((a, b) => a + b, 0)).toBe(9642)
    // the same grid under w:tblLayout fixed keeps the declared columns
    const fixed: TableModel = { ...model, fixedLayout: true }
    expect(expandAutofitColWidths(fixed, 10772, 9638, metrics)).toBe(fixed)
  })

  it('reclaims growth past the fit width from columns with surplus', () => {
    const model: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell('a')]],
      colWidthsTwips: [567, 9071],
      autoLayout: true,
    }
    const expanded = expandAutofitColWidths(model, 10772, 9638)
    expect(expanded.colWidthsTwips![0]).toBe(MIN_ARABIC_COL)
    // growth (+194) comes out of the wide column; total stays at the declared/fit width
    expect(expanded.colWidthsTwips!.reduce((a, b) => a + b, 0)).toBe(9638)
  })

  it('compresses tblW-auto preferred widths past the text column back to it', () => {
    // centered autofit table whose tcW sum (11338) runs ~24% past a 9122-twip
    // text column: Word treats tcW as preferred and fits the table to the
    // column plus its two 108-twip side cell margins (measured 623px at 96dpi)
    const model: TableModel = {
      rows: [[cell('Marco'), cell('Autores'), cell('Autores'), cell('Aporte')]],
      colWidthsTwips: [2551, 2948, 2721, 3118],
      autoLayout: true,
      align: 'center',
    }
    const fitted = expandAutofitColWidths(model, 12240, 9122)
    const total = fitted.colWidthsTwips!.reduce((a, b) => a + b, 0)
    expect(Math.abs(total - 9338)).toBeLessThanOrEqual(2)
    // proportions survive the cut
    expect(fitted.colWidthsTwips![3]).toBeGreaterThan(fitted.colWidthsTwips![0])
    expect(model.colWidthsTwips).toEqual([2551, 2948, 2721, 3118])
  })

  it('keeps a tblW-auto grid that hangs into the margins by no more than its cell margins', () => {
    // six-column timesheet grid of 9824 twips in a 9749-twip text column: Word
    // draws it at full width, the border 108 twips outside each margin edge
    const model: TableModel = {
      rows: [[cell('Date'), cell('Org'), cell('Code'), cell('Work'), cell('Time'), cell('Level')]],
      colWidthsTwips: [1296, 1008, 1296, 4032, 1008, 1184],
      autoLayout: true,
    }
    expect(expandAutofitColWidths(model, 10829, 9749)).toBe(model)
    // the hang follows the table's own cell margins
    const narrowMar = { ...model, cellMarTwips: { left: 28, right: 28 } }
    const fitted = expandAutofitColWidths(narrowMar, 10829, 9749)
    expect(fitted.colWidthsTwips!.reduce((a, b) => a + b, 0)).toBe(9749 + 56)
    // past the hang the grid compresses to column + margins, wide column first
    const wide = { ...model, colWidthsTwips: [1296, 1008, 1296, 4332, 1008, 1184] }
    const fittedWide = expandAutofitColWidths(wide, 10829, 9749).colWidthsTwips!
    expect(Math.abs(fittedWide.reduce((a, b) => a + b, 0) - 9965)).toBeLessThanOrEqual(2)
    expect(fittedWide[3]).toBeLessThan(4332)
  })

  it('leaves a full-width pct table alone even when its indent pushes it past the column', () => {
    const model: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsPct: [50, 50],
      widthPct: 100,
      indentTwips: 200,
      autoLayout: true,
    }
    expect(expandAutofitColWidths(model, 10772, 9638)).toBe(model)
  })

  it('floors pct-width autofit columns at min-content, converting to absolute widths', () => {
    // prod sample: w:tblW w="100" pct (= 2%) tip boxes broke one character per line
    const model: TableModel = {
      rows: [[cell(ARABIC_HEADER)]],
      colWidthsTwips: [100],
      colWidthsPct: [100],
      widthPct: 2,
      autoLayout: true,
    }
    const expanded = expandAutofitColWidths(model, 10772, 9638)
    expect(expanded.colWidthsTwips).toEqual([MIN_ARABIC_COL])
    expect(expanded.widthPct).toBeUndefined()
    expect(model.widthPct).toBe(2)
    // pct tables whose resolved columns already fit their words keep the pct width
    const wide: TableModel = {
      rows: [[cell('ok')]],
      colWidthsPct: [100],
      widthPct: 50,
      autoLayout: true,
    }
    expect(expandAutofitColWidths(wide, 10772, 9638)).toBe(wide)
    // pct resolves against the full text column: the render draws width:N% of
    // the content box and shifts by the indent, so the indent must not shrink
    // the resolved widths into false growth (8% of 9638 = 771 >= 761 min)
    const indented: TableModel = {
      rows: [[cell(ARABIC_HEADER)]],
      colWidthsPct: [100],
      widthPct: 8,
      indentTwips: 1450,
      autoLayout: true,
    }
    expect(expandAutofitColWidths(indented, 10772, 9638)).toBe(indented)
  })

  it('counts the list indent inside cells toward min-content', () => {
    const bullet: TableModel = {
      rows: [
        [
          {
            paras: ['word'],
            richParas: [
              { runs: [{ text: 'word' }], list: { kind: 'bullet', numId: '1', ilvl: 0 } },
            ],
          },
        ],
      ],
      colWidthsTwips: [400],
      autoLayout: true,
    }
    const expanded = expandAutofitColWidths(bullet, 10772, 9638)
    // 'word' = 2.18em x 16px = 34.88px, + 0.55in default .doc-li padding (52.8px)
    expect(expanded.colWidthsTwips![0]).toBe(Math.ceil((34.88 + 52.8) * 1.08 * 15) + 216)
  })

  it('compresses proportionally when min-contents exceed the budget', () => {
    const model: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell(ARABIC_HEADER), cell('x')]],
      colWidthsTwips: [100, 100, 500],
      autoLayout: true,
    }
    const expanded = expandAutofitColWidths(model, 800, 700)
    const widths = expanded.colWidthsTwips!
    // total stays within the fit budget (rounding slack) with no zero columns,
    // so the downstream clamp never starves the trailing columns
    expect(widths.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(703)
    for (const w of widths) expect(w).toBeGreaterThan(0)
    const clamped = clampTableColWidths(expanded, 800)
    for (const w of clamped.colWidthsTwips!) expect(w).toBeGreaterThan(0)
  })
})

describe('display clamp wiring', () => {
  const GARBAGE_TABLE =
    '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="6236"/><w:gridCol w:w="130620472"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6236"/></w:tcPr><w:p><w:r><w:t>h1</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="130620472"/></w:tcPr><w:p><w:r><w:t>h2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

  it('blocksToPmDoc clamps table attrs with section geometry, raw without', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: GARBAGE_TABLE }))
    const sections = readSections(parsed)
    const clamped = blocksToPmDoc(parsed.blocks, sections).content![0]
    // hard cap: the paper width (over-wide tables spill the margins, never past the paper)
    const paperPx = Math.round(sections[0].settings.pageWidth / 15)
    expect(clamped.attrs?.widthPx as number).toBeGreaterThanOrEqual(paperPx - 3)
    expect(clamped.attrs?.widthPx as number).toBeLessThanOrEqual(paperPx + 3)
    const raw = blocksToPmDoc(parsed.blocks).content![0]
    expect(raw.attrs?.widthPx).toBeGreaterThan(8_000_000)
    // parse keeps the document values; saving untouched tables is byte-preserving
    expect(parsed.blocks[0].table!.colWidthsTwips).toEqual([6236, 130620472])
  })

  it('tableModelToPmNode leaves models alone without an avail width', () => {
    const model: TableModel = { rows: [[cell('a')]], colWidthsTwips: [130618601] }
    expect(tableModelToPmNode(model).attrs?.widthPx).toBe(Math.round(130618601 / 15))
    expect(tableModelToPmNode(model, null, null, null, 10886).attrs?.widthPx).toBe(
      Math.round(10886 / 15),
    )
  })
})

describe('autofit expansion wiring', () => {
  const AUTO_TABLE =
    '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="567"/><w:gridCol w:w="9071"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="567"/></w:tcPr><w:p><w:r><w:t>الموقف البيئي</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9071"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

  it('parse flags autofit tables and blocksToPmDoc widens their narrow columns', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: AUTO_TABLE }))
    const table = parsed.blocks[0].table!
    expect(table.autoLayout).toBe(true)
    expect(table.colWidthsTwips).toEqual([567, 9071])
    const pm = blocksToPmDoc(parsed.blocks, readSections(parsed)).content![0]
    const colwidth = pm.content![0].content![0].attrs!.colwidth as number[]
    expect(colwidth[0]).toBeGreaterThanOrEqual(Math.floor(720 / 15))
  })

  // form grid: 12 slots summing to the 9026-twip text column, five of them
  // slivers (49-114 twips) that only ever sit inside spanned cells
  const FORM_GRID_COLS = [446, 57, 1705, 114, 1391, 49, 2193, 110, 537, 78, 1151, 1195]
  const formRow = (spans: number[]) => {
    let at = 0
    return (
      '<w:tr>' +
      spans
        .map((span) => {
          const w = FORM_GRID_COLS.slice(at, at + span).reduce((a, b) => a + b, 0)
          at += span
          return (
            `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:gridSpan w:val="${span}"/></w:tcPr>` +
            '<w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>'
          )
        })
        .join('') +
      '</w:tr>'
    )
  }
  const FORM_TABLE =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="56" w:type="dxa"/>' +
    '<w:tblCellMar><w:left w:w="28" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
    `<w:tblGrid>${FORM_GRID_COLS.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    formRow([2, 2, 2, 2, 2, 2]) +
    formRow([3, 3, 3, 3]) +
    formRow([4, 4, 4]) +
    '</w:tbl>'

  it('keeps a text-column-wide form grid with sliver columns at the text column', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: FORM_TABLE }))
    const sections = readSections(parsed)
    const fit = sections[0].settings.pageWidth - 2 * 1440
    expect(FORM_GRID_COLS.reduce((a, b) => a + b, 0)).toBe(fit)
    const pm = blocksToPmDoc(parsed.blocks, sections).content![0]
    // grid plus indent fits the column: the slivers must not be floored per column
    // (24px each pushed the whole table past the paper edge)
    const widthPx = pm.attrs!.widthPx as number
    expect(Math.abs(widthPx - fit / 15)).toBeLessThanOrEqual(4)
    const firstRow = pm.content![0].content!.flatMap((c) => c.attrs!.colwidth as number[])
    expect(firstRow.reduce((a, b) => a + b, 0)).toBe(widthPx)
  })

  const WIDE_TCW_TABLE =
    '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:jc w:val="center"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/><w:gridCol w:w="2256"/></w:tblGrid>' +
    '<w:tr>' +
    [2551, 2948, 2721, 3118]
      .map(
        (w) =>
          `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`,
      )
      .join('') +
    '</w:tr></w:tbl>'

  it('fits a centered tblW-auto table with over-wide tcW into the text column', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: WIDE_TCW_TABLE }))
    // the parser prefers the disagreeing tcW over the even grid ...
    expect(parsed.blocks[0].table!.colWidthsTwips).toEqual([2551, 2948, 2721, 3118])
    const sections = readSections(parsed)
    const pm = blocksToPmDoc(parsed.blocks, sections).content![0]
    // ... but the display width is the text column plus the side cell margins, not the 11338-twip tcW sum
    const fitPx = (sections[0].settings.pageWidth - 2 * 1440 + 216) / 15
    expect(Math.abs((pm.attrs!.widthPx as number) - fitPx)).toBeLessThanOrEqual(2)
  })

  it('parse does not flag fixed-layout tables but keeps pct tables autofit', async () => {
    const fixed = AUTO_TABLE.replace('</w:tblPr>', '<w:tblLayout w:type="fixed"/></w:tblPr>')
    const fixedTable = (await parseDocx(await buildDocx({ bodyXml: fixed }))).blocks[0].table!
    expect(fixedTable.autoLayout).toBe(undefined)
    expect(fixedTable.fixedLayout).toBe(true)
    expect(tableModelToPmNode(fixedTable).attrs?.tblFixedLayout).toBe(true)
    // dxa width without w:tblLayout fixed keeps the fit-to-page display clamp
    const auto = (await parseDocx(await buildDocx({ bodyXml: AUTO_TABLE }))).blocks[0].table!
    expect(auto.fixedLayout).toBe(undefined)
    expect(tableModelToPmNode(auto).attrs?.tblFixedLayout).toBe(false)
    // pct width picks the table size, not the layout algorithm: min-content still floors columns
    const pct = AUTO_TABLE.replace('w:type="auto" w:w="0"', 'w:type="pct" w:w="4000"')
    expect((await parseDocx(await buildDocx({ bodyXml: pct }))).blocks[0].table!.autoLayout).toBe(
      true,
    )
  })
})

describe('renderTableSpec width budget', () => {
  it('lets top-level tables spill into the right margin, caps nested ones at the cell', () => {
    const model: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [2340, 7020],
      indentTwips: 1450,
    }
    const spec = renderTableSpec(model) as Spec
    expect(spec[1].style).toContain(
      'width:min(624px,calc(var(--doc-content-w,100%) + var(--doc-margin-right,0px) - 96.7px))',
    )
    expect(spec[1].style).toContain('margin-left:96.7px')
    const nestedSpec = renderTableSpec(model, true) as Spec
    expect(nestedSpec[1].style).toContain('width:min(624px,calc(100% - 96.7px))')
  })

  it('fixed-layout tables hold the declared width past the paper edge (Word clips there)', () => {
    const model: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [2340, 7020],
      fixedLayout: true,
    }
    const style = (renderTableSpec(model) as Spec)[1].style
    expect(style).toContain('width:624px')
    expect(style).toContain('max-width:none')
    expect(style).not.toContain('min(')
    const centered = renderTableSpec({ ...model, align: 'center' }) as Spec
    expect(centered[1].style).toContain('margin-left:calc((var(--doc-content-w,100%) - 624px)/2)')
    // nested fixed tables stay capped by their cell
    expect((renderTableSpec(model, true) as Spec)[1].style).toContain('width:min(624px,100%)')
  })

  it('centered tables ignore the indent and spill both margins symmetrically', () => {
    const centered: TableModel = {
      rows: [[cell('a')]],
      colWidthsTwips: [3000],
      indentTwips: 1450,
      align: 'center',
    }
    const paper =
      'calc(var(--doc-content-w,100%) + var(--doc-margin-left,var(--doc-margin-right,0px)) + var(--doc-margin-right,0px))'
    const style = (renderTableSpec(centered) as Spec)[1].style
    expect(style).toContain(`width:min(200px,${paper})`)
    // auto margins resolve to 0 on overflow: symmetric spill needs the negative-capable calc
    expect(style).toContain(`margin-left:calc((var(--doc-content-w,100%) - min(200px,${paper}))/2)`)
    const nestedStyle = (renderTableSpec(centered, true) as Spec)[1].style
    expect(nestedStyle).toContain('width:min(200px,100%)')
    expect(nestedStyle).toContain('margin-left:auto;margin-right:auto')

    const pct: TableModel = {
      rows: [[cell('a'), cell('b')]],
      colWidthsTwips: [3000, 3000],
      colWidthsPct: [50, 50],
      widthPct: 80,
    }
    // a top-level 'pct' table is a share of its SECTION column (differing-margin
    // sections resolve through --doc-content-w); nested tables stay cell-relative
    const pctStyle = (renderTableSpec(pct) as Spec)[1].style
    expect(pctStyle).toContain('width:calc(var(--doc-content-w,100%) * 0.8)')
    expect(pctStyle).not.toContain('width:80%')
    expect((renderTableSpec(pct, true) as Spec)[1].style).toContain('width:80%')
  })
})

describe('nested table width solving', () => {
  // 8px per character, so 'w' x 20 + 2px edge = 162px -> ceil(162 x 1.02 x 15) + 216 = 2695 twips
  const perChar = {
    measure: (t: string) => t.length * 8,
    metrics: () => ({ ascent: 0, descent: 0, lineHeight: 0 }),
  }
  const LONG = 2695
  const X = Math.ceil(10 * 1.02 * 15) + 216
  const deep: TableModel = {
    rows: [[cell('w'.repeat(20)), cell('x')]],
    colWidthsTwips: [335, 335],
    autoLayout: true,
  }
  // dxa tblW (no autoLayout flag): its tcW split the cell evenly
  const inner: TableModel = {
    rows: [[{ paras: [''], nestedTables: [deep] }, cell('x')]],
    colWidthsTwips: [2016, 2017],
  }
  const outer: TableModel = {
    rows: [[{ paras: [''], nestedTables: [inner] }, cell('x')]],
    colWidthsTwips: [4259, 4255],
    autoLayout: true,
  }

  it('grows a nested dxa table column to hold the deeper grid, shrinking the rest to the cell', () => {
    const expanded = expandAutofitColWidths(outer, 8640, 8640, perChar)
    expect(expanded.colWidthsTwips).toEqual([4259, 4255])
    const mid = expanded.rows[0][0].nestedTables![0]
    // deep needs LONG + X; the column holding it adds the cell padding; the
    // total stays at the outer cell content width 4259 - 216 (no hang)
    expect(mid.colWidthsTwips).toEqual([LONG + X + 216, 4043 - (LONG + X + 216)])
    expect(mid.rows[0][0].nestedTables![0].colWidthsTwips).toEqual([LONG, X])
    expect(inner.colWidthsTwips).toEqual([2016, 2017])
  })

  it('shrinks an empty nested column to its side margins, never a sliver below them', () => {
    // Word (eight nesting levels): a nested tcW 360/360 grid whose
    // second cell is empty lays out at 222/328 inside a 544-twip cell, so the
    // host column reserves only the margins for it; a 100-twip sliver keeps its width
    const deepEmpty: TableModel = { ...deep, rows: [[cell('w'.repeat(20)), cell('')]] }
    const hostFor = (deepest: TableModel): TableModel => ({
      ...outer,
      rows: [
        [
          {
            paras: [''],
            nestedTables: [
              { ...inner, rows: [[{ paras: [''], nestedTables: [deepest] }, cell('x')]] },
            ],
          },
          cell('x'),
        ],
      ],
    })
    const mid = expandAutofitColWidths(hostFor(deepEmpty), 8640, 8640, perChar).rows[0][0]
      .nestedTables![0]
    expect(mid.colWidthsTwips![0]).toBe(LONG + 216 + 216)
    expect(mid.rows[0][0].nestedTables![0].colWidthsTwips).toEqual([LONG, 216])

    const sliver: TableModel = { ...deepEmpty, colWidthsTwips: [335, 100] }
    const midSliver = expandAutofitColWidths(hostFor(sliver), 8640, 8640, perChar).rows[0][0]
      .nestedTables![0]
    expect(midSliver.colWidthsTwips![0]).toBe(LONG + 100 + 216)
    expect(midSliver.rows[0][0].nestedTables![0].colWidthsTwips).toEqual([LONG, 100])
  })

  it('adds half the outer vertical borders to a nested table min width', () => {
    const bordered: TableModel = {
      ...deep,
      borders: {
        left: { style: 'single', szEighths: 4 },
        right: { style: 'single', szEighths: 4 },
      },
    }
    const host: TableModel = { ...outer, colWidthsTwips: [1000, 7514] }
    host.rows = [[{ paras: [''], nestedTables: [bordered] }, cell('x')]]
    // 0.5pt borders straddled on both sides = 10 twips on top of the grid + margins
    expect(expandAutofitColWidths(host, 8640, 8640, perChar).colWidthsTwips![0]).toBe(
      LONG + X + 216 + 10,
    )
  })

  it('counts a nested table toward the outer column min-content', () => {
    const tight: TableModel = { ...outer, colWidthsTwips: [1000, 7514] }
    const expanded = expandAutofitColWidths(tight, 8640, 8640, perChar)
    const innerMin = LONG + X + 216 + X
    expect(expanded.colWidthsTwips).toEqual([innerMin + 216, 8640 - innerMin - 216])
  })

  it('leaves fixed-layout and pct nested tables to their own rules', () => {
    const fixedInner: TableModel = { ...inner, fixedLayout: true, colWidthsTwips: [3000, 3000] }
    const host: TableModel = {
      ...outer,
      rows: [[{ paras: [''], nestedTables: [fixedInner] }, cell('x')]],
      colWidthsTwips: [1000, 7514],
    }
    const expanded = expandAutofitColWidths(host, 8640, 8640, perChar)
    expect(expanded.colWidthsTwips![0]).toBe(6000 + 216)
    expect(expanded.rows[0][0].nestedTables![0].colWidthsTwips).toEqual([3000, 3000])
    const pctInner: TableModel = { ...inner, widthPct: 50, colWidthsPct: [50, 50] }
    const pctHost: TableModel = {
      ...outer,
      rows: [[{ paras: [''], nestedTables: [pctInner] }, cell('x')]],
      colWidthsTwips: [1000, 7514],
    }
    expect(expandAutofitColWidths(pctHost, 8640, 8640, perChar).colWidthsTwips![0]).toBe(1000)
  })
})

describe('legacyIndentTable', () => {
  // compatibilityMode < 15: tblInd is measured to the cell text, so the border
  // sits one left cell margin outside it (a regression sample: no tblInd, TableNormal
  // 108 -> table edge 7.2px left of the margin at 96dpi; 024: tblInd 108 -> on it)
  it('moves the border edge one left cell margin out', () => {
    const noInd: TableModel = { rows: [[cell('a')]], colWidthsTwips: [4000] }
    expect(legacyIndentTable(noInd).indentTwips).toBe(-108)
    const onMargin: TableModel = { ...noInd, indentTwips: 108 }
    expect(legacyIndentTable(onMargin).indentTwips).toBe(0)
    const custom: TableModel = { ...noInd, indentTwips: 32, cellMarTwips: { left: 24 } }
    expect(legacyIndentTable(custom).indentTwips).toBe(8)
  })

  it('leaves centered, right-aligned and floating tables alone', () => {
    const centered: TableModel = { rows: [[cell('a')]], align: 'center' }
    expect(legacyIndentTable(centered)).toBe(centered)
    const floating: TableModel = { rows: [[cell('a')]], floatSide: 'left' }
    expect(legacyIndentTable(floating)).toBe(floating)
  })

  // a host without settings.xml (compat 0) whose table comes from
  // an HTML altChunk; Word lays the chunk out by its own modern conventions
  it('does not apply to tables expanded from a w:altChunk', () => {
    const table: TableModel = { rows: [[cell('a')]], colWidthsTwips: [4000], indentTwips: 108 }
    const block = (altChunk: boolean): Block => ({
      id: 'b0',
      type: 'table',
      docxIndex: 0,
      originalXml: null,
      table,
      ...(altChunk ? { altChunk } : {}),
    })
    const indentOf = (b: Block): number | null =>
      blocksToPmDoc([b], undefined, { legacyTableIndent: true }).content![0].attrs!.indentTwips as
        number | null
    expect(indentOf(block(false))).toBe(0)
    expect(indentOf(block(true))).toBe(108)
  })
})

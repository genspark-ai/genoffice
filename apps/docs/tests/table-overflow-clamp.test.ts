import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, type TableModel } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  clampTableColWidths,
  expandAutofitColWidths,
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

  it('leaves fixed-layout tables and already-wide autofit columns alone', () => {
    const fixed: TableModel = {
      rows: [[cell(ARABIC_HEADER)]],
      colWidthsTwips: [567],
    }
    expect(expandAutofitColWidths(fixed, 10772, 9638)).toBe(fixed)

    const wide: TableModel = {
      rows: [[cell(ARABIC_HEADER), cell('ok')]],
      colWidthsTwips: [2000, 4000],
      autoLayout: true,
    }
    expect(expandAutofitColWidths(wide, 10772, 9638)).toBe(wide)
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

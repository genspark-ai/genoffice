/** rebuild-xlsx unit tests (P26): hand-built IR pages, verified by unzipping. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { Rect } from '../src/geometry'
import type { IrPage, Line, Span, TableBlock, TableCellBlock, TextBlock } from '../src/ir'
import { parseCellValue } from '../src/rebuild-xlsx/numbers'
import { ptToColumnChars, rebuildXlsx } from '../src/rebuild-xlsx/rebuild'

const span = (text: string, over: Partial<Span> = {}): Span => ({
  text,
  box: { x0: 0, y0: 0, x1: 10, y1: 10 },
  fontSize: 11,
  fontFamily: 'Helvetica',
  bold: false,
  italic: false,
  color: '000000',
  dir: 'ltr',
  script: 'latin',
  ...over,
})

const line = (text: string, box: Rect, over: Partial<Line> = {}): Line => ({
  spans: [span(text, { box })],
  box,
  baseline: box.y0 + 2,
  endsWithHyphen: false,
  ...over,
})

const textBlock = (text: string, box: Rect, over: Partial<TextBlock> = {}): TextBlock => ({
  kind: 'text',
  lines: [line(text, box)],
  box,
  align: 'left',
  firstLineIndentPt: 0,
  dir: 'ltr',
  ...over,
})

const cell = (text: string, box: Rect, over: Partial<TableCellBlock> = {}): TableCellBlock => ({
  box,
  gridSpan: 1,
  blocks: text === '' ? [] : [textBlock(text, box)],
  ...over,
})

const page = (over: Partial<IrPage> = {}): IrPage => ({
  index: 0,
  widthPt: 612,
  heightPt: 792,
  rotation: 0,
  blocks: [],
  degraded: false,
  scanned: false,
  hasStructTree: false,
  ...over,
})

async function unzip(xlsx: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(xlsx)
  const out = new Map<string, string>()
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir) out.set(path, await file.async('string'))
  }
  return out
}

// ── number classification matrix ──

describe('parseCellValue', () => {
  const num = (v: number, numFmt = 'General') => ({ kind: 'number', value: v, numFmt })

  it('parses plain and grouped numbers', () => {
    expect(parseCellValue('1234')).toEqual(num(1234))
    expect(parseCellValue('-3.5')).toEqual(num(-3.5))
    expect(parseCellValue('0')).toEqual(num(0))
    expect(parseCellValue('0.25')).toEqual(num(0.25))
    // authored trailing zeros keep a fixed-decimal display format
    expect(parseCellValue('31.10')).toEqual(num(31.1, '0.00'))
    expect(parseCellValue('2.500')).toEqual(num(2.5, '0.000'))
    expect(parseCellValue('1,234')).toEqual(num(1234, '#,##0'))
    expect(parseCellValue('1,234.56')).toEqual(num(1234.56, '#,##0.00'))
    expect(parseCellValue('12,345,678')).toEqual(num(12345678, '#,##0'))
    expect(parseCellValue(' 42 ')).toEqual(num(42))
  })

  it('parses percentages as fractions with percent formats', () => {
    expect(parseCellValue('45%')).toEqual(num(0.45, '0%'))
    expect(parseCellValue('45.3%')).toEqual(num(0.453, '0.0%'))
    expect(parseCellValue('-2.75%')).toEqual(num(-0.0275, '0.00%'))
    expect(parseCellValue('100%')).toEqual(num(1, '0%'))
    expect(parseCellValue('1,200%')).toEqual(num(12, '0%'))
  })

  it('parses currency prefixes and suffixes', () => {
    expect(parseCellValue('$1,200')).toEqual(num(1200, '"$"#,##0'))
    expect(parseCellValue('$1,200.50')).toEqual(num(1200.5, '"$"#,##0.00'))
    expect(parseCellValue('-$45')).toEqual(num(-45, '"$"#,##0'))
    expect(parseCellValue('$-45')).toEqual(num(-45, '"$"#,##0'))
    expect(parseCellValue('-$-45').kind).toBe('text')
    expect(parseCellValue('¥500')).toEqual(num(500, '"¥"#,##0'))
    expect(parseCellValue('￥500')).toEqual(num(500, '"¥"#,##0'))
    expect(parseCellValue('€99.99')).toEqual(num(99.99, '"€"#,##0.00'))
    expect(parseCellValue('£10')).toEqual(num(10, '"£"#,##0'))
    expect(parseCellValue('1,200元')).toEqual(num(1200, '#,##0"元"'))
    expect(parseCellValue('3.50 元')).toEqual(num(3.5, '#,##0.00"元"'))
  })

  it('keeps leading zeros, phone-like and date-like strings as text', () => {
    for (const text of [
      '007',
      '0042',
      '00.5',
      '013800138000',
      '+86 138 0013 8000',
      '138-0013-8000',
      '2024-01-05',
      '1/2/2024',
      '2024/01/05',
      '01.02.2024',
      '12:30',
      '1.2.3',
      '1,23',
      '1,2345',
      ',123',
      '9781234567890123456',
      'abc',
      '12 34',
      '12%%',
      '$',
      '元',
      '',
      '  ',
      'NaN',
      'Infinity',
    ]) {
      expect(parseCellValue(text).kind, text).toBe('text')
    }
  })

  it('keeps huge integers as text but accepts exponent notation', () => {
    expect(parseCellValue('123456789012345678').kind).toBe('text')
    expect(parseCellValue('1e5')).toEqual(num(100000))
    expect(parseCellValue('1.5E-3')).toEqual(num(0.0015))
  })
})

// ── column width conversion ──

describe('ptToColumnChars', () => {
  it('converts pt widths to Excel character units', () => {
    // 96 pt ≈ 128 px ≈ (128−5)/7 ≈ 17.57 chars
    expect(ptToColumnChars(96)).toBeCloseTo(17.57, 1)
    // never below the readable floor
    expect(ptToColumnChars(1)).toBe(2)
  })
})

// ── workbook assembly ──

const GRID: Rect = { x0: 72, y0: 600, x1: 272, y1: 640 }

/** 2×2 lattice with a numeric column and a vMerge in column A */
function sampleTable(over: Partial<TableBlock> = {}): TableBlock {
  const rowTop = { y0: 620, y1: 640 }
  const rowBottom = { y0: 600, y1: 620 }
  return {
    kind: 'table',
    box: GRID,
    colWidthsPt: [100, 100],
    rows: [
      [
        cell('Item', { x0: 72, x1: 172, ...rowTop }, { vMerge: 'restart', fill: 'ddeeff' }),
        cell('1,234.56', { x0: 172, x1: 272, ...rowTop }),
      ],
      [
        cell('', { x0: 72, x1: 172, ...rowBottom }, { vMerge: 'continue' }),
        cell('45%', { x0: 172, x1: 272, ...rowBottom }, { vAlign: 'center' }),
      ],
    ],
    ...over,
  }
}

describe('rebuildXlsx', () => {
  it('emits one worksheet per page named Page N', async () => {
    const { xlsx } = await rebuildXlsx([page(), page({ index: 1 })])
    const parts = await unzip(xlsx)
    const workbook = parts.get('xl/workbook.xml')!
    expect(workbook).toContain('name="Page 1"')
    expect(workbook).toContain('name="Page 2"')
    expect(parts.has('xl/worksheets/sheet1.xml')).toBe(true)
    expect(parts.has('xl/worksheets/sheet2.xml')).toBe(true)
    expect(parts.get('[Content_Types].xml')).toContain('/xl/worksheets/sheet2.xml')
  })

  it('carries dropped furniture as the print header/footer on every sheet', async () => {
    const hf = {
      pageNo: false,
      fontSizePt: 9,
      fontFamily: 'Helvetica',
      bold: false,
      italic: false,
      color: '000000',
      coversFirstPage: true,
    }
    const { xlsx } = await rebuildXlsx(
      [page(), page({ index: 1 })],
      [
        { ...hf, band: 'top' as const, text: 'ACME & Co', x0: 240, x1: 360, edgeDistPt: 20 },
        {
          ...hf,
          band: 'bottom' as const,
          text: 'Page \uE001',
          pageNo: true,
          x0: 500,
          x1: 560,
          edgeDistPt: 20,
        },
      ],
    )
    const parts = await unzip(xlsx)
    for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) {
      const sheet = parts.get(name)!
      // centered ink → &C, ampersand doubled; page-number mark → &P at &R
      expect(sheet).toContain('<oddHeader>&amp;CACME &amp;&amp; Co</oddHeader>')
      expect(sheet).toContain('<oddFooter>&amp;RPage &amp;P</oddFooter>')
    }
    // &P restarts per worksheet — each sheet pins its printed page number
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain(
      '<pageSetup paperSize="1" orientation="portrait" firstPageNumber="1" useFirstPageNumber="1"/>',
    )
    expect(parts.get('xl/worksheets/sheet2.xml')).toContain(
      '<pageSetup paperSize="1" orientation="portrait" firstPageNumber="2" useFirstPageNumber="1"/>',
    )
  })

  it('orders furniture lines by edge distance and skips cover pages', async () => {
    const hf = {
      pageNo: false,
      fontSizePt: 9,
      fontFamily: 'Helvetica',
      bold: false,
      italic: false,
      color: '000000',
      x0: 200,
      x1: 400,
    }
    const { xlsx } = await rebuildXlsx(
      [page(), page({ index: 1 })],
      [
        // detection order reversed vs reading order: the footer's UPPER line
        // sits FURTHER from the bottom edge
        { ...hf, band: 'bottom' as const, text: 'closing', edgeDistPt: 20, coversFirstPage: false },
        { ...hf, band: 'bottom' as const, text: 'opening', edgeDistPt: 32, coversFirstPage: false },
      ],
    )
    const parts = await unzip(xlsx)
    expect(parts.get('xl/worksheets/sheet2.xml')).toContain(
      '<oddFooter>&amp;Copening\nclosing</oddFooter>',
    )
    // no slot covers the first (cover) page — sheet 1 stays clean
    expect(parts.get('xl/worksheets/sheet1.xml')).not.toContain('<headerFooter>')
  })

  it('maps tables to grids with merges, numbers, fills and borders', async () => {
    const { xlsx, warnings } = await rebuildXlsx([page({ blocks: [sampleTable()] })])
    const parts = await unzip(xlsx)
    const sheet = parts.get('xl/worksheets/sheet1.xml')!
    // vMerge restart in col A spans two rows
    expect(sheet).toContain('<mergeCell ref="A1:A2"/>')
    // grouped number became numeric with its own style
    expect(sheet).toMatch(/<c r="B1" s="\d+"><v>1234.56<\/v><\/c>/)
    // percent stored as fraction
    expect(sheet).toMatch(/<c r="B2" s="\d+"><v>0.45<\/v><\/c>/)
    // text cell stays an inline string
    expect(sheet).toContain('<is><t xml:space="preserve">Item</t></is>')
    // measured row heights ride along
    expect(sheet).toContain('customHeight="1"')
    // column widths from colWidthsPt
    // 100 pt → (100 × 4/3 − 5) / 7 ≈ 18.33 chars
    expect(sheet).toMatch(/<col min="1" max="1" width="18.33" customWidth="1"\/>/)

    const styles = parts.get('xl/styles.xml')!
    expect(styles).toContain('patternFill patternType="solid"')
    expect(styles).toContain('FFDDEEFF')
    // lattice ⇒ thin borders on all edges
    expect(styles).toContain('<left style="thin">')
    // builtin percent format id 9 needs no custom numFmt, grouped needs id 4
    expect(styles).toContain('numFmtId="9"')
    expect(styles).toContain('numFmtId="4"')
    expect(warnings).toEqual([])
  })

  it('keeps stream tables borderless and leading zeros as text', async () => {
    const table = sampleTable({ confidence: 0.9 })
    table.rows[0]![1] = cell('007', { x0: 172, x1: 272, y0: 620, y1: 640 })
    const { xlsx } = await rebuildXlsx([page({ blocks: [table] })])
    const parts = await unzip(xlsx)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain(
      '<is><t xml:space="preserve">007</t></is>',
    )
    expect(parts.get('xl/styles.xml')).not.toContain('style="thin"')
  })

  it('separates adjacent tables by one blank row and keeps text rows in column A', async () => {
    const p = page({
      blocks: [
        textBlock('Intro paragraph', { x0: 72, y0: 700, x1: 300, y1: 715 }),
        sampleTable(),
        sampleTable(),
      ],
    })
    const { xlsx } = await rebuildXlsx([p])
    const parts = await unzip(xlsx)
    const sheet = parts.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toContain('<is><t xml:space="preserve">Intro paragraph</t></is>')
    // text row 1, table rows 2–3, spacer row 4, second table rows 5–6
    expect(sheet).toContain('<mergeCell ref="A2:A3"/>')
    expect(sheet).toContain('<mergeCell ref="A5:A6"/>')
    expect(sheet).not.toMatch(/<row r="4"/)
  })

  it('emits a gridSpan-only merge without vMerge', async () => {
    const table = sampleTable()
    table.rows[0] = [cell('Wide header', { x0: 72, x1: 272, y0: 620, y1: 640 }, { gridSpan: 2 })]
    table.rows[1] = [
      cell('a', { x0: 72, x1: 172, y0: 600, y1: 620 }),
      cell('b', { x0: 172, x1: 272, y0: 600, y1: 620 }),
    ]
    const { xlsx } = await rebuildXlsx([page({ blocks: [table] })])
    const parts = await unzip(xlsx)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain('<mergeCell ref="A1:B1"/>')
  })

  it('adds the no-tables warning on text-only documents', async () => {
    const p = page({ blocks: [textBlock('Just prose', { x0: 72, y0: 700, x1: 300, y1: 715 })] })
    const { xlsx, warnings } = await rebuildXlsx([p])
    expect(warnings).toContain('no tables detected')
    const parts = await unzip(xlsx)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain('Just prose')
  })

  it('writes a single notice row for scanned pages', async () => {
    const { xlsx, warnings } = await rebuildXlsx([page({ scanned: true })])
    const parts = await unzip(xlsx)
    expect(parts.get('xl/worksheets/sheet1.xml')).toContain(
      'Page 1: scanned page — not convertible to cells',
    )
    // the xlsx layer itself adds nothing (pipeline already warned)
    expect(warnings).toEqual([])
  })

  it('survives an empty document', async () => {
    const { xlsx } = await rebuildXlsx([])
    const parts = await unzip(xlsx)
    expect(parts.get('xl/workbook.xml')).toContain('name="Page 1"')
  })
})

describe('P27 x-aware sheet columns', () => {
  it('places an indented table beside margin prose in its own columns', async () => {
    const prose = textBlock('Body paragraph at the left margin.', {
      x0: 56,
      x1: 400,
      y0: 700,
      y1: 715,
    })
    const table = sampleTable({ box: { x0: 210, x1: 410, y0: 600, y1: 640 } })
    for (const row of table.rows) {
      row[0]!.box = { ...row[0]!.box, x0: 210, x1: 310 }
      row[1]!.box = { ...row[1]!.box, x0: 310, x1: 410 }
    }
    const { xlsx } = await rebuildXlsx([page({ blocks: [prose, table] })])
    const parts = await unzip(xlsx)
    const sheet = parts.get('xl/worksheets/sheet1.xml')!
    // prose at slot 0 (col A), table columns at slots 1-2 (B/C)
    expect(sheet).toContain('<c r="A1"')
    expect(sheet).toMatch(/<c r="B2" s="\d+" t="inlineStr"><is><t xml:space="preserve">Item<\/t>/)
    expect(sheet).toMatch(/<c r="C2" s="\d+"><v>1234.56<\/v>/)
    expect(sheet).toContain('<mergeCell ref="B2:B3"/>')
  })

  it('keeps a margin-aligned table in column A (baseline layout unchanged)', async () => {
    const prose = textBlock('Body paragraph.', { x0: 72, x1: 400, y0: 700, y1: 715 })
    const table = sampleTable() // box.x0 = 72 = prose margin
    const { xlsx } = await rebuildXlsx([page({ blocks: [prose, table] })])
    const parts = await unzip(xlsx)
    const sheet = parts.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(/<c r="A2" s="\d+" t="inlineStr"><is><t xml:space="preserve">Item<\/t>/)
    expect(sheet).toMatch(/<c r="B2" s="\d+"><v>1234.56<\/v>/)
  })

  it('suppresses softEdges borders on split-run lattice cells', async () => {
    const table = sampleTable()
    table.rows[0]![1] = cell(
      '1,234.56',
      { x0: 172, x1: 272, y0: 620, y1: 640 },
      {
        softEdges: { left: true },
      },
    )
    const { xlsx } = await rebuildXlsx([page({ blocks: [table] })])
    const parts = await unzip(xlsx)
    const styles = parts.get('xl/styles.xml')!
    // at least one border record omits the left edge while keeping the rest
    expect(styles).toMatch(/<border><left\/><right style="thin">/)
  })
})

// ── cross-page table runs (P39) ──

/** header + one glued band row, statement style */
function statementTable(band: string, over: Partial<TableBlock> = {}): TableBlock {
  const head = { y0: 700, y1: 720 }
  const body = { y0: 100, y1: 700 }
  return {
    kind: 'table',
    box: { x0: 40, y0: 100, x1: 240, y1: 720 },
    colWidthsPt: [100, 100],
    rows: [
      [cell('Ref', { x0: 40, x1: 140, ...head }), cell('Amount', { x0: 140, x1: 240, ...head })],
      [cell(band, { x0: 40, x1: 140, ...body }), cell('9.99', { x0: 140, x1: 240, ...body })],
    ],
    ...over,
  }
}

describe('cross-page table runs (P39)', () => {
  it('merges continuing statement pages into one sheet and drops the restated header', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const p2 = page({ index: 1, blocks: [statementTable('RC-2')] })
    const { xlsx, sheets } = await rebuildXlsx([p1, p2])
    expect(sheets.map((s) => s.name)).toEqual(['Pages 1-2'])
    const parts = await unzip(xlsx)
    expect(parts.get('xl/workbook.xml')).toContain('name="Pages 1-2"')
    const sheet1 = parts.get('xl/worksheets/sheet1.xml')!
    expect(sheet1.match(/>Ref</g)).toHaveLength(1) // header once
    expect(sheet1).toContain('>RC-1<')
    expect(sheet1).toContain('>RC-2<')
  })

  it('keeps pages separate when the header differs', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const other = statementTable('RC-2')
    other.rows[0] = [
      cell('Code', { x0: 40, x1: 140, y0: 700, y1: 720 }),
      cell('Total', { x0: 140, x1: 240, y0: 700, y1: 720 }),
    ]
    const p2 = page({ index: 1, blocks: [other] })
    const { sheets } = await rebuildXlsx([p1, p2])
    expect(sheets.map((s) => s.name)).toEqual(['Page 1', 'Page 2'])
  })

  it('keeps pages separate when column boundaries move', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const shifted = statementTable('RC-2', {
      box: { x0: 60, y0: 100, x1: 260, y1: 720 },
    })
    shifted.rows = shifted.rows.map((r) =>
      r.map((c) => ({ ...c, box: { ...c.box, x0: c.box.x0 + 20, x1: c.box.x1 + 20 } })),
    )
    const p2 = page({ index: 1, blocks: [shifted] })
    const { sheets } = await rebuildXlsx([p1, p2])
    expect(sheets.map((s) => s.name)).toEqual(['Page 1', 'Page 2'])
  })

  it('a scanned page breaks the run', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const p2 = page({ index: 1, scanned: true })
    const p3 = page({ index: 2, blocks: [statementTable('RC-3')] })
    const { sheets } = await rebuildXlsx([p1, p2, p3])
    expect(sheets.map((s) => s.name)).toEqual(['Page 1', 'Page 2', 'Page 3'])
  })

  it('a mid-document new statement starts a new sheet (cover table first)', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const p2 = page({ index: 1, blocks: [statementTable('RC-2')] })
    const cover: TableBlock = {
      kind: 'table',
      box: { x0: 300, y0: 730, x1: 500, y1: 780 },
      colWidthsPt: [200],
      rows: [[cell('RIB box', { x0: 300, x1: 500, y0: 730, y1: 780 })]],
    }
    const p3 = page({ index: 2, blocks: [cover, statementTable('RC-3')] })
    const { sheets } = await rebuildXlsx([p1, p2, p3])
    expect(sheets.map((s) => s.name)).toEqual(['Pages 1-2', 'Page 3'])
  })

  it('continuation rows stay in the anchor columns under sub-tolerance drift', async () => {
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    // whole grid drifts +1.8pt: within RUN_GEOM_TOL_PT but past SLOT_TOL_PT
    const drifted = statementTable('RC-2', {
      box: { x0: 41.8, y0: 100, x1: 241.8, y1: 720 },
    })
    drifted.rows = drifted.rows.map((r) =>
      r.map((c) => ({ ...c, box: { ...c.box, x0: c.box.x0 + 1.8, x1: c.box.x1 + 1.8 } })),
    )
    const p2 = page({ index: 1, blocks: [drifted] })
    const { sheets } = await rebuildXlsx([p1, p2])
    expect(sheets.map((s) => s.name)).toEqual(['Pages 1-2'])
    const colOf = (needle: string): number =>
      sheets[0]!.cells.find((c) => c.value?.kind === 'text' && c.value.text === needle)!.col
    expect(colOf('RC-2')).toBe(colOf('RC-1'))
  })

  it('a run starting mid-group anchors to its own table, not the group head', async () => {
    // p1: table A · p2: A continues, then NEW table B starts · p3: B continues
    // with sub-tolerance drift — B's rows must map through B's bounds, not A's
    const tableB = (band: string, dx = 0): TableBlock => ({
      kind: 'table',
      box: { x0: 300 + dx, y0: 100, x1: 540 + dx, y1: 400 },
      colWidthsPt: [80, 80, 80],
      rows: [
        [
          cell('Code', { x0: 300 + dx, x1: 380 + dx, y0: 380, y1: 400 }),
          cell('Qty', { x0: 380 + dx, x1: 460 + dx, y0: 380, y1: 400 }),
          cell('Total', { x0: 460 + dx, x1: 540 + dx, y0: 380, y1: 400 }),
        ],
        [
          cell(band, { x0: 300 + dx, x1: 380 + dx, y0: 100, y1: 380 }),
          cell('7', { x0: 380 + dx, x1: 460 + dx, y0: 100, y1: 380 }),
          cell('9.99', { x0: 460 + dx, x1: 540 + dx, y0: 100, y1: 380 }),
        ],
      ],
    })
    const p1 = page({ index: 0, blocks: [statementTable('RC-1')] })
    const p2 = page({ index: 1, blocks: [statementTable('RC-2'), tableB('B-1')] })
    const p3 = page({ index: 2, blocks: [tableB('B-2', 1.8)] })
    const { sheets } = await rebuildXlsx([p1, p2, p3])
    expect(sheets.map((s) => s.name)).toEqual(['Pages 1-3'])
    const colOf = (needle: string): number =>
      sheets[0]!.cells.find((c) => c.value?.kind === 'text' && c.value.text === needle)!.col
    expect(colOf('B-2')).toBe(colOf('B-1'))
    expect(colOf('RC-2')).toBe(colOf('RC-1'))
  })
})
// ── band-row splitting (P40) ──

describe('splitBandRows via rebuildXlsx (P40)', () => {
  /** statement band: header + one giant row; per txn the ref/amount sit on
   * the boundary line and the description adds a second line below it */
  const Y = (i: number) => 700 - i * 40 // txn i boundary top (IR y up)
  const bandTable = (txns: number): TableBlock => {
    const mk = (
      text: string,
      x0: number,
      x1: number,
      top: number,
      h = 12,
    ): { line: Line; box: Rect } => {
      const box = { x0, x1, y0: top - h, y1: top }
      return { line: line(text, box), box }
    }
    const refLines: Line[] = []
    const descLines: Line[] = []
    const amtLines: Line[] = []
    for (let i = 0; i < txns; i++) {
      refLines.push(mk(`R-${i}`, 40, 100, Y(i)).line)
      descLines.push(mk(`LABEL-${i}`, 140, 220, Y(i)).line)
      descLines.push(mk(`DETAIL-${i}`, 140, 260, Y(i) - 14).line)
      amtLines.push(mk(`${i + 1},000.00`, 300, 380, Y(i)).line)
    }
    const bandBox = { y0: Y(txns - 1) - 30, y1: 712 }
    const blockOf = (lines: Line[], x0: number, x1: number): TextBlock => ({
      kind: 'text',
      lines,
      box: { x0, x1, ...bandBox },
      align: 'left',
      firstLineIndentPt: 0,
      dir: 'ltr',
    })
    const cellOf = (lines: Line[], x0: number, x1: number): TableCellBlock => ({
      box: { x0, x1, ...bandBox },
      gridSpan: 1,
      blocks: [blockOf(lines, x0, x1)],
    })
    const head = { y0: 715, y1: 730 }
    return {
      kind: 'table',
      box: { x0: 40, y0: bandBox.y0, x1: 380, y1: 730 },
      colWidthsPt: [100, 160, 80],
      rows: [
        [
          cell('Ref', { x0: 40, x1: 140, ...head }),
          cell('Desc', { x0: 140, x1: 300, ...head }),
          cell('Amount', { x0: 300, x1: 380, ...head }),
        ],
        [cellOf(refLines, 40, 140), cellOf(descLines, 140, 300), cellOf(amtLines, 300, 380)],
      ],
    }
  }

  it('splits a glued statement band into one row per transaction', async () => {
    const { sheets } = await rebuildXlsx([page({ blocks: [bandTable(4)] })])
    const cells = sheets[0]!.cells.filter((c) => c.value !== undefined)
    const textAt = (needle: string) =>
      cells.find(
        (c) => c.value!.kind === 'text' && (c.value as { text: string }).text.includes(needle),
      )
    for (let i = 0; i < 4; i++) {
      const ref = textAt(`R-${i}`)!
      const desc = textAt(`LABEL-${i}`)!
      expect(desc.row).toBe(ref.row)
      // the wrapped detail line stays with its own transaction
      expect((desc.value as { text: string }).text).toContain(`DETAIL-${i}`)
      const amt = cells.find((c) => c.value!.kind === 'number' && c.row === ref.row)!
      expect((amt.value as { value: number }).value).toBe(1000 * (i + 1))
    }
  })

  it('keeps two-line wrapped cells intact (no split below the line floor)', async () => {
    const { sheets } = await rebuildXlsx([page({ blocks: [bandTable(2)] })])
    // 2 txns → band cells carry 2-4 lines but only 2 boundaries < MIN_SPLIT_ROWS
    const texts = sheets[0]!.cells
      .filter((c) => c.value?.kind === 'text')
      .map((c) => (c.value as { text: string }).text)
    expect(texts.some((t) => t.includes('R-0') && t.includes('R-1'))).toBe(true)
  })

  it('never splits rows containing merged cells', async () => {
    const t = bandTable(4)
    t.rows[1]![0]!.vMerge = 'restart'
    const { sheets } = await rebuildXlsx([page({ blocks: [t] })])
    const texts = sheets[0]!.cells
      .filter((c) => c.value?.kind === 'text')
      .map((c) => (c.value as { text: string }).text)
    expect(texts.some((t2) => t2.includes('R-0') && t2.includes('R-3'))).toBe(true)
  })

  it('staggered column baselines collapse into one row (no phantom 3pt rows)', async () => {
    // two column groups start each transaction 3pt apart — both clusters can
    // win the vote, but boundaries closer than the row pitch must merge
    const mk = (text: string, x0: number, x1: number, top: number): Line =>
      line(text, { x0, x1, y0: top - 10, y1: top })
    const colCell = (lines: Line[], x0: number, x1: number): TableCellBlock => ({
      box: { x0, x1, y0: 480, y1: 700 },
      gridSpan: 1,
      blocks: [
        {
          kind: 'text',
          lines,
          box: { x0, x1, y0: 480, y1: 700 },
          align: 'left',
          firstLineIndentPt: 0,
          dir: 'ltr',
        },
      ],
    })
    const Y = (i: number) => 690 - i * 40
    const a: Line[] = []
    const b: Line[] = []
    const c: Line[] = []
    const d: Line[] = []
    for (let i = 0; i < 4; i++) {
      a.push(mk(`A-${i}`, 40, 90, Y(i)))
      b.push(mk(`B-${i}`, 100, 150, Y(i)))
      c.push(mk(`C-${i}`, 160, 210, Y(i) - 3))
      d.push(mk(`D-${i}`, 220, 270, Y(i) - 3))
    }
    const head = { y0: 705, y1: 718 }
    const t: TableBlock = {
      kind: 'table',
      box: { x0: 40, y0: 480, x1: 280, y1: 718 },
      colWidthsPt: [60, 60, 60, 60],
      rows: [
        [
          cell('H1', { x0: 40, x1: 100, ...head }),
          cell('H2', { x0: 100, x1: 160, ...head }),
          cell('H3', { x0: 160, x1: 220, ...head }),
          cell('H4', { x0: 220, x1: 280, ...head }),
        ],
        [colCell(a, 40, 100), colCell(b, 100, 160), colCell(c, 160, 220), colCell(d, 220, 280)],
      ],
    }
    const { sheets } = await rebuildXlsx([page({ blocks: [t] })])
    const cells = sheets[0]!.cells.filter((x) => x.value?.kind === 'text')
    const rowOfText = (needle: string) =>
      cells.find((x) => (x.value as { text: string }).text.includes(needle))!.row
    for (let i = 0; i < 4; i++) {
      expect(rowOfText(`C-${i}`)).toBe(rowOfText(`A-${i}`))
      expect(rowOfText(`D-${i}`)).toBe(rowOfText(`A-${i}`))
    }
  })
})

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

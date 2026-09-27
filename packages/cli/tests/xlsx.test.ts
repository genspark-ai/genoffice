import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { RECALC_CELL_CAP, displayText, isoDateSerial, recalcBands } from '../src/formats/xlsx'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file(name)!.async('string')
}

describe('genoffice create --type xlsx / sheet', () => {
  it('builds a workbook from a JSON table with formulas', async () => {
    const dir = tempDir()
    const table = join(dir, 'table.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['item', 'qty'],
        ['Apple', 2],
        ['Pear', 3],
        ['Total', '=SUM(B2:B3)'],
      ]),
    )
    const out = join(dir, 'table.xlsx')
    const r = await run(['create', '--type', 'xlsx', '--from', table, '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ sheets: 1, cells: 8, formulas: 1 })
    if (sidecar) expect(r.json().detail.cached_values).toBe(true)
    const sheet = await part(out, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<f>SUM(B2:B3)</f>')
    expect(sheet).toContain('Apple')
    if (sidecar) expect(sheet).toContain('<v>5</v>')
    expect(await part(out, 'xl/styles.xml')).toContain('<cellXfs')
  })

  it.skipIf(!sidecar)('reads back values, formulas and evaluated results', async () => {
    const dir = tempDir()
    const table = join(dir, 't.json')
    writeFileSync(table, JSON.stringify([[1, 2, '=A1+B1']]))
    const out = join(dir, 't.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
    const r = await run(['sheet', 'read', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.rows[0]).toEqual([1, 2, 3])
    expect(r.json().detail.formulas).toEqual({ C1: '=A1+B1' })
    const ranged = await run(['sheet', 'read', out, '--range', 'B1:C1', '--json'])
    expect(ranged.json().detail.rows[0]).toEqual([2, 3])
    expect(ranged.json().detail.range).toBe('B1:C1')
  })

  it('creates multi-sheet workbooks and imports csv with numeric cells', async () => {
    const dir = tempDir()
    const multi = join(dir, 'multi.json')
    writeFileSync(
      multi,
      JSON.stringify({
        sheets: [
          { name: 'Data', rows: [['a', 1]] },
          { name: 'Notes', rows: [['hello']] },
        ],
      }),
    )
    const out = join(dir, 'multi.xlsx')
    const r = await run(['create', '--type', 'xlsx', '--from', multi, '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.sheets).toBe(2)
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="Data"')
    expect(wb).toContain('name="Notes"')

    const csv = join(dir, 'sales.csv')
    writeFileSync(csv, 'region,amount\nEast,10.5\nWest,7\n')
    const csvOut = join(dir, 'sales.xlsx')
    const c = await run(['create', '--type', 'xlsx', '--from', csv, '--out', csvOut, '--json'])
    expect(c.code).toBe(0)
    expect(c.json().detail).toMatchObject({ sheets: 1, cells: 6 })
    expect(await part(csvOut, 'xl/workbook.xml')).toContain('name="sales"')
    expect(await part(csvOut, 'xl/worksheets/sheet1.xml')).toContain('<v>10.5</v>')
  })

  it('derives a valid sheet name from a file stem Excel would reject', async () => {
    const dir = tempDir()
    const csv = join(dir, "'data[final]:v2'.csv")
    writeFileSync(csv, 'a,b\n1,2\n')
    const out = join(dir, 'data.xlsx')
    const c = await run(['create', '--type', 'xlsx', '--from', csv, '--out', out, '--json'])
    expect(c.code).toBe(0)
    expect(await part(out, 'xl/workbook.xml')).toContain('name="data_final__v2"')
    const conv = await run([
      'convert',
      csv,
      '--to',
      'xlsx',
      '--out',
      join(dir, 'conv.xlsx'),
      '--json',
    ])
    expect(conv.code).toBe(0)
    expect(conv.json().detail.sheet).toBe('data_final__v2')
  })

  it('types ISO dates, honours sep= and --decimal, and --header freezes and filters', async () => {
    const dir = tempDir()
    const csv = join(dir, 'orders.csv')
    writeFileSync(csv, 'sep=;\nday;amount;code\n2024-01-01;1.234,5;007\n2024-01-02 09:30;7;x\n')
    const out = join(dir, 'orders.xlsx')
    const r = await run([
      'create',
      '--type',
      'xlsx',
      '--from',
      csv,
      '--out',
      out,
      '--decimal',
      ',',
      '--header',
      '--json',
    ])
    expect(r.code).toBe(0)
    const sheet = await part(out, 'xl/worksheets/sheet1.xml')
    expect(sheet).not.toContain('sep=')
    expect(sheet).toMatch(/<c r="A2" s="\d+"><v>45292<\/v><\/c>/)
    expect(sheet).toMatch(/<c r="A3" s="\d+"><v>45293\.39583333/)
    expect(sheet).toContain('<v>1234.5</v>')
    expect(sheet).toContain('007')
    expect(sheet).toMatch(/<pane [^>]*ySplit="1"/)
    expect(sheet).toContain('<autoFilter ref="A1:C3"')
    const styles = await part(out, 'xl/styles.xml')
    expect(styles).toContain('yyyy-mm-dd')
    expect(styles).toContain('yyyy-mm-dd hh:mm')
    expect(isoDateSerial('2024-02-30')).toBeUndefined()
    expect(isoDateSerial('2024-01-01T12:00:00')).toEqual({
      serial: 45292.5,
      format: 'yyyy-mm-dd hh:mm:ss',
    })
  })

  it.skipIf(!sidecar)(
    'sheet apply writes values, formulas and styles into an existing workbook',
    async () => {
      const dir = tempDir()
      const table = join(dir, 't.json')
      writeFileSync(table, JSON.stringify([['x', 'y']]))
      const out = join(dir, 't.xlsx')
      await run(['create', '--type', 'xlsx', '--from', table, '--out', out])
      const cells = join(dir, 'cells.json')
      writeFileSync(
        cells,
        JSON.stringify([
          { cell: 'A2', value: 10 },
          { cell: 'B2', value: 32 },
          { cell: 'C2', formula: 'A2+B2', style: { bold: true } },
          { cell: 'A1', style: { fillColor: '#FFFF00' } },
        ]),
      )
      const applied = await run(['sheet', 'apply', out, '--cells', cells, '--json'])
      expect(applied.code).toBe(0)
      expect(applied.json().detail).toMatchObject({ cells: 4, formulas: 1 })
      const read = await run(['sheet', 'read', out, '--json'])
      expect(read.json().detail.rows[1]).toEqual([10, 32, 42])
      expect(read.json().detail.rows[0][0]).toBe('x')
      const styles = await part(out, 'xl/styles.xml')
      expect(styles).toContain('FFFF00')
      expect(styles).toContain('<b/>')

      const bad = join(dir, 'bad.json')
      writeFileSync(bad, JSON.stringify([{ cell: 'A1', value: 1, sheet: 'Nope' }]))
      const r = await run(['sheet', 'apply', out, '--cells', bad, '--json'])
      expect(r.code).toBe(1)
      expect(r.json().detail.sheets).toEqual(['t'])
    },
  )

  it.skipIf(!sidecar)('writes theme colors and pattern fills and reads them back', async () => {
    const dir = tempDir()
    const table = join(dir, 't.json')
    writeFileSync(table, JSON.stringify([['a', 'b', 'c', 'd']]))
    const out = join(dir, 'theme.xlsx')
    await run(['create', '--type', 'xlsx', '--from', table, '--out', out])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'format_range', range: 'A1', format: { fillColor: 'accent1+40%', fontColor: 'lt1' } },
        {
          op: 'format_range',
          range: 'B1',
          format: { fill: { pattern: 'lightGray', fg: { theme: 'accent2' }, bg: '#FFFFFF' } },
        },
        {
          op: 'format_range',
          range: 'C1',
          format: {
            fill: {
              gradient: {
                angle: 90,
                stops: [
                  { position: 0, color: 'accent1' },
                  { position: 1, color: '#FFFFFF' },
                ],
              },
            },
          },
        },
        { op: 'format_range', range: 'D1', format: { fillColor: 'nope' } },
      ]),
    )
    const rejected = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(rejected.code).toBe(1)
    writeFileSync(ops, JSON.stringify(JSON.parse(readFileSync(ops, 'utf8')).slice(0, 3)))
    const applied = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    const styles = await part(out, 'xl/styles.xml')
    expect(styles).toContain('<fgColor theme="4" tint="0.4"/>')
    expect(styles).toContain('<color theme="0"/>')
    expect(styles).toContain(
      '<patternFill patternType="lightGray"><fgColor theme="5"/><bgColor rgb="FFFFFFFF"/></patternFill>',
    )
    expect(styles).toContain(
      '<gradientFill degree="90"><stop position="0"><color theme="4"/></stop>',
    )
    expect(styles).not.toContain('8EA9DB')

    const read = await run(['sheet', 'read', out, '--formats', '--json'])
    expect(read.code).toBe(0)
    const formats = read.json().detail.formats
    expect(formats.A1).toMatchObject({
      fontColor: '#FFFFFF',
      fontColorTheme: { theme: 'lt1' },
      fill: { pattern: 'solid', fg: { theme: 'accent1', tint: 0.4 } },
    })
    expect(formats.A1.fillColor).toBe(formats.A1.fill.fg.rgb)
    expect(formats.A1.fillColor).toMatch(/^#8[EF]A[9A]D[BC]$/)
    expect(formats.B1).toMatchObject({
      fill: {
        pattern: 'lightGray',
        fg: { rgb: '#ED7D31', theme: 'accent2' },
        bg: { rgb: '#FFFFFF' },
      },
    })
    expect(formats.C1.fill.gradient).toMatchObject({
      angle: 90,
      stops: [
        { position: 0, color: { theme: 'accent1' } },
        { position: 1, color: { rgb: '#FFFFFF' } },
      ],
    })
    // B1 and C1 keep the Normal font (theme="1" in every Excel stylesheet): no font color echoed
    expect(formats.B1.fontColorTheme).toBeUndefined()
    expect(formats.C1.fontColor).toBeUndefined()
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'format_range',
          range: 'D1',
          format: {
            fill: {
              gradient: {
                type: 'path',
                left: 0.5,
                right: 0.5,
                top: 0.25,
                bottom: 0.75,
                stops: [
                  { position: 0, color: '#FFFFFF' },
                  { position: 1, color: 'accent6-25%' },
                ],
              },
            },
          },
        },
      ]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', ops, '--json'])).code).toBe(0)
    const again = await run(['sheet', 'read', out, '--formats', '--json'])
    expect(again.json().detail.formats.D1.fill.gradient).toMatchObject({
      type: 'path',
      left: 0.5,
      right: 0.5,
      top: 0.25,
      bottom: 0.75,
      stops: [{ color: { rgb: '#FFFFFF' } }, { color: { theme: 'accent6', tint: -0.25 } }],
    })
  })

  it.skipIf(!sidecar)(
    'flattens theme colors to rgb when the workbook has no theme part',
    async () => {
      const dir = tempDir()
      const table = join(dir, 't.json')
      writeFileSync(table, JSON.stringify([['a']]))
      const out = join(dir, 'notheme.xlsx')
      await run(['create', '--type', 'xlsx', '--from', table, '--out', out])
      const zip = await JSZip.loadAsync(readFileSync(out))
      expect(zip.file('xl/theme/theme1.xml')).not.toBeNull()
      zip.remove('xl/theme/theme1.xml')
      writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer' }))
      const ops = join(dir, 'ops.json')
      writeFileSync(
        ops,
        JSON.stringify([{ op: 'format_range', range: 'A1', format: { fillColor: 'accent1+40%' } }]),
      )
      const applied = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
      expect(applied.code).toBe(0)
      expect(applied.json().warnings).toMatchObject([{ code: 'theme_colors_flattened' }])
      const styles = await part(out, 'xl/styles.xml')
      expect(styles).not.toContain('theme="4"')
      expect(styles).toMatch(/<fgColor rgb="FF8[EF]A[9A]D[BC]"\/>/)
    },
  )

  it('reports usage errors for missing inputs', async () => {
    const dir = tempDir()
    expect((await run(['create', '--type', 'xlsx', '--out', join(dir, 'x.xlsx')])).code).toBe(1)
    const notTable = join(dir, 'x.json')
    writeFileSync(notTable, '{"nope":1}')
    expect(
      (await run(['create', '--type', 'xlsx', '--from', notTable, '--out', join(dir, 'y.xlsx')]))
        .code,
    ).toBe(1)
    expect((await run(['sheet', 'nope', 'a.xlsx'])).code).toBe(1)
  })

  it('splits recalc reads into bands under the sidecar cap', () => {
    const one = recalcBands({ startRow: 3, endRow: 10, startColumn: 0, endColumn: 4 })
    expect(one).toEqual([{ startRow: 3, endRow: 10, startColumn: 0, endColumn: 4 }])
    // 600 rows × 100 columns = 60k cells → 200-row bands
    const bands = recalcBands({ startRow: 0, endRow: 599, startColumn: 0, endColumn: 99 })
    expect(bands).toHaveLength(3)
    expect(bands[0]).toEqual({ startRow: 0, endRow: 199, startColumn: 0, endColumn: 99 })
    expect(bands[2]!.endRow).toBe(599)
    for (const b of bands) {
      expect((b.endRow - b.startRow + 1) * (b.endColumn - b.startColumn + 1)).toBeLessThanOrEqual(
        RECALC_CELL_CAP,
      )
    }
    // wider than the cap: column bands too
    const wide = recalcBands({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 25_000 })
    expect(wide).toHaveLength(2)
    expect(wide[1]).toEqual({ startRow: 0, endRow: 0, startColumn: 20_000, endColumn: 25_000 })
  })

  it.skipIf(!sidecar)(
    'reads a wide window with uncached formulas without hitting the recalc cap',
    async () => {
      const dir = tempDir()
      const rows = Array.from({ length: 220 }, (_, r) => [r, `=A${r + 1}*2`])
      const table = join(dir, 'wide.json')
      writeFileSync(table, JSON.stringify(rows))
      const out = join(dir, 'wide.xlsx')
      expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
      // strip the cached values so read has to evaluate every formula in its window
      const zip = await JSZip.loadAsync(readFileSync(out))
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      zip.file('xl/worksheets/sheet1.xml', sheet.replace(/(<f>[^<]*<\/f>)<v>[^<]*<\/v>/g, '$1'))
      writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer' }))
      const r = await run(['sheet', 'read', out, '--range', 'A1:ZZ220', '--json'])
      expect(r.code).toBe(0)
      expect(r.json().detail.rows[219][1]).toBe(438)
    },
  )
})

describe('genoffice sheet read: --cols, --max-rows, --where, --stats', () => {
  async function book(dir: string): Promise<string> {
    const table = join(dir, 'r.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['name', 'qty', 'price', 'total'],
        ['Apple', 2, 1.5, '=B2*C2'],
        ['Pear', 3, 2, '=B3*C3'],
        ['Fig', null, 4, '=B4/0'],
        ['Sum', '=SUM(B2:B4)', null, '=SUM(D2:D3)'],
      ]),
    )
    const xlsx = join(dir, 'r.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)
    return xlsx
  }

  it.skipIf(!sidecar)('--cols keeps only the named columns and labels them', async () => {
    const xlsx = await book(tempDir())
    const r = await run(['sheet', 'read', xlsx, '--cols', 'A,C:D', '--json'])
    expect(r.code).toBe(0)
    const d = r.json().detail
    expect(d.columns).toEqual(['A', 'C', 'D'])
    expect(d.rows[1]).toEqual(['Apple', 1.5, 3])
    expect(Object.keys(d.formulas)).toEqual(['D2', 'D3', 'D4', 'D5'])
    expect(d.range).toBe('A1:D5')
    const outside = await run(['sheet', 'read', xlsx, '--range', 'A1:B5', '--cols', 'D', '--json'])
    expect(outside.code).toBe(1)
    expect(outside.json().error).toBe('out_of_range')
    const bad = await run(['sheet', 'read', xlsx, '--cols', 'A;B', '--json'])
    expect(bad.code).toBe(1)
    expect(bad.json().error).toBe('invalid_argument')
  })

  it.skipIf(!sidecar)('--max-rows stops early and says how to continue', async () => {
    const xlsx = await book(tempDir())
    const r = await run(['sheet', 'read', xlsx, '--max-rows', '2', '--json'])
    expect(r.code).toBe(0)
    const d = r.json().detail
    expect(d.rows).toHaveLength(2)
    expect(d).toMatchObject({ rowsShown: 2, rowsTotal: 5, truncated: true, range: 'A1:D2' })
    expect(Object.keys(d.formulas)).toEqual(['D2'])
    expect(d.note).toBe('2 of 5 rows shown; continue with --range A3:D5')
    const page = await run(['sheet', 'read', xlsx, '--range', 'A2:D5', '--max-rows', '2', '--json'])
    expect(page.json().detail).toMatchObject({ range: 'A2:D3', rowsShown: 2, rowsTotal: 4 })
    expect(page.json().detail.note).toBe('2 of 4 rows shown; continue with --range A4:D5')
    const text = await run(['sheet', 'read', xlsx, '--max-rows', '2'])
    expect(text.stdout.trim().split('\n').at(-1)).toContain('2 of 5 rows shown')
    const whole = await run(['sheet', 'read', xlsx, '--json'])
    expect(whole.json().detail).toMatchObject({ rowsShown: 5, rowsTotal: 5, truncated: false })
    expect(whole.json().detail.note).toBeUndefined()
    expect((await run(['sheet', 'read', xlsx, '--max-rows', 'x', '--json'])).code).toBe(1)
  })

  it.skipIf(!sidecar)('--where lists one kind of cell without the row grid', async () => {
    const xlsx = await book(tempDir())
    const formulas = (await run(['sheet', 'read', xlsx, '--where', 'formula', '--json'])).json()
    expect(formulas.detail.rows).toBeUndefined()
    expect(formulas.detail.matches).toBe(5)
    expect(formulas.detail.cells[0]).toEqual({ ref: 'D2', value: 3, formula: '=B2*C2' })
    const errors = (await run(['sheet', 'read', xlsx, '--where', 'error', '--json'])).json()
    expect(errors.detail.cells.map((c: { ref: string }) => c.ref)).toEqual(['D4'])
    expect(errors.detail.cells[0].value).toBe('#DIV/0!')
    expect(errors.summary).toContain('1 error cells')
    const empty = (await run(['sheet', 'read', xlsx, '--where', 'empty', '--json'])).json()
    expect(empty.detail.cells.map((c: { ref: string }) => c.ref)).toEqual(['B4', 'C5'])
    const text = (
      await run(['sheet', 'read', xlsx, '--where', 'text', '--cols', 'A', '--json'])
    ).json()
    expect(text.detail.matches).toBe(5)
    const numbers = (
      await run(['sheet', 'read', xlsx, '--where', 'number', '--range', 'B2:C3', '--json'])
    ).json()
    expect(numbers.detail.matches).toBe(4)
    const capped = (
      await run(['sheet', 'read', xlsx, '--where', 'empty', '--max-rows', '4', '--json'])
    ).json().detail
    expect(capped.note).toBe('4 of 5 rows shown; continue with --range A5:D5')
    const typo = await run(['sheet', 'read', xlsx, '--where', 'formlua', '--json'])
    expect(typo.code).toBe(1)
    expect(typo.json()).toMatchObject({
      error: 'invalid_argument',
      suggestion: 'did you mean --where formula',
    })
  })

  it.skipIf(!sidecar)('--stats counts the sheet instead of listing it', async () => {
    const xlsx = await book(tempDir())
    const r = await run(['sheet', 'read', xlsx, '--stats', '--json'])
    expect(r.code).toBe(0)
    const d = r.json().detail
    expect(d.rows).toBeUndefined()
    expect(d.formulas).toBeUndefined()
    expect(d.stats).toMatchObject({
      sheets: [{ name: 'r', rows: 5, columns: 4 }],
      usedRange: 'A1:D5',
      scanned: 'A1:D5',
      rows: 5,
      columns: 4,
      nonEmpty: 18,
      formulas: 5,
      errors: 1,
      numbers: 9,
      text: 8,
      booleans: 0,
      merges: 0,
    })
    expect(d.features).toBeDefined()
    expect(r.json().summary).toContain('18 non-empty cells, 5 formulas, 1 errors')
  })
})

describe('genoffice convert xlsx → csv', () => {
  it.skipIf(!sidecar)('writes the active sheet as displayed text with a BOM', async () => {
    const dir = tempDir()
    const table = join(dir, 'table.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['item', 'qty', 'note'],
        ['Apple, red', 2, 'say "hi"'],
        ['Pear', 3, true],
        ['Total', '=SUM(B2:B3)', null],
      ]),
    )
    const xlsx = join(dir, 'table.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)
    const r = await run(['convert', xlsx, '--to', 'csv', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ sheet: 'table', sheets: 1, formulas: 1 })
    const text = readFileSync(join(dir, 'table.csv'), 'utf-8')
    expect(text.charCodeAt(0)).toBe(0xfeff)
    expect(text.slice(1)).toBe(
      'item,qty,note\r\n"Apple, red",2,"say ""hi"""\r\nPear,3,TRUE\r\nTotal,5,\r\n',
    )
    const missing = await run([
      'convert',
      xlsx,
      '--to',
      'csv',
      '--sheet',
      'Nope',
      '--force',
      '--json',
    ])
    expect(missing.code).toBe(1)
    expect(missing.json().detail.sheets).toEqual(['table'])
  })
})

describe('displayText', () => {
  it('formats like the grid and shifts only calendar dates in 1904 workbooks', () => {
    expect(displayText('x', undefined, false)).toBe('x')
    expect(displayText(true, undefined, false)).toBe('TRUE')
    expect(displayText(null, undefined, false)).toBe('')
    expect(displayText(1234.5, undefined, false)).toBe('1234.5')
    expect(displayText(0.256, '0.0%', false)).toBe('25.6%')
    expect(displayText(45000, 'm/d/yyyy', false)).toBe('3/15/2023')
    expect(displayText(45000, 'm/d/yyyy', true)).toBe('3/16/2027')
    expect(displayText(1.5, '[h]:mm', true)).toBe('36:00')
    expect(displayText(0.5, 'h:mm AM/PM', true)).toBe('12:00 PM')
  })
})

describe('sheet guard rails', () => {
  async function book(dir: string, rows: unknown[][]): Promise<string> {
    const table = join(dir, 'b.json')
    writeFileSync(table, JSON.stringify(rows))
    const xlsx = join(dir, 'b.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)
    return xlsx
  }

  it('refuses cell addresses beyond the sheet limits and malformed ranges as usage errors', async () => {
    const dir = tempDir()
    const xlsx = await book(dir, [['a'], [1]])
    const cells = join(dir, 'cells.json')
    writeFileSync(cells, JSON.stringify([{ cell: 'ZZZZ99999', value: 1 }]))
    const bad = await run(['sheet', 'apply', xlsx, '--cells', cells, '--json'])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('ZZZZ99999')
    if (sidecar) {
      expect((await run(['sheet', 'read', xlsx, '--range', 'ZZZ', '--json'])).code).toBe(1)
      const beyond = await run(['sheet', 'read', xlsx, '--range', 'D50:E51', '--json'])
      expect(beyond.code).toBe(0)
      expect(beyond.json().detail.rows).toEqual([
        [null, null],
        [null, null],
      ])
    }
  })

  it.skipIf(!sidecar)(
    'applies ops in order, so a sort followed by a replace sees the sorted sheet',
    async () => {
      const dir = tempDir()
      const xlsx = await book(dir, [['name'], ['c'], ['a'], ['b']])
      const ops = join(dir, 'ops.json')
      writeFileSync(
        ops,
        JSON.stringify([
          { op: 'sort_range', range: 'A1:A4', byColumn: 'A', order: 'asc', hasHeader: true },
          { op: 'find_replace', range: 'A1:A4', find: 'c', replace: 'gamma' },
        ]),
      )
      expect((await run(['sheet', 'apply', xlsx, '--ops', ops, '--json'])).code).toBe(0)
      const r = await run(['sheet', 'read', xlsx, '--range', 'A1:A4', '--json'])
      expect(r.json().detail.rows.map((row: unknown[]) => row[0])).toEqual([
        'name',
        'a',
        'b',
        'gamma',
      ])
    },
  )

  it.skipIf(!sidecar)(
    'needs --force to write --out over another existing file, but edits in place freely',
    async () => {
      const dir = tempDir()
      const xlsx = await book(dir, [['a'], [1]])
      const other = join(dir, 'other.xlsx')
      writeFileSync(other, 'not really a workbook')
      const cells = join(dir, 'cells.json')
      writeFileSync(cells, JSON.stringify([{ cell: 'A2', value: 2 }]))
      const refused = await run([
        'sheet',
        'apply',
        xlsx,
        '--cells',
        cells,
        '--out',
        other,
        '--json',
      ])
      expect(refused.code).toBe(2)
      expect(refused.json().message).toContain('use --force')
      expect(readFileSync(other, 'utf-8')).toBe('not really a workbook')
      expect(
        (await run(['sheet', 'apply', xlsx, '--cells', cells, '--out', other, '--force'])).code,
      ).toBe(0)
      expect((await run(['sheet', 'apply', xlsx, '--cells', cells])).code).toBe(0)
    },
  )
})

describe('formula cache policy', () => {
  it.skipIf(!sidecar)(
    'marks spill formulas as dynamic arrays, types errors, and never caches #NAME?',
    async () => {
      const dir = tempDir()
      const table = join(dir, 't.json')
      writeFileSync(
        table,
        JSON.stringify([
          [1, 2, 3],
          ['=FILTER(A1:C1,A1:C1>1)', '=1/0', '=SUM(A1:C1)', '="#N/A"'],
        ]),
      )
      const out = join(dir, 't.xlsx')
      const r = await run(['create', '--type', 'xlsx', '--from', table, '--out', out, '--json'])
      expect(r.code).toBe(0)
      expect(r.json().warnings[0].message).toContain('A2')
      const sheet = await part(out, 'xl/worksheets/sheet1.xml')
      expect(sheet).toMatch(
        /<c r="A2"[^>]*\bcm="1"[^>]*><f t="array" ref="A2">_xlfn\._xlws\.FILTER\(/,
      )
      expect(sheet).not.toContain('#NAME?')
      expect(sheet).toMatch(/<c r="B2"[^>]*\bt="e"[^>]*>.*?<v>#DIV\/0!<\/v>/)
      // text that spells an error is still text
      expect(sheet).toMatch(/<c r="D2"[^>]*\bt="str"[^>]*>.*?<v>#N\/A<\/v>/)
      expect(sheet).toContain('<v>6</v>')
      expect(await part(out, 'xl/metadata.xml')).toContain('name="XLDAPR"')
      expect(await part(out, '[Content_Types].xml')).toContain('/xl/metadata.xml')
      expect(await part(out, 'xl/_rels/workbook.xml.rels')).toContain('Target="metadata.xml"')
      const read = await run(['sheet', 'read', out, '--json'])
      expect(read.code).toBe(0)
    },
  )
})

describe('package validity', () => {
  /// Duplicate Relationship Ids or Override PartNames are invalid under the Open
  /// Packaging Conventions, and Excel opens such a workbook only after repairing it.
  async function expectSingleStylesPart(out: string): Promise<void> {
    const rels = await part(out, 'xl/_rels/workbook.xml.rels')
    const ids = [...rels.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1])
    expect(ids).toEqual([...new Set(ids)])
    expect([...rels.matchAll(/\bType="[^"]*\/relationships\/styles"/g)].length).toBe(1)

    const types = await part(out, '[Content_Types].xml')
    const names = [...types.matchAll(/\bPartName="([^"]+)"/g)].map((m) => m[1])
    expect(names).toEqual([...new Set(names)])
    expect(names.filter((name) => name === '/xl/styles.xml')).toEqual(['/xl/styles.xml'])

    expect(await part(out, 'xl/styles.xml')).toContain('<cellXfs')
  }

  it('writes the styles part once in a new workbook', async () => {
    const dir = tempDir()
    const table = join(dir, 'table.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['item', 'qty'],
        ['Apple', 2],
      ]),
    )
    const out = join(dir, 'table.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
    await expectSingleStylesPart(out)
  })

  it('writes the styles part once in a multi-sheet workbook', async () => {
    const dir = tempDir()
    const multi = join(dir, 'multi.json')
    writeFileSync(
      multi,
      JSON.stringify({
        sheets: [
          { name: 'Data', rows: [['a', 1]] },
          { name: 'Notes', rows: [['hello']] },
        ],
      }),
    )
    const out = join(dir, 'multi.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', multi, '--out', out])).code).toBe(0)
    await expectSingleStylesPart(out)
  })
})

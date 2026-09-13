import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { RECALC_CELL_CAP, displayText, recalcBands } from '../src/formats/xlsx'
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

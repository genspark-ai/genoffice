import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { patchToStyleEdit } from '../src/formats/xlsx-dsl'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file(name)!.async('string')
}

async function book(dir: string, rows: unknown[][]): Promise<string> {
  const table = join(dir, 'table.json')
  writeFileSync(table, JSON.stringify(rows))
  const out = join(dir, 'book.xlsx')
  expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
  return out
}

describe('genoffice sheet apply --ops (workbook DSL)', () => {
  it('runs content, format and layout ops through the in-memory workbook and saves them', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['item', 'qty', 'price'],
      ['Apple', 2, 1.5],
      ['Pear', 3, 2],
    ])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'set_range',
          sheet: 'table',
          range: 'D1:D3',
          values: [['total'], ['=B2*C2'], ['=B3*C3']],
        },
        {
          op: 'format_range',
          sheet: 'table',
          range: 'A1:D1',
          format: { bold: true, fillColor: '#FFFF00', horizontalAlign: 'center' },
        },
        { op: 'find_replace', sheet: 'table', range: 'A1:A3', find: 'Pear', replace: 'Plum' },
        { op: 'merge_cells', sheet: 'table', range: 'A5:D5' },
        { op: 'set_col_width', sheet: 'table', column: 'A', widthPx: 140 },
        { op: 'set_row_height', sheet: 'table', row: 1, heightPoints: 24 },
      ]),
    )
    const dry = await run(['sheet', 'apply', out, '--ops', ops, '--dry-run', '--json'])
    expect(dry.code).toBe(0)
    expect(dry.json().detail.cells).toBeGreaterThan(0)
    const before = readFileSync(out)

    const r = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(readFileSync(out).equals(before)).toBe(false)
    const sheet = await part(out, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<f>B2*C2</f>')
    expect(sheet).toContain('Plum')
    expect(sheet).not.toContain('Pear')
    expect(sheet).toContain('<mergeCell ref="A5:D5"/>')
    expect(sheet).toMatch(/<col [^>]*min="1" max="1" [^>]*width="20"/)
    expect(sheet).toMatch(/<row r="1"[^>]* ht="24"/)
    const styles = await part(out, 'xl/styles.xml')
    expect(styles).toContain('FFFF00')
    expect(styles).toContain('<b/>')
    if (sidecar) {
      const read = await run(['sheet', 'read', out, '--range', 'D1:D3', '--json'])
      expect(read.json().detail.rows.map((r: unknown[]) => r[0])).toEqual(['total', 3, 6])
    }
  })

  it('handles structural and sheet ops without diffing shifted cells', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['a', 'b'],
      [1, 2],
    ])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
        { op: 'add_sheet', name: 'Notes' },
        { op: 'rename_sheet', sheet: 'table', name: 'Data' },
      ]),
    )
    const r = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ cells: 0, structural: 1, sheets_changed: 2 })
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="Data"')
    expect(wb).toContain('name="Notes"')
    const sheet = await part(out, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<row r="2"')
    expect(sheet).toMatch(/<c r="A2"[^>]*>/)
  })

  it('refuses unsupported ops and mixed batches with a usage error', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a']])
    const bad = join(dir, 'bad.json')
    writeFileSync(
      bad,
      JSON.stringify([{ op: 'add_sparkline', sheet: 'table', dataRange: 'A1:B2', type: 'line' }]),
    )
    const r = await run(['sheet', 'apply', out, '--ops', bad, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('not available headless')
    expect(r.json().detail.supported).toContain('set_range')
    expect(r.json().detail.supported).toContain('add_chart')

    const mixed = join(dir, 'mixed.json')
    writeFileSync(
      mixed,
      JSON.stringify([
        { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
        { op: 'set_cell', sheet: 'table', address: 'A1', value: 'x' },
      ]),
    )
    const m = await run(['sheet', 'apply', out, '--ops', mixed, '--json'])
    expect(m.code).toBe(1)
    expect(m.json().message).toContain('rejected')

    const nosheet = join(dir, 'nosheet.json')
    writeFileSync(
      nosheet,
      JSON.stringify([{ op: 'set_cell', sheet: 'Nope', address: 'A1', value: 1 }]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', nosheet, '--json'])).code).toBe(1)
  })

  it('maps DSL format patches onto the gateway style delta', () => {
    expect(
      patchToStyleEdit({
        bold: true,
        italic: null,
        fontColor: '#FF0000',
        fillColor: null,
        numberFormat: null,
        horizontalAlign: 'center',
        verticalAlign: 'top',
        textRotation: -45,
        indent: null,
        border: { type: 'all', color: '#000000' },
      }),
    ).toEqual({
      bold: true,
      italic: false,
      fontColor: '#FF0000',
      fillColor: null,
      numberFormat: 'General',
      horizontalAlignment: 'center',
      verticalAlignment: 'top',
      textRotation: 135,
      indent: 0,
      borderTop: { style: 'thin', color: '#000000' },
      borderBottom: { style: 'thin', color: '#000000' },
      borderLeft: { style: 'thin', color: '#000000' },
      borderRight: { style: 'thin', color: '#000000' },
    })
    expect(patchToStyleEdit({ textRotation: 'vertical', border: { type: 'none' } })).toEqual({
      textRotation: 255,
      borderTop: null,
      borderBottom: null,
      borderLeft: null,
      borderRight: null,
    })
    expect(patchToStyleEdit({ border: { type: 'left' } })).toEqual({
      borderLeft: { style: 'thin' },
    })
  })

  it('merges overlapping format patches, validates existing merges, and defaults to the active sheet', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a', 'b', 'c']])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'format_range', range: 'A1:C1', format: { bold: true } },
        { op: 'format_range', range: 'A1:A1', format: { fillColor: '#00FF00' } },
        { op: 'merge_cells', range: 'A2:B2' },
      ]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', ops])).code).toBe(0)
    const styles = await part(out, 'xl/styles.xml')
    // A1 keeps bold from the first patch and gains the fill from the second: one xf with both
    expect(styles).toMatch(/<xf [^>]*fontId="(\d+)"[^>]*fillId="(\d+)"/)
    expect(styles).toContain('00FF00')
    expect(styles).toContain('<b/>')

    const overlap = join(dir, 'overlap.json')
    writeFileSync(overlap, JSON.stringify([{ op: 'merge_cells', range: 'B2:C2' }]))
    const r = await run(['sheet', 'apply', out, '--ops', overlap, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toMatch(/merge/i)

    const unknown = await run(['sheet', 'apply', out, '--ops', ops, '--sheet', 'Nope', '--json'])
    expect(unknown.code).toBe(1)
    expect(unknown.json().detail.sheets).toEqual(['table'])
  })

  it.skipIf(!sidecar)('caches formula results when the same batch renames the sheet', async () => {
    const dir = tempDir()
    const out = await book(dir, [[2, 3]])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'rename_sheet', sheet: 'table', name: 'Data' },
        { op: 'set_cell', sheet: 'table', address: 'C1', value: '=A1+B1' },
      ]),
    )
    const r = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.cached_values).toBe(true)
    expect(await part(out, 'xl/worksheets/sheet1.xml')).toContain('<v>5</v>')
    expect(await part(out, 'xl/workbook.xml')).toContain('name="Data"')
  })

  it('resolves sheets with escaped names through the gateway parsers', async () => {
    const dir = tempDir()
    const multi = join(dir, 'multi.json')
    writeFileSync(
      multi,
      JSON.stringify({
        sheets: [
          { name: 'P&L > 2026', rows: [['x']] },
          { name: 'Other', rows: [['y']] },
        ],
      }),
    )
    const out = join(dir, 'multi.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', multi, '--out', out])).code).toBe(0)
    const ops = join(dir, 'ops.json')
    // no "sheet": the active tab (the first one, with the escaped name) receives the write
    writeFileSync(ops, JSON.stringify([{ op: 'set_cell', address: 'B1', value: 'hit' }]))
    const r = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(await part(out, 'xl/worksheets/sheet1.xml')).toContain('hit')
    expect(await part(out, 'xl/worksheets/sheet2.xml')).not.toContain('hit')
  })
})

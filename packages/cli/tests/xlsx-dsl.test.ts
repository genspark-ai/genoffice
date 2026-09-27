import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { patchToStyleEdit } from '../src/formats/xlsx-dsl'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'
import { book, cells } from './xlsx-helpers'

const sidecar = Boolean(xlsxSidecarPath())

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file(name)!.async('string')
}

describe('genoffice sheet apply --ops (workbook DSL)', () => {
  it('set_cell writes a string starting with "=" as a formula unless type is text', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a', 1]])
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'set_cell', sheet: 'table', address: 'C1', value: '=B1*2' },
        { op: 'set_cell', sheet: 'table', address: 'D1', value: '=B1*2', type: 'text' },
      ]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', ops])).code).toBe(0)
    const sheet = await cells(out, 'xl/worksheets/sheet1.xml')
    expect(sheet.get('C1')?.formula).toBe('B1*2')
    expect(sheet.get('D1')).toEqual({ value: '=B1*2' })
  })

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

  it('takes the source sheet of a copy by name', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['item', 'qty'],
      ['Apple', 2],
    ])
    const add = join(dir, 'add.json')
    writeFileSync(add, JSON.stringify([{ op: 'add_sheet', name: 'Notes' }]))
    expect((await run(['sheet', 'apply', out, '--ops', add, '--json'])).code).toBe(0)
    const copy = join(dir, 'copy.json')
    writeFileSync(
      copy,
      JSON.stringify([
        { op: 'copy_range', sheet: 'Notes', sourceSheet: 'table', source: 'A1:B2', target: 'A1' },
      ]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', copy, '--json'])).code).toBe(0)
    const notes = await part(out, 'xl/worksheets/sheet2.xml')
    expect(notes).toMatch(/<c r="A2"[^>]*>/)
    expect(notes).toMatch(/<c r="B2"[^>]*><v>2<\/v>/)
    const guide = await run(['guide', 'sheets', 'copy_range'])
    expect(guide.stdout).toContain('sourceSheet?')
    expect(guide.stdout).not.toContain('sourceSheetId')
  })

  it('refuses unsupported ops, runs structural and content ops in one batch', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a']])
    const bad = join(dir, 'bad.json')
    writeFileSync(bad, JSON.stringify([{ op: 'refresh_pivot', sheet: 'table' }]))
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
    expect(m.code).toBe(0)
    expect(m.json().detail).toMatchObject({ cells: 1, structural: 1 })
    const shifted = await cells(out, 'xl/worksheets/sheet1.xml')
    expect(shifted.get('A1')).toEqual({ value: 'x' })
    expect(shifted.get('A2')).toEqual({ value: 'a' })

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
    expect(
      patchToStyleEdit({
        fontColor: 'accent1+40%',
        fillColor: { theme: 'dk2', tint: -0.25 },
        border: { type: 'top', color: 'tx1' },
      }),
    ).toEqual({
      fontColor: { theme: 4, tint: 0.4 },
      fillColor: { theme: 3, tint: -0.25 },
      borderTop: { style: 'thin', color: { theme: 1 } },
    })
    expect(
      patchToStyleEdit({
        fillColor: '#ffffff',
        fill: {
          gradient: {
            angle: 90,
            stops: [
              { position: 0, color: 'accent2' },
              { position: 1, color: '#FFFFFF' },
            ],
          },
        },
      }),
    ).toEqual({
      fill: {
        gradient: {
          angle: 90,
          stops: [
            { position: 0, color: { theme: 5 } },
            { position: 1, color: '#FFFFFF' },
          ],
        },
      },
    })
    expect(patchToStyleEdit({ fill: { pattern: 'lightGray', fg: '#ff0000' } })).toEqual({
      fill: { pattern: 'lightGray', fg: '#FF0000' },
    })
    expect(patchToStyleEdit({ fill: null })).toEqual({ fill: null })
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

  it('lets a later fillColor replace an earlier pattern or gradient fill on the same cell', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a', 'b', 'c']])
    const ops = join(dir, 'ops.json')
    const gradient = {
      gradient: {
        angle: 90,
        stops: [
          { position: 0, color: '#112233' },
          { position: 1, color: '#FFFFFF' },
        ],
      },
    }
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'format_range',
          range: 'A1:C1',
          format: { fill: { pattern: 'lightGray', fg: '#445566' } },
        },
        { op: 'format_range', range: 'A1', format: { fillColor: '#00FF00' } },
        { op: 'format_range', range: 'B1', format: { fill: gradient } },
        { op: 'format_range', range: 'B1', format: { fillColor: null } },
        { op: 'format_range', range: 'C1', format: { fillColor: '#FF0000' } },
        { op: 'format_range', range: 'C1', format: { fill: gradient } },
      ]),
    )
    expect((await run(['sheet', 'apply', out, '--ops', ops])).code).toBe(0)
    const styles = await part(out, 'xl/styles.xml')
    expect(styles).toContain('<patternFill patternType="solid"><fgColor rgb="FF00FF00"/>')
    expect(styles).toContain('<gradientFill degree="90">')
    expect(styles).not.toContain('lightGray')
    expect(styles).not.toContain('FF0000')
    const sheet = await part(out, 'xl/worksheets/sheet1.xml')
    const b1 = /<c r="B1"[^>]*>/.exec(sheet)![0]
    expect(b1).not.toMatch(/ s="[1-9]/)
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

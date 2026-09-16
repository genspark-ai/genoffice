import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'
import { book, cells, part } from './xlsx-helpers'

const SHEET1 = 'xl/worksheets/sheet1.xml'

async function apply(out: string, dir: string, ops: unknown[]) {
  const file = join(dir, `ops-${Math.random().toString(36).slice(2)}.json`)
  writeFileSync(file, JSON.stringify(ops))
  return run(['sheet', 'apply', out, '--ops', file, '--json'])
}

async function comments(out: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(readFileSync(out))
  const names = Object.keys(zip.files).filter((f) => /^xl\/comments\d*\.xml$/.test(f))
  return Promise.all(names.map((f) => zip.file(f)!.async('string')))
}

describe('sheet apply: structural and content ops in one batch', () => {
  it('fills a row inserted earlier in the batch; existing formulas shift', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['a', 'b', 'sum'],
      [1, 2, '=A2+B2'],
    ])
    const r = await apply(out, dir, [
      { op: 'insert_rows', sheet: 'table', row: 2, count: 1 },
      { op: 'set_range', sheet: 'table', start: 'A2', values: [[10, 20, '=A2+B2']] },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ cells: 3, structural: 1 })
    const grid = await cells(out, SHEET1)
    expect(grid.get('A2')).toEqual({ value: 10 })
    expect(grid.get('B2')).toEqual({ value: 20 })
    expect(grid.get('C2')?.formula).toBe('A2+B2')
    expect(grid.get('A3')).toEqual({ value: 1 })
    expect(grid.get('C3')?.formula).toBe('A3+B3')
  })

  it('addresses the grid right of a deleted column by its new letters', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a', 'b', 'c', 'd']])
    const r = await apply(out, dir, [
      { op: 'delete_cols', sheet: 'table', column: 'B', count: 1 },
      { op: 'set_cell', sheet: 'table', address: 'B1', value: 'z' },
    ])
    expect(r.code).toBe(0)
    const grid = await cells(out, SHEET1)
    expect(grid.get('A1')).toEqual({ value: 'a' })
    expect(grid.get('B1')).toEqual({ value: 'z' })
    expect(grid.get('C1')).toEqual({ value: 'd' })
    expect(grid.has('D1')).toBe(false)
  })

  it('writes into a sheet added earlier in the batch and follows a rename', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a']])
    const r = await apply(out, dir, [
      { op: 'add_sheet', name: 'Notes' },
      { op: 'set_cell', sheet: 'Notes', address: 'A1', value: 'hello' },
      { op: 'rename_sheet', sheet: 'table', name: 'Data' },
      { op: 'set_cell', sheet: 'Data', address: 'B1', value: 'renamed' },
      { op: 'set_cell', sheet: 'table', address: 'C1', value: 'old name still works' },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ cells: 3, sheets_changed: 2 })
    expect(await cells(out, 'xl/worksheets/sheet2.xml')).toEqual(
      new Map([['A1', { value: 'hello' }]]),
    )
    const first = await cells(out, SHEET1)
    expect(first.get('B1')).toEqual({ value: 'renamed' })
    expect(first.get('C1')).toEqual({ value: 'old name still works' })
  })

  it('renames a sheet added in the same batch and collapses a chain of renames', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a']])
    const r = await apply(out, dir, [
      { op: 'add_sheet', name: 'Notes' },
      { op: 'rename_sheet', sheet: 'Notes', name: 'Memo' },
      { op: 'set_cell', sheet: 'Memo', address: 'A1', value: 'hello' },
      { op: 'rename_sheet', sheet: 'table', name: 'Step' },
      { op: 'rename_sheet', sheet: 'Step', name: 'Data' },
    ])
    expect(r.code).toBe(0)
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="Memo"')
    expect(wb).toContain('name="Data"')
    expect(wb).not.toMatch(/name="(Notes|Step|table)"/)
    expect(await cells(out, 'xl/worksheets/sheet2.xml')).toEqual(
      new Map([['A1', { value: 'hello' }]]),
    )
  })

  it('refuses writes to a renamed sheet and a new sheet that took its file name', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a']])
    const r = await apply(out, dir, [
      { op: 'rename_sheet', sheet: 'table', name: 'Archive' },
      { op: 'add_sheet', name: 'table' },
      { op: 'set_cell', sheet: 'Archive', address: 'B1', value: 'old' },
    ])
    expect(r.code).toBe(1)
    expect(r.json().error).toBe('op_rejected')
    expect(r.json().message).toContain(
      'add_sheet "table" reuses the file name of the sheet renamed to "Archive" while ops[2] (set_cell)',
    )
    expect(await cells(out, SHEET1)).toEqual(new Map([['A1', { value: 'a' }]]))

    const bare = await apply(out, dir, [
      { op: 'rename_sheet', sheet: 'table', name: 'Archive' },
      { op: 'add_sheet', name: 'table' },
    ])
    expect(bare.code).toBe(0)
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="Archive"')
    expect(wb).toContain('name="table"')
  })

  it('moves content and formats written before a shift along with the grid', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a'], ['b']])
    const r = await apply(out, dir, [
      { op: 'set_cell', sheet: 'table', address: 'A3', value: 'late' },
      { op: 'set_formula', sheet: 'table', address: 'B1', formula: '=A1*2' },
      { op: 'format_range', sheet: 'table', range: 'A1:B1', format: { bold: true } },
      { op: 'insert_rows', sheet: 'table', row: 1, count: 2 },
      { op: 'set_cell', sheet: 'table', address: 'A1', value: 'top' },
    ])
    expect(r.code).toBe(0)
    const grid = await cells(out, SHEET1)
    expect(grid.get('A1')).toEqual({ value: 'top' })
    expect(grid.get('A3')).toEqual({ value: 'a' })
    expect(grid.get('B3')?.formula).toBe('A3*2')
    expect(grid.get('A5')).toEqual({ value: 'late' })
    const xml = await part(out, SHEET1)
    expect(xml).toMatch(/<c r="A3"[^>]* s="[1-9]\d*"/)
    expect(xml).toMatch(/<c r="B3"[^>]* s="[1-9]\d*"/)
    expect(xml).not.toMatch(/<c r="A1"[^>]* s="[1-9]\d*"/)
  })

  it('drops content that a later delete removes and keeps deleted-sheet edits out of the file', async () => {
    const dir = tempDir()
    const out = await book(dir, [['a'], ['b'], ['c']])
    const r = await apply(out, dir, [
      { op: 'set_cell', sheet: 'table', address: 'A2', value: 'gone' },
      { op: 'delete_rows', sheet: 'table', row: 2, count: 1 },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ cells: 0, structural: 1 })
    const grid = await cells(out, SHEET1)
    expect(grid.get('A2')).toEqual({ value: 'c' })
  })

  it('drops formats, payloads and edits aimed at a sheet deleted later in the batch', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5]])
    const r = await apply(out, dir, [
      { op: 'set_cell', sheet: 'table', address: 'B1', value: 'gone' },
      { op: 'format_range', sheet: 'table', range: 'A1:B1', format: { bold: true } },
      { op: 'set_row_height', sheet: 'table', row: 1, heightPoints: 30 },
      {
        op: 'add_conditional_format',
        sheet: 'table',
        range: 'A2:A2',
        rule: { kind: 'number', operator: 'greaterThan', value: 1, format: { bold: true } },
      },
      { op: 'add_sheet', name: 'Keep' },
      { op: 'set_cell', sheet: 'Keep', address: 'A1', value: 'kept' },
      { op: 'delete_sheet', sheet: 'table' },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ cells: 1, structural: 0, sheets_changed: 2 })
    const dropped = r.stderr + JSON.stringify(r.json().warnings ?? [])
    expect(dropped).toContain('ops[1] (format_range) dropped: sheet table is deleted later')
    expect(dropped).toContain('ops[3] (add_conditional_format) dropped')
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="Keep"')
    expect(wb).not.toContain('name="table"')
  })

  it('filters a sheet recreated under a deleted name against the new sheet', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5]])
    expect((await apply(out, dir, [{ op: 'add_sheet', name: 'Other' }])).code).toBe(0)
    const r = await apply(out, dir, [
      { op: 'delete_sheet', sheet: 'table' },
      { op: 'add_sheet', name: 'table' },
      { op: 'set_range', sheet: 'table', start: 'A1', values: [['n'], [1], [2], [3]] },
      { op: 'set_filter', sheet: 'table', range: 'A1:A4' },
      { op: 'set_filter_criteria', sheet: 'table', column: 'A', values: ['2', '3'] },
    ])
    expect(r.code).toBe(0)
    const xml = await part(out, 'xl/worksheets/sheet3.xml')
    expect(xml).toContain('<autoFilter ref="A1:A4"')
    expect(xml).toMatch(/<row r="2"[^>]*hidden="1"/)
    expect(xml).not.toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(xml).not.toMatch(/<row r="4"[^>]*hidden="1"/)
  })

  it('sends sheet-less ops to the default sheet recreated under its name', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5]])
    expect((await apply(out, dir, [{ op: 'add_sheet', name: 'Other' }])).code).toBe(0)
    const r = await apply(out, dir, [
      { op: 'delete_sheet', sheet: 'table' },
      { op: 'add_sheet', name: 'table' },
      { op: 'set_cell', address: 'A1', value: 'fresh' },
    ])
    expect(r.code).toBe(0)
    expect(await cells(out, 'xl/worksheets/sheet3.xml')).toEqual(
      new Map([['A1', { value: 'fresh' }]]),
    )
  })

  it('keeps a copy made from a sheet the same batch deletes', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5]])
    expect((await apply(out, dir, [{ op: 'add_sheet', name: 'Other' }])).code).toBe(0)
    const r = await apply(out, dir, [
      { op: 'duplicate_sheet', sheet: 'table' },
      { op: 'delete_sheet', sheet: 'table' },
    ])
    expect(r.code).toBe(0)
    const wb = await part(out, 'xl/workbook.xml')
    expect(wb).toContain('name="table (2)"')
    expect(wb).not.toContain('name="table"')
    expect(await cells(out, 'xl/worksheets/sheet3.xml')).toEqual(
      new Map([
        ['A1', { value: 'n' }],
        ['A2', { value: 5 }],
      ]),
    )
  })

  it('extends the file filter and notes at their shifted positions after a row insert', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [1], [2], [3]])
    const setup = await apply(out, dir, [
      { op: 'set_filter', sheet: 'table', range: 'A1:A4' },
      { op: 'set_note', sheet: 'table', address: 'A1', text: 'header' },
    ])
    expect(setup.code).toBe(0)
    const r = await apply(out, dir, [
      { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
      { op: 'set_filter_criteria', sheet: 'table', column: 'A', values: ['2', '3'] },
      { op: 'set_note', sheet: 'table', address: 'A6', text: 'tail' },
    ])
    expect(r.code).toBe(0)
    const xml = await part(out, SHEET1)
    expect(xml).toContain('<autoFilter ref="A2:A5"')
    expect(xml).not.toMatch(/<row r="[12]"[^>]*hidden="1"/)
    expect(xml).toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(xml).not.toMatch(/<row r="[45]"[^>]*hidden="1"/)
    const notes = (await comments(out)).join('')
    expect(notes).toMatch(/<comment ref="A2"[^>]*>[\s\S]*?header/)
    expect(notes).toMatch(/<comment ref="A6"[^>]*>[\s\S]*?tail/)
    expect(notes).not.toContain('ref="A1"')
  })

  it("gives a sheet recreated under a deleted name none of the old sheet's file state", async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5]])
    expect((await apply(out, dir, [{ op: 'add_sheet', name: 'Other' }])).code).toBe(0)
    const setup = await apply(out, dir, [
      { op: 'set_filter', sheet: 'table', range: 'A1:A2' },
      { op: 'set_note', sheet: 'table', address: 'A1', text: 'old' },
    ])
    expect(setup.code).toBe(0)
    const criteria = await apply(out, dir, [
      { op: 'delete_sheet', sheet: 'table' },
      { op: 'add_sheet', name: 'table' },
      { op: 'set_filter_criteria', sheet: 'table', column: 'A', values: ['5'] },
    ])
    expect(criteria.code).toBe(1)
    expect(criteria.json().message).toContain('table has no filter; run set_filter first')

    const r = await apply(out, dir, [
      { op: 'delete_sheet', sheet: 'table' },
      { op: 'add_sheet', name: 'table' },
      { op: 'set_range', sheet: 'table', start: 'A1', values: [['n'], [1]] },
      { op: 'set_note', sheet: 'table', address: 'A2', text: 'fresh' },
    ])
    expect(r.code).toBe(0)
    expect(await part(out, 'xl/worksheets/sheet3.xml')).not.toContain('<autoFilter')
    const notes = (await comments(out)).join('')
    expect(notes).toContain('fresh')
    expect(notes).not.toContain('old')
  })

  it('refuses address-carrying payload ops placed before a shift of their sheet', async () => {
    const dir = tempDir()
    const out = await book(dir, [['n'], [5], [50]])
    const rule = {
      op: 'add_conditional_format',
      sheet: 'table',
      range: 'A2:A3',
      rule: { kind: 'number', operator: 'greaterThan', value: 10, format: { bold: true } },
    }
    const early = await apply(out, dir, [
      rule,
      { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
    ])
    expect(early.code).toBe(1)
    expect(early.json().error).toBe('op_rejected')
    expect(early.json().message).toContain(
      'ops[0] (add_conditional_format) addresses table before ops[1] (insert_rows)',
    )
    expect(early.json().detail.failures[0]).toMatchObject({ index: 0 })

    const late = await apply(out, dir, [
      { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
      { ...rule, range: 'A3:A4' },
    ])
    expect(late.code).toBe(0)
    const xml = await part(out, SHEET1)
    expect(xml).toContain('<conditionalFormatting sqref="A3:A4"')
  })
})

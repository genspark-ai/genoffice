import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { parseCsv } from '../src/gateway/csv-import'
import {
  csvField,
  csvFromDisplayRows,
  serializeActiveSheetCsv,
  withoutFormulaView,
} from '../src/renderer/csv-export'
import {
  workbookExportCsvRequestSchema,
  workbookExportCsvResultSchema,
  workbookSaveResultSchema,
} from '../src/shared/desktop-api'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

describe('csvField', () => {
  it('passes plain text through unquoted', () => {
    expect(csvField('hello')).toBe('hello')
    expect(csvField('12.5')).toBe('12.5')
    expect(csvField('')).toBe('')
  })

  it('quotes fields carrying commas, quotes, or newlines', () => {
    expect(csvField('a,b')).toBe('"a,b"')
    expect(csvField('say "hi"')).toBe('"say ""hi"""')
    expect(csvField('line1\nline2')).toBe('"line1\nline2"')
  })

  it('normalizes CR/CRLF line breaks before quoting', () => {
    expect(csvField('a\r\nb')).toBe('"a\nb"')
    expect(csvField('a\rb')).toBe('"a\nb"')
  })

  it('leaves tabs unquoted (CSV only cares about its own delimiter)', () => {
    expect(csvField('a\tb')).toBe('a\tb')
  })
})

describe('csvFromDisplayRows', () => {
  it('joins rows with CRLF and ends the file with one', () => {
    expect(
      csvFromDisplayRows([
        ['a', 'b'],
        ['c', ''],
      ]),
    ).toBe('a,b\r\nc,\r\n')
  })

  it('serializes an empty grid to an empty string', () => {
    expect(csvFromDisplayRows([])).toBe('')
  })

  it('round-trips through the CSV importer', () => {
    const rows = [
      ['name', 'note', 'amount'],
      ['Ada, Bob', 'says "ok"', '1,234.50'],
      ['multi\nline', '', '-3'],
    ]
    expect(parseCsv(csvFromDisplayRows(rows), ',')).toEqual(rows)
  })
})

function fakeCsvSheet(rows: string[][]) {
  const columns = Math.max(1, ...rows.map((row) => row.length))
  return {
    getLastRow: () => rows.length - 1,
    getLastColumn: () => columns - 1,
    getSheetId: () => 's1',
    getSheetName: () => 'Sheet1',
    getRange: (row: number, column: number, numRows: number, numColumns: number) => ({
      getDisplayValues: () =>
        rows
          .slice(row, row + numRows)
          .map((cells) =>
            Array.from({ length: numColumns }, (_, offset) => cells[column + offset] ?? ''),
          ),
    }),
    getSheet: () => ({ getCellMatrix: () => ({ forValue: () => undefined }) }),
  }
}

describe('serializeActiveSheetCsv', () => {
  it('serializes the display grid with quoting and rectangular padding', () => {
    const sheet = fakeCsvSheet([['name', 'note'], ['a,b', 'say "hi"', 'extra'], ['solo']])
    expect(serializeActiveSheetCsv(sheet as never, null)).toBe(
      'name,note,\r\n"a,b","say ""hi""",extra\r\nsolo,,\r\n',
    )
  })
})

describe('withoutFormulaView', () => {
  it('reads with the flag suspended and restores it afterwards', () => {
    const sheets = new Set(['s1', 's2'])
    const during = withoutFormulaView(sheets, 's1', () => sheets.has('s1'))
    expect(during).toBe(false)
    expect(sheets.has('s1')).toBe(true)
  })

  it('restores the flag when the read throws', () => {
    const sheets = new Set(['s1'])
    expect(() =>
      withoutFormulaView(sheets, 's1', () => {
        throw new Error('boom')
      }),
    ).toThrow('boom')
    expect(sheets.has('s1')).toBe(true)
  })

  it('never adds the flag to a sheet that did not have it', () => {
    const sheets = new Set<string>()
    withoutFormulaView(sheets, 's1', () => undefined)
    expect(sheets.size).toBe(0)
  })
})

describe('workbookExportCsv schemas', () => {
  it('accepts a minimal request and one with the active-sheet name', () => {
    expect(
      workbookExportCsvRequestSchema.parse({
        fileName: 'data.csv',
        content: 'a,b\r\n',
        hasFormulas: false,
      }),
    ).toBeTruthy()
    expect(
      workbookExportCsvRequestSchema.parse({
        fileName: 'data.csv',
        content: '',
        hasFormulas: true,
        activeSheetName: 'Sheet2',
      }),
    ).toBeTruthy()
  })

  it('rejects unknown keys and oversized names', () => {
    expect(() =>
      workbookExportCsvRequestSchema.parse({
        fileName: 'data.csv',
        content: '',
        hasFormulas: false,
        html: 'nope',
      }),
    ).toThrow()
    expect(() =>
      workbookExportCsvRequestSchema.parse({
        fileName: 'x'.repeat(256),
        content: '',
        hasFormulas: false,
      }),
    ).toThrow()
  })

  it('accepts a direct targetPath for the Save As CSV reroute', () => {
    expect(
      workbookExportCsvRequestSchema.parse({
        fileName: 'data.csv',
        content: 'a\r\n',
        hasFormulas: false,
        targetPath: '/tmp/data.csv',
      }),
    ).toBeTruthy()
  })

  it('lets a canceled save carry the picked CSV path', () => {
    expect(
      workbookSaveResultSchema.parse({ canceled: true, csvSaveAsPath: '/tmp/data.csv' }),
    ).toEqual({ canceled: true, csvSaveAsPath: '/tmp/data.csv' })
    expect(workbookSaveResultSchema.parse({ canceled: true })).toEqual({ canceled: true })
  })

  it('parses both result variants including the Save-As-xlsx reroute', () => {
    expect(workbookExportCsvResultSchema.parse({ canceled: true })).toEqual({ canceled: true })
    expect(
      workbookExportCsvResultSchema.parse({ canceled: true, saveAsXlsxInstead: true }),
    ).toEqual({ canceled: true, saveAsXlsxInstead: true })
    expect(workbookExportCsvResultSchema.parse({ canceled: false, path: '/tmp/a.csv' })).toEqual({
      canceled: false,
      path: '/tmp/a.csv',
    })
  })
})

describe('Export CSV wiring', () => {
  it('has the File-menu item forwarding export-csv to the renderer', () => {
    const mainSrc = read('src/main/sheets-main.ts')
    expect(mainSrc).toMatch(/menuExportCsv[\s\S]{0,80}sendMenuAction\('export-csv'\)/)
    expect(mainSrc).toContain("action === 'export-csv'")
  })

  it('dispatches the menu action into handleExportCsv', () => {
    const appSrc = read('src/renderer/App.tsx')
    expect(appSrc).toMatch(/action === 'export-csv'[\s\S]{0,80}handleExportCsvImpl/)
  })

  it('allows export-csv through the preload menu-action allowlist', () => {
    const preloadSrc = read('src/preload/index.ts')
    expect(preloadSrc).toContain("action === 'export-csv'")
  })
})

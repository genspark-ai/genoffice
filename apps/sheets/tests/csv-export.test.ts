import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { parseCsv } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { INumfmtService } from '@univerjs/sheets'
import { ConditionalFormattingService } from '@univerjs/preset-sheets-conditional-formatting'

import {
  type CellNumberFormat,
  cellPatternLookup,
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

function fakeCsvSheet(rows: string[][], raws?: unknown[][]) {
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
      getRawValues: () =>
        (raws ?? rows)
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

const explicit =
  (pattern: string) =>
  (_row: number, _column: number, raw: number): CellNumberFormat => ({ pattern, renderValue: raw })

describe('serializeActiveSheetCsv precision', () => {
  it('writes the full model value for General cells whose display was width-rounded', () => {
    // a narrow General column renders 645577.79 as "645577.8": exporting that
    // display string silently loses precision in the written file
    const sheet = fakeCsvSheet([['Debit'], ['645577.8']], [['Debit'], [645577.79]])
    expect(serializeActiveSheetCsv(sheet as never, null, () => null)).toBe('Debit\r\n645577.79\r\n')
  })

  it('keeps the formatted text of explicitly formatted cells (values as displayed)', () => {
    const sheet = fakeCsvSheet([['645577.8', '645,577.79']], [[645577.79, 645577.79]])
    expect(serializeActiveSheetCsv(sheet as never, null, explicit('0.0'))).toBe(
      '645577.8,"645,577.79"\r\n',
    )
  })

  it('leaves non-numeric cells alone even under the default format', () => {
    const sheet = fakeCsvSheet([['TRUE', 'note']], [[true, 'note']])
    expect(serializeActiveSheetCsv(sheet as never, null, () => null)).toBe('TRUE,note\r\n')
  })

  it('re-renders a hash-overflow display from the cell pattern, width-free', () => {
    const sheet = fakeCsvSheet([['########']], [[645577.79]])
    expect(serializeActiveSheetCsv(sheet as never, null, explicit('0.000'))).toBe('645577.790\r\n')
  })

  it('formats a hash-overflow date from the lookup serial (1904 workbooks shift it)', () => {
    const sheet = fakeCsvSheet([['########']], [[1]])
    const lookup = (_r: number, _c: number, raw: number) => ({
      pattern: 'yyyy-mm-dd',
      renderValue: raw + 1462,
    })
    expect(serializeActiveSheetCsv(sheet as never, null, lookup)).toBe('1904-01-02\r\n')
  })

  it('keeps a genuine hash display when the lookup says the value has no rendering', () => {
    const sheet = fakeCsvSheet([['########']], [[-1]])
    const lookup = () => ({ pattern: 'yyyy-mm-dd', renderValue: null })
    expect(serializeActiveSheetCsv(sheet as never, null, lookup)).toBe('########\r\n')
  })

  it('exports the display grid unchanged when no format lookup is supplied', () => {
    const sheet = fakeCsvSheet([['645577.8']], [[645577.79]])
    expect(serializeActiveSheetCsv(sheet as never, null)).toBe('645577.8\r\n')
  })
})

function fakeRuntime(options: { cellPattern?: string; cfPattern?: string; formula?: string }) {
  const services = new Map<unknown, unknown>([
    [
      INumfmtService,
      { getValue: () => (options.cellPattern ? { pattern: options.cellPattern } : null) },
    ],
    [
      ConditionalFormattingService,
      {
        composeStyle: () =>
          options.cfPattern ? { style: { n: { pattern: options.cfPattern } } } : undefined,
      },
    ],
  ])
  const worksheet = {
    getCellRaw: () => (options.formula ? { v: 1, f: options.formula } : { v: 1 }),
  }
  return {
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => {
          const service = services.get(token)
          if (!service) throw new Error('unregistered')
          return service
        },
      }),
    },
    univerAPI: {
      getActiveWorkbook: () => ({
        getId: () => 'wb',
        getSheetBySheetId: () => ({ getSheet: () => worksheet }),
      }),
    },
  } as never
}

describe('cellPatternLookup', () => {
  it('treats a cell with no explicit pattern as General', () => {
    expect(cellPatternLookup(fakeRuntime({}), 's1')!(0, 0, 645577.79)).toBeNull()
    expect(cellPatternLookup(fakeRuntime({ cellPattern: 'General' }), 's1')!(0, 0, 1)).toBeNull()
  })

  it('lets a conditional-formatting pattern win over the cell pattern, as on screen', () => {
    const lookup = cellPatternLookup(fakeRuntime({ cellPattern: '0.0', cfPattern: '0.000' }), 's1')!
    expect(lookup(0, 0, 645577.79)).toEqual({ pattern: '0.000', renderValue: 645577.79 })
  })

  it('reports a conditional-formatting pattern on an otherwise General cell', () => {
    const lookup = cellPatternLookup(fakeRuntime({ cfPattern: '#,##0' }), 's1')!
    expect(lookup(0, 0, 1234.5)).toEqual({ pattern: '#,##0', renderValue: 1234.5 })
  })

  it('shifts file-loaded calendar dates in a 1904 workbook but not formula results', () => {
    const date = fakeRuntime({ cellPattern: 'yyyy-mm-dd' })
    expect(cellPatternLookup(date, 's1', true)!(0, 0, 10)).toEqual({
      pattern: 'yyyy-mm-dd',
      renderValue: 10 + 1462,
    })
    expect(cellPatternLookup(date, 's1', false)!(0, 0, 10)!.renderValue).toBe(10)
    const formula = fakeRuntime({ cellPattern: 'yyyy-mm-dd', formula: '=TODAY()' })
    expect(cellPatternLookup(formula, 's1', true)!(0, 0, 10)!.renderValue).toBe(10)
    const time = fakeRuntime({ cellPattern: 'h:mm' })
    expect(cellPatternLookup(time, 's1', true)!(0, 0, 0.5)!.renderValue).toBe(0.5)
  })

  it('marks a negative 1900-system date as unrenderable so its hashes stay', () => {
    const lookup = cellPatternLookup(fakeRuntime({ cellPattern: 'yyyy-mm-dd' }), 's1')!
    expect(lookup(0, 0, -1)).toEqual({ pattern: 'yyyy-mm-dd', renderValue: null })
    expect(
      cellPatternLookup(fakeRuntime({ cellPattern: '0.00' }), 's1')!(0, 0, -1)!.renderValue,
    ).toBe(-1)
  })

  it('works without a conditional-formatting service', () => {
    const runtime = fakeRuntime({ cellPattern: '0.0' })
    const injector = (runtime as { univer: { __getInjector(): { get(t: unknown): unknown } } })
      .univer
    const original = injector.__getInjector
    injector.__getInjector = () => ({
      get: (token: unknown) => {
        if (token === ConditionalFormattingService) throw new Error('missing')
        return original().get(token)
      },
    })
    expect(cellPatternLookup(runtime, 's1')!(0, 0, 2)).toEqual({ pattern: '0.0', renderValue: 2 })
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

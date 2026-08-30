import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  executeWorkbookTool,
  type ActiveSheetInfo,
  type CreateDocumentToolOutcome,
  type SheetsSkillDeps,
  type ToolExecution,
} from '../src/renderer/ai/tools'
import { createAiDocument, type AiCreateDocumentContext } from '../src/renderer/ai/create-document'
import type { CsvWorksheet } from '../src/renderer/csv-export'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

function call(name: string, input: Record<string, unknown>) {
  return { id: 'call-1', name, input }
}

function fakeDeps(overrides: Partial<SheetsSkillDeps> = {}): SheetsSkillDeps {
  const info: ActiveSheetInfo = {
    mode: 'demo',
    sheetId: 'sheet-1',
    sheetName: 'Sheet1',
    revision: 0,
    knownAddresses: ['A1'],
    sheets: [
      { id: 'sheet-1', name: 'Sheet1' },
      { id: 'sheet-2', name: 'ja' },
    ],
  }
  return {
    getActiveSheetInfo: () => info,
    readCells: () => ({}),
    readFormats: () => ({}),
    readSheetFeatures: () => '',
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    selectRange: () => ({ ok: true, sheetName: 'Sheet1' }),
    tracePrecedents: () => ({ refs: [] }),
    traceDependents: () => ({ dependents: [], truncated: false, incompleteSheets: [] }),
    proposeOperations: () => ({ ok: false, error: 'not configured' }),
    ...overrides,
  }
}

async function exec(
  input: Record<string, unknown>,
  createDocument: NonNullable<SheetsSkillDeps['createDocument']>,
): Promise<ToolExecution> {
  return executeWorkbookTool(call('create_document', input), fakeDeps({ createDocument }))
}

describe('create_document tool (executor)', () => {
  it('defaults to xlsx on the active sheet', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'Sheet1.xlsx',
      path: '/tmp/Sheet1.xlsx',
      sheetName: 'Sheet1',
    }))
    const result = await exec({}, createDocument)
    expect(result.isError).toBeUndefined()
    expect(createDocument).toHaveBeenCalledWith({ type: 'xlsx', sheetId: undefined })
    expect(result.output).toContain('Sheet1.xlsx')
    expect(result.output).toContain('unchanged')
    expect(result.mutated).toBe(false)
  })

  it('forwards csv with sheetId and title', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'locale-ja.csv',
      path: '/tmp/locale-ja.csv',
      sheetName: 'ja',
    }))
    const result = await exec(
      { type: 'csv', sheetId: 'sheet-2', title: 'locale-ja' },
      createDocument,
    )
    expect(result.isError).toBeUndefined()
    expect(createDocument).toHaveBeenCalledWith({
      type: 'csv',
      sheetId: 'sheet-2',
      title: 'locale-ja',
    })
  })

  it('appends the formula-loss note when the sheet has formulas', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'Sheet1.csv',
      sheetName: 'Sheet1',
      hadFormulas: true,
    }))
    const result = await exec({ type: 'csv' }, createDocument)
    expect(result.output).toContain('computed values only')
  })

  it('rejects an unknown sheetId without calling the bridge', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'x',
    }))
    const result = await exec({ type: 'csv', sheetId: 'nope' }, createDocument)
    expect(result.isError).toBe(true)
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('rejects content for xlsx/csv with the write-to-a-sheet-first hint', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'x',
    }))
    const result = await exec({ type: 'csv', content: 'a,b\n1,2' }, createDocument)
    expect(result.isError).toBe(true)
    expect(result.output).toContain('add_sheet')
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('requires title and content for docx/pdf/md', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'x',
    }))
    for (const input of [
      { type: 'docx', content: '<p>x</p>' },
      { type: 'pdf', title: 'T' },
      { type: 'md', title: '  ', content: 'x' },
    ]) {
      const result = await exec(input, createDocument)
      expect(result.isError).toBe(true)
    }
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('rejects echoed tool-protocol output as docx/pdf content', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'x',
    }))
    const result = await exec(
      { type: 'pdf', title: 'T', content: '<tool_response>…</tool_response>' },
      createDocument,
    )
    expect(result.isError).toBe(true)
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('skips the echo guard for markdown content', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'notes.md',
      path: '/tmp/notes.md',
    }))
    const result = await exec(
      { type: 'md', title: 'notes', content: '{"index": 1, "type": "note"} raw dump' },
      createDocument,
    )
    expect(result.isError).toBeUndefined()
    expect(createDocument).toHaveBeenCalledOnce()
  })

  it('surfaces a bridge failure', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: false,
      error: 'disk full',
    }))
    const result = await exec({ type: 'docx', title: 'T', content: '<p>x</p>' }, createDocument)
    expect(result.isError).toBe(true)
    expect(result.output).toContain('disk full')
  })

  it('rejects an unknown type', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentToolOutcome> => ({
      ok: true,
      name: 'x',
    }))
    const result = await exec({ type: 'pptx', title: 'T', content: 'x' }, createDocument)
    expect(result.isError).toBe(true)
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('fails gracefully when the dep is not wired', async () => {
    const result = await executeWorkbookTool(call('create_document', {}), fakeDeps())
    const settled = result instanceof Promise ? await result : result
    expect(settled.isError).toBe(true)
    expect(settled.output).toContain('not available')
  })
})

// ---- renderer bridge: createAiDocument (serialization + availability) ----

interface FakeSheetOptions {
  id?: string
  name?: string
  hasFormulas?: boolean
}

function fakeSheet(rows: string[][], options: FakeSheetOptions = {}): CsvWorksheet {
  const { id = 's-live', name = 'Data', hasFormulas = false } = options
  return {
    getLastRow: () => rows.length - 1,
    getLastColumn: () => Math.max(...rows.map((row) => row.length)) - 1,
    getSheetId: () => id,
    getSheetName: () => name,
    getRange: (row, column, numRows, numColumns) => ({
      getDisplayValues: () =>
        rows.slice(row, row + numRows).map((sourceRow) => {
          const out: string[] = []
          for (let c = column; c < column + numColumns; c += 1) out.push(sourceRow[c] ?? '')
          return out
        }),
    }),
    getSheet: () => ({
      getCellMatrix: () => ({
        forValue: (callback) => {
          if (hasFormulas) callback(0, 0, { f: '=SUM(A1)' })
        },
      }),
    }),
  }
}

function fakeContext(
  sheets: CsvWorksheet[],
  state: LazyWorkbookState | null = null,
): AiCreateDocumentContext {
  const univerAPI = {
    getActiveWorkbook: () => ({
      getActiveSheet: () => sheets[0],
      getSheetBySheetId: (id: string) => sheets.find((sheet) => sheet.getSheetId() === id) ?? null,
      getSheets: () => sheets,
    }),
  }
  return {
    univerRef: { current: { univerAPI } as unknown as UniverRuntime },
    lazyWorkbookRef: { current: state },
  }
}

function lazyState(fileSheetIds: string[], preloadComplete: boolean): LazyWorkbookState {
  return {
    flags: { preloadComplete },
    file: { sheets: fileSheetIds.map((id) => ({ id })) },
    showFormulaSheets: new Set<string>(),
  } as unknown as LazyWorkbookState
}

const createDocumentBridge = vi.fn()

beforeEach(() => {
  createDocumentBridge.mockReset()
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { createDocument: createDocumentBridge },
  }
})

describe('createAiDocument (renderer bridge)', () => {
  it('serializes the sheet as CSV and defaults the title to its name', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/ja.csv' })
    const outcome = await createAiDocument(
      fakeContext([
        fakeSheet(
          [
            ['a', 'b'],
            ['1', '2'],
          ],
          { name: 'ja' },
        ),
      ]),
      { type: 'csv' },
    )
    expect(createDocumentBridge).toHaveBeenCalledWith({
      type: 'csv',
      title: 'ja',
      content: 'a,b\r\n1,2\r\n',
    })
    expect(outcome).toMatchObject({
      ok: true,
      name: 'ja.csv',
      path: '/tmp/ja.csv',
      sheetName: 'ja',
    })
  })

  it('passes the worksheet name through for xlsx', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/out.xlsx' })
    const outcome = await createAiDocument(fakeContext([fakeSheet([['x']], { name: 'de' })]), {
      type: 'xlsx',
      title: 'winback_de',
    })
    expect(createDocumentBridge).toHaveBeenCalledWith({
      type: 'xlsx',
      title: 'winback_de',
      content: 'x\r\n',
      sheetName: 'de',
    })
    expect(outcome).toMatchObject({ ok: true, name: 'winback_de.xlsx' })
  })

  it('reports formulas so the tool can warn about value-only export', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/f.csv' })
    const outcome = await createAiDocument(
      fakeContext([fakeSheet([['1']], { hasFormulas: true })]),
      { type: 'csv' },
    )
    expect(outcome).toMatchObject({ ok: true, hadFormulas: true })
  })

  it('resolves a non-active sheet by id', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/second.csv' })
    const second = fakeSheet([['s2']], { id: 's-2', name: 'Second' })
    const outcome = await createAiDocument(
      fakeContext([fakeSheet([['s1']], { id: 's-1' }), second]),
      { type: 'csv', sheetId: 's-2' },
    )
    expect(createDocumentBridge).toHaveBeenCalledWith({
      type: 'csv',
      title: 'Second',
      content: 's2\r\n',
    })
    expect(outcome).toMatchObject({ ok: true, sheetName: 'Second' })
  })

  it('errors on an unknown sheetId', async () => {
    const outcome = await createAiDocument(fakeContext([fakeSheet([['x']])]), {
      type: 'csv',
      sheetId: 'missing',
    })
    expect(outcome.ok).toBe(false)
    expect(createDocumentBridge).not.toHaveBeenCalled()
  })

  it('refuses a file-backed sheet before the workbook is fully loaded', async () => {
    const outcome = await createAiDocument(
      fakeContext([fakeSheet([['x']], { id: 's-file' })], lazyState(['s-file'], false)),
      { type: 'csv' },
    )
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('copy_range')
    expect(createDocumentBridge).not.toHaveBeenCalled()
  })

  it('exports a session-added sheet even while the file streams', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/ja.csv' })
    const outcome = await createAiDocument(
      fakeContext(
        [fakeSheet([['ja-row']], { id: 's-added', name: 'ja' })],
        lazyState(['s-file'], false),
      ),
      { type: 'csv' },
    )
    expect(outcome).toMatchObject({ ok: true, name: 'ja.csv' })
    expect(createDocumentBridge).toHaveBeenCalledOnce()
  })

  it('exports a file-backed sheet once preload completes', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true, path: '/tmp/Data.csv' })
    const outcome = await createAiDocument(
      fakeContext([fakeSheet([['x']], { id: 's-file' })], lazyState(['s-file'], true)),
      { type: 'csv' },
    )
    expect(outcome).toMatchObject({ ok: true })
  })

  it('forwards docx content untouched', async () => {
    createDocumentBridge.mockResolvedValue({ ok: true })
    const outcome = await createAiDocument(fakeContext([]), {
      type: 'docx',
      title: 'Report',
      content: '<h1>R</h1><p>body</p>',
    })
    expect(createDocumentBridge).toHaveBeenCalledWith({
      type: 'docx',
      title: 'Report',
      content: '<h1>R</h1><p>body</p>',
    })
    expect(outcome).toMatchObject({ ok: true, name: 'Report.docx' })
  })
})

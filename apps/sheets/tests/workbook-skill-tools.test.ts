import { describe, expect, it, vi } from 'vitest'

import {
  buildWorkbookContext,
  executeWorkbookTool,
  type ActiveSheetInfo,
  type SheetsSkillDeps,
  type ToolExecution,
} from '../src/renderer/ai/tools'
import type { ChangePlan } from '../src/domain/workbook.types'

function call(name: string, input: Record<string, unknown>) {
  return { id: 'call-1', name, input }
}

/** Only propose_operations' formula read-back branch is async; sync-asserted cases stay sync. */
function execSync(c: ReturnType<typeof call>, d: SheetsSkillDeps): ToolExecution {
  const result = executeWorkbookTool(c, d)
  if (result instanceof Promise) throw new Error('expected sync tool execution')
  return result
}

function fakeDeps(overrides: Partial<SheetsSkillDeps> = {}): SheetsSkillDeps {
  const info: ActiveSheetInfo = {
    mode: 'demo',
    sheetId: 'sheet-1',
    sheetName: 'Sheet1',
    revision: 0,
    knownAddresses: ['A1', 'B1'],
    sheets: [
      { id: 'sheet-1', name: 'Sheet1' },
      { id: 'sheet-2', name: 'Summary' },
    ],
    selection: 'B2:C4',
  }
  return {
    getActiveSheetInfo: () => info,
    readCells: () => ({}),
    readFormats: () => ({}),
    readSheetFeatures: () => 'Feature state of sheet Sheet1 (id=sheet-1):\nAutoFilter: none',
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    selectRange: () => ({ ok: true, sheetName: 'Sheet1' }),
    tracePrecedents: () => ({ refs: [] }),
    traceDependents: () => ({ dependents: [], truncated: false, incompleteSheets: [] }),
    proposeOperations: () => ({ ok: false, error: 'not configured' }),
    ...overrides,
  }
}

const EMPTY_PLAN: ChangePlan = {
  transactionId: 'agent-1',
  baseRevision: 0,
  cellChanges: [
    { sheetId: 'sheet-1', address: 'A1', before: { value: 'old' }, after: { value: 'new' } },
  ],
  sheetRenames: [],
  structuralChanges: [],
  formatChanges: [],
  warnings: [],
}

describe('buildWorkbookContext', () => {
  it('reports no open workbook', () => {
    const deps = fakeDeps({
      getActiveSheetInfo: () => ({
        mode: 'none',
        sheetId: '',
        sheetName: '',
        knownAddresses: [],
        sheets: [],
      }),
    })
    expect(buildWorkbookContext(deps)).toContain('No workbook is currently open')
  })

  it('reports the demo sheet, revision, and known addresses', () => {
    const text = buildWorkbookContext(fakeDeps())
    expect(text).toContain('Sheet1')
    expect(text).toContain('sheet-1')
    expect(text).toContain('revision=0')
    expect(text).toContain('A1')
  })

  it('lists all sheets and the current selection', () => {
    const text = buildWorkbookContext(fakeDeps())
    expect(text).toContain('Summary (id=sheet-2)')
    expect(text).toContain('Current selection: B2:C4')
  })

  it('lists merged ranges and charts when present', () => {
    const deps = fakeDeps({
      getActiveSheetInfo: () => ({
        mode: 'lazy',
        sheetId: 'sheet-1',
        sheetName: 'Data',
        knownAddresses: [],
        sheets: [{ id: 'sheet-1', name: 'Data' }],
        merges: ['A1:C1'],
        charts: [
          { path: 'xl/charts/chart1.xml', title: 'Revenue', types: 'column', sheetId: 'sheet-1' },
        ],
      }),
    })
    const text = buildWorkbookContext(deps)
    expect(text).toContain('Merged ranges on the active sheet: A1:C1')
    expect(text).toContain('xl/charts/chart1.xml')
    expect(text).toContain('Revenue')
  })

  it('reports per-sheet data dimensions when known', () => {
    const deps = fakeDeps({
      getActiveSheetInfo: () => ({
        mode: 'lazy',
        sheetId: 'sheet-1',
        sheetName: 'Data',
        knownAddresses: [],
        sheets: [
          { id: 'sheet-1', name: 'Data', rows: 1006, columns: 17 },
          { id: 'sheet-2', name: 'Summary', rows: 20, columns: 3 },
        ],
      }),
    })
    const text = buildWorkbookContext(deps)
    expect(text).toContain('1006 rows × 17 columns')
    expect(text).toContain('A1:Q1006')
    expect(text).toContain('20 rows × 3 columns')
  })

  it('reports a lazy (imported) sheet without a revision', () => {
    const deps = fakeDeps({
      getActiveSheetInfo: () => ({
        mode: 'lazy',
        sheetId: 'sheet-2',
        sheetName: 'Budget',
        knownAddresses: [],
        sheets: [{ id: 'sheet-2', name: 'Budget' }],
        loadedRange: 'A80:E160',
      }),
    })
    const text = buildWorkbookContext(deps)
    expect(text).toContain('Budget')
    expect(text).toContain('streaming in')
    expect(text).toContain('Currently loaded viewport: A80:E160 (not the worksheet data extent)')
  })
})

describe('executeWorkbookTool: read_range', () => {
  it('rejects a missing or unparsable range', () => {
    expect(execSync(call('read_range', {}), fakeDeps()).isError).toBe(true)
    expect(execSync(call('read_range', { range: 'nope' }), fakeDeps()).isError).toBe(true)
  })

  it('rejects ranges above the cell ceiling', () => {
    const result = execSync(call('read_range', { range: 'A1:Z100' }), fakeDeps())
    expect(result.isError).toBe(true)
    expect(result.output).toContain('2000')
  })

  it('rejects a range outside the authoritative worksheet extent before loading it', () => {
    const ensureRangeLoaded = vi.fn()
    const readCells = vi.fn()
    const result = execSync(
      call('read_range', { range: 'A1:F10' }),
      fakeDeps({
        getActiveSheetInfo: () => ({
          mode: 'lazy',
          sheetId: 'sheet-1',
          sheetName: 'Data',
          knownAddresses: [],
          sheets: [{ id: 'sheet-1', name: 'Data', rows: 10, columns: 5 }],
        }),
        ensureRangeLoaded,
        readCells,
      }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('outside the worksheet data extent A1:E10')
    expect(ensureRangeLoaded).not.toHaveBeenCalled()
    expect(readCells).not.toHaveBeenCalled()
  })

  it('returns a labeled grid with formulas and blanks', () => {
    const readCells = vi.fn().mockReturnValue({
      A1: { value: 'Name' },
      B1: { value: 'Total' },
      B2: { value: null, formula: '=SUM(C1:C9)' },
    })
    const result = execSync(call('read_range', { range: 'A1:B2' }), fakeDeps({ readCells }))
    expect(readCells).toHaveBeenCalledWith(['A1', 'B1', 'A2', 'B2'], undefined)
    const lines = result.output.split('\n')
    expect(lines[0]).toContain('requested range A1:B2')
    expect(lines[0]).toContain('Do not infer total rows')
    expect(lines[1]).toBe('\tA\tB')
    expect(lines[2]).toBe('1\tName\tTotal')
    expect(lines[3]).toBe('2\t\t=SUM(C1:C9)')
    expect(result.mutated).toBe(false)
  })

  it('reads a non-active sheet when sheetId is given', () => {
    const readCells = vi.fn().mockReturnValue({ A1: { value: 'from summary' } })
    const result = execSync(
      call('read_range', { range: 'A1', sheetId: 'sheet-2' }),
      fakeDeps({ readCells }),
    )
    expect(readCells).toHaveBeenCalledWith(['A1'], 'sheet-2')
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('from summary')
  })

  it('rejects an unknown sheetId with a clear error', () => {
    const readCells = vi.fn()
    const result = execSync(
      call('read_range', { range: 'A1', sheetId: 'nope' }),
      fakeDeps({ readCells }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: nope')
    expect(readCells).not.toHaveBeenCalled()
  })

  it('escapes control characters in cell text so each grid row stays one physical line', () => {
    const readCells = vi.fn().mockReturnValue({
      A1: { value: 'Question' },
      B1: { value: 'Answer' },
      A2: { value: 'How did you\ndo it?' },
      B2: { value: 'VO: line one\n\nline two\twith tab and C:\\path' },
    })
    const result = execSync(call('read_range', { range: 'A1:B2' }), fakeDeps({ readCells }))
    const lines = result.output.split('\n')
    expect(lines).toHaveLength(4)
    expect(lines[2]).toBe('1\tQuestion\tAnswer')
    expect(lines[3]).toBe(
      '2\tHow did you\\ndo it?\tVO: line one\\n\\nline two\\twith tab and C:\\\\path',
    )
  })

  it('normalizes Univer \\r paragraph breaks and CRLF to the \\n escape', () => {
    // Univer streams in-cell line breaks as \r; the model must only ever see
    // the \n convention documented in the base prompt.
    const readCells = vi.fn().mockReturnValue({
      A1: { value: 'univer\rparagraph' },
      B1: { value: 'windows\r\nbreak and lone\rcr' },
    })
    const result = execSync(call('read_range', { range: 'A1:B1' }), fakeDeps({ readCells }))
    const lines = result.output.split('\n')
    expect(lines[2]).toBe('1\tuniver\\nparagraph\twindows\\nbreak and lone\\ncr')
    expect(result.output).not.toContain('\\r')
  })

  it('reports the authoritative sheet extent separately from the requested range', () => {
    const result = execSync(
      call('read_range', { range: 'A1:E120' }),
      fakeDeps({
        getActiveSheetInfo: () => ({
          mode: 'lazy',
          sheetId: 'sheet-1',
          sheetName: 'Data',
          knownAddresses: [],
          sheets: [{ id: 'sheet-1', name: 'Data', rows: 201, columns: 5 }],
        }),
      }),
    )
    expect(result.output).toContain('requested range A1:E120')
    expect(result.output).toContain('authoritative worksheet data extent A1:E201')
    expect(result.output).toContain('201 worksheet rows')
  })

  it('waits for a lazy range to load before reading cells', async () => {
    const events: string[] = []
    const result = await executeWorkbookTool(
      call('read_range', { range: 'A1:B2' }),
      fakeDeps({
        ensureRangeLoaded: async () => {
          events.push('loaded')
          return true
        },
        readCells: () => {
          events.push('read')
          return {}
        },
      }),
    )
    expect(result.isError).toBeFalsy()
    expect(events).toEqual(['loaded', 'read'])
  })

  it('fails instead of treating cells as blank when a lazy range cannot load', async () => {
    const readCells = vi.fn()
    const result = await executeWorkbookTool(
      call('read_range', { range: 'A1:B2' }),
      fakeDeps({
        ensureRangeLoaded: async () => false,
        readCells,
      }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('could not be fully loaded')
    expect(readCells).not.toHaveBeenCalled()
  })

  it('accepts a single-cell range in lowercase', () => {
    const readCells = vi.fn().mockReturnValue({ B2: { value: 7 } })
    const result = execSync(call('read_range', { range: 'b2' }), fakeDeps({ readCells }))
    expect(result.output).toContain('7')
    expect(result.isError).toBeFalsy()
  })
})

describe('executeWorkbookTool: get_workbook_context', () => {
  it('never mutates', () => {
    const result = execSync(call('get_workbook_context', {}), fakeDeps())
    expect(result.mutated).toBe(false)
    expect(result.output).toContain('Sheet1')
  })
})

describe('executeWorkbookTool: load_guide', () => {
  it('loads one or more guides by name', () => {
    const result = execSync(call('load_guide', { guides: ['writing', 'structure'] }), fakeDeps())
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('set_range')
    expect(result.output).toContain('insert_rows')
    expect(result.mutated).toBe(false)
  })

  it('rejects unknown guide names, listing the valid ones', () => {
    // 'pivot' became a real guide — use a name that stays unregistered.
    const result = execSync(call('load_guide', { guides: ['no-such-guide'] }), fakeDeps())
    expect(result.isError).toBe(true)
    expect(result.output).toContain('writing')
    expect(execSync(call('load_guide', {}), fakeDeps()).isError).toBe(true)
  })
})

describe('executeWorkbookTool: read_formats', () => {
  it('lists only cells with explicit formats', () => {
    const readFormats = vi.fn().mockReturnValue({
      A1: { bold: true, fillColor: '#FFF2CC' },
      B2: { numberFormat: '0.00%', border: { type: 'all' } },
    })
    const result = execSync(call('read_formats', { range: 'A1:B2' }), fakeDeps({ readFormats }))
    expect(readFormats).toHaveBeenCalledWith(['A1', 'B1', 'A2', 'B2'], undefined)
    expect(result.output).toContain('A1: bold, fill #FFF2CC')
    expect(result.output).toContain('B2: number format 0.00%, border all')
  })

  it('reports an unformatted range and rejects oversized ones', () => {
    const empty = execSync(call('read_formats', { range: 'A1:B2' }), fakeDeps())
    expect(empty.output).toContain('No explicit formats')
    const oversized = execSync(call('read_formats', { range: 'A1:Z100' }), fakeDeps())
    expect(oversized.isError).toBe(true)
  })
})

describe('executeWorkbookTool: read_sheet_features', () => {
  it('returns the feature report and forwards the sheetId', () => {
    const seen: (string | undefined)[] = []
    const deps = fakeDeps({
      readSheetFeatures: (sheetId) => {
        seen.push(sheetId)
        return 'Status: visible, unprotected'
      },
    })
    const result = execSync({ id: '1', name: 'read_sheet_features', input: {} }, deps)
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('unprotected')
    execSync({ id: '2', name: 'read_sheet_features', input: { sheetId: 'sheet-2' } }, deps)
    expect(seen).toEqual([undefined, 'sheet-2'])
  })

  it('rejects an unknown sheetId before reading', () => {
    const readSheetFeatures = vi.fn()
    const result = execSync(
      { id: '1', name: 'read_sheet_features', input: { sheetId: 'nope' } },
      fakeDeps({ readSheetFeatures }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: nope')
    expect(readSheetFeatures).not.toHaveBeenCalled()
  })

  it('streams the sheet in before reading grid-backed feature models', async () => {
    const readSheetFeatures = vi.fn().mockReturnValue('AutoFilter: A1:C9')
    const ensureRangeLoaded = vi.fn().mockResolvedValue(true)
    const result = await executeWorkbookTool(
      { id: '1', name: 'read_sheet_features', input: { sheetId: 'sheet-2' } },
      fakeDeps({ readSheetFeatures, ensureRangeLoaded }),
    )
    expect(ensureRangeLoaded).toHaveBeenCalledWith(
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      'sheet-2',
    )
    expect(result.output).toContain('AutoFilter: A1:C9')
  })

  it('fails closed when the sheet cannot be streamed in', async () => {
    const readSheetFeatures = vi.fn()
    const ensureRangeLoaded = vi.fn().mockResolvedValue(false)
    const result = await executeWorkbookTool(
      { id: '1', name: 'read_sheet_features', input: {} },
      fakeDeps({ readSheetFeatures, ensureRangeLoaded }),
    )
    expect(result.isError).toBe(true)
    expect(readSheetFeatures).not.toHaveBeenCalled()
  })
})

describe('executeWorkbookTool: read_cells', () => {
  it('rejects a missing addresses array', () => {
    const result = execSync(call('read_cells', {}), fakeDeps())
    expect(result.isError).toBe(true)
  })

  it('formats values and formulas for each requested address', () => {
    const readCells = vi.fn().mockReturnValue({
      A1: { value: 42 },
      B1: { value: null, formula: '=SUM(A1:A10)' },
    })
    const result = execSync(
      call('read_cells', { addresses: ['A1', 'B1', 'C1'] }),
      fakeDeps({ readCells }),
    )
    expect(readCells).toHaveBeenCalledWith(['A1', 'B1', 'C1'], undefined)
    expect(result.output).toContain('A1: 42')
    expect(result.output).toContain('B1: =SUM(A1:A10)')
    expect(result.output).toContain('C1: (unknown)')
    expect(result.mutated).toBe(false)
  })

  it('ensures the covering range of the addresses is streamed in before reading', async () => {
    const readCells = vi.fn().mockReturnValue({ B2: { value: 1 }, D5: { value: 2 } })
    const ensureRangeLoaded = vi.fn().mockResolvedValue(true)
    const result = await executeWorkbookTool(
      call('read_cells', { addresses: ['B2', 'D5'], sheetId: 'sheet-2' }),
      fakeDeps({ readCells, ensureRangeLoaded }),
    )
    expect(ensureRangeLoaded).toHaveBeenCalledWith(
      { startRow: 1, endRow: 4, startColumn: 1, endColumn: 3 },
      'sheet-2',
    )
    expect(result.output).toContain('B2: 1')
    expect(readCells).toHaveBeenCalledWith(['B2', 'D5'], 'sheet-2')
  })

  it('reads wide scatters when no streaming is needed (ensure returns true)', async () => {
    const readCells = vi.fn().mockReturnValue({ A1: { value: 1 }, ZZ9999: { value: 2 } })
    const ensureRangeLoaded = vi.fn().mockResolvedValue(true)
    const result = await executeWorkbookTool(
      call('read_cells', { addresses: ['A1', 'ZZ9999'] }),
      fakeDeps({ readCells, ensureRangeLoaded }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('A1: 1')
    expect(result.output).toContain('ZZ9999: 2')
  })

  it('tells the model to cluster its reads when the covering box cannot load', async () => {
    const readCells = vi.fn()
    const ensureRangeLoaded = vi.fn().mockResolvedValue(false)
    const result = await executeWorkbookTool(
      call('read_cells', { addresses: ['A1', 'ZZ9999'] }),
      fakeDeps({ readCells, ensureRangeLoaded }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('closer-together read_cells calls')
    expect(readCells).not.toHaveBeenCalled()
  })

  it('fails closed when the requested cells cannot be loaded', async () => {
    const readCells = vi.fn()
    const ensureRangeLoaded = vi.fn().mockResolvedValue(false)
    const result = await executeWorkbookTool(
      call('read_cells', { addresses: ['B2'] }),
      fakeDeps({ readCells, ensureRangeLoaded }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('could not be fully loaded')
    expect(readCells).not.toHaveBeenCalled()
  })

  it('formula cell with a computed value displays "value (formula)"', () => {
    const readCells = vi.fn().mockReturnValue({
      B1: { value: 550, formula: '=SUM(A1:A10)' },
      C1: { value: '#REF!', formula: '=D1+E1' },
    })
    const result = execSync(
      call('read_cells', { addresses: ['B1', 'C1'] }),
      fakeDeps({ readCells }),
    )
    expect(result.output).toContain('B1: 550 (=SUM(A1:A10))')
    expect(result.output).toContain('C1: #REF! (=D1+E1)')
  })

  it('escapes multi-line cell text so each address stays on its own line', () => {
    const readCells = vi.fn().mockReturnValue({
      A1: { value: 'first\nsecond' },
      B1: { value: 7 },
    })
    const result = execSync(
      call('read_cells', { addresses: ['A1', 'B1'] }),
      fakeDeps({ readCells }),
    )
    expect(result.output.split('\n')).toEqual(['A1: first\\nsecond', 'B1: 7'])
  })
})

describe('executeWorkbookTool: propose_operations', () => {
  it('rejects an empty operations array', () => {
    const result = execSync(
      call('propose_operations', { operations: [], summary: 'x' }),
      fakeDeps(),
    )
    expect(result.isError).toBe(true)
  })

  it('rejects a blank summary', () => {
    const result = execSync(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 1 }],
        summary: '   ',
      }),
      fakeDeps(),
    )
    expect(result.isError).toBe(true)
  })

  it('rejects operations that fail the DSL schema without calling proposeOperations', () => {
    const proposeOperations = vi.fn()
    const result = execSync(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'not-a-cell', value: 1 }],
        summary: 'bad address',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(result.isError).toBe(true)
    expect(proposeOperations).not.toHaveBeenCalled()
  })

  it('forwards validated operations and reports auto-applied success', () => {
    const proposeOperations = vi.fn().mockReturnValue({ ok: true, plan: EMPTY_PLAN })
    const result = execSync(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'new' }],
        summary: 'Update A1',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(proposeOperations).toHaveBeenCalledWith(
      [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'new' }],
      'Update A1',
    )
    expect(result.mutated).toBe(true)
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('old → new')
    expect(result.output).toContain('Auto-applied')
    expect(result.output).toContain('Undo')
  })

  it('after writing a formula, reads back the computed value asynchronously (write → verify)', async () => {
    const plan: ChangePlan = {
      ...EMPTY_PLAN,
      cellChanges: [
        {
          sheetId: 'sheet-1',
          address: 'B4',
          before: { value: null },
          after: { value: null, formula: '=SUM(B1:B3)' },
        },
      ],
    }
    const proposeOperations = vi.fn().mockReturnValue({ ok: true, plan })
    const readCells = vi.fn().mockReturnValue({ B4: { value: 60, formula: '=SUM(B1:B3)' } })
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [
          { op: 'set_formula', sheetId: 'sheet-1', address: 'B4', formula: '=SUM(B1:B3)' },
        ],
        summary: 'Sum B',
      }),
      fakeDeps({ proposeOperations, readCells }),
    )
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('Formula results: B4 = 60')
    expect(readCells).toHaveBeenCalledWith(['B4'], 'sheet-1')
  })

  it('reads back formulas per target sheet, labeling addresses when sheets mix', async () => {
    const plan: ChangePlan = {
      ...EMPTY_PLAN,
      cellChanges: [
        {
          sheetId: 'sheet-1',
          address: 'B4',
          before: { value: null },
          after: { value: null, formula: '=SUM(B1:B3)' },
        },
        {
          sheetId: 'sheet-2',
          address: 'C1',
          before: { value: null },
          after: { value: null, formula: '=Sheet1!B4*2' },
        },
      ],
    }
    const proposeOperations = vi.fn().mockReturnValue({ ok: true, plan })
    const readCells = vi
      .fn()
      .mockImplementation((addresses: string[], sheetId?: string) =>
        sheetId === 'sheet-2' ? { C1: { value: 120 } } : { B4: { value: 60 } },
      )
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [
          { op: 'set_formula', sheetId: 'sheet-1', address: 'B4', formula: '=SUM(B1:B3)' },
          { op: 'set_formula', sheetId: 'sheet-2', address: 'C1', formula: '=Sheet1!B4*2' },
        ],
        summary: 'Cross-sheet sums',
      }),
      fakeDeps({ proposeOperations, readCells }),
    )
    expect(readCells).toHaveBeenCalledWith(['B4'], 'sheet-1')
    expect(readCells).toHaveBeenCalledWith(['C1'], 'sheet-2')
    expect(result.output).toContain('Sheet1!B4 = 60')
    expect(result.output).toContain('Summary!C1 = 120')
  })

  it('warns explicitly when read-back finds a formula error value', async () => {
    const plan: ChangePlan = {
      ...EMPTY_PLAN,
      cellChanges: [
        {
          sheetId: 'sheet-1',
          address: 'C1',
          before: { value: null },
          after: { value: null, formula: '=A1/A2' },
        },
      ],
    }
    const proposeOperations = vi.fn().mockReturnValue({ ok: true, plan })
    const readCells = vi.fn().mockReturnValue({ C1: { value: '#DIV/0!', formula: '=A1/A2' } })
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [{ op: 'set_formula', sheetId: 'sheet-1', address: 'C1', formula: '=A1/A2' }],
        summary: 'Divide',
      }),
      fakeDeps({ proposeOperations, readCells }),
    )
    expect(result.output).toContain('#DIV/0!')
    expect(result.output).toContain('⚠️ Formula error values present')
  })

  it('waits for the async apply and reports success only after it lands', async () => {
    const proposeOperations = vi.fn().mockReturnValue({
      ok: true,
      plan: EMPTY_PLAN,
      applied: Promise.resolve({ ok: true }),
    })
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'new' }],
        summary: 'Update A1',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.mutated).toBe(true)
    expect(result.output).toContain('Auto-applied')
  })

  it('returns an error (not success) when the async apply fails', async () => {
    const proposeOperations = vi.fn().mockReturnValue({
      ok: true,
      plan: EMPTY_PLAN,
      applied: Promise.resolve({ ok: false, reason: 'workbook changed since preview' }),
    })
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'new' }],
        summary: 'Update A1',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(result.isError).toBe(true)
    expect(result.mutated).toBe(false)
    expect(result.output).toContain('UNCHANGED')
    expect(result.output).toContain('workbook changed since preview')
  })

  it('reports a mid-batch failure as partially committed, never as unchanged', async () => {
    const proposeOperations = vi.fn().mockReturnValue({
      ok: true,
      plan: EMPTY_PLAN,
      applied: Promise.resolve({ ok: false, reason: 'Unknown chart', partiallyApplied: true }),
    })
    const result = await executeWorkbookTool(
      call('propose_operations', {
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'new' }],
        summary: 'Update A1',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('MID-BATCH')
    expect(result.output).toContain('already committed')
    expect(result.output).not.toContain('UNCHANGED')
  })

  it('propagates a conflict/streaming-guard error from proposeOperations', () => {
    const proposeOperations = vi.fn().mockReturnValue({ ok: false, error: 'still streaming in' })
    const result = execSync(
      call('propose_operations', {
        operations: [{ op: 'clear_cell', sheetId: 'sheet-1', address: 'A1' }],
        summary: 'Clear A1',
      }),
      fakeDeps({ proposeOperations }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toBe('still streaming in')
  })
})

describe('executeWorkbookTool: find_cells', () => {
  it('rejects a missing query unless errors_only is set', () => {
    const result = execSync(call('find_cells', {}), fakeDeps())
    expect(result.isError).toBe(true)
    expect(result.output).toContain('query')
  })

  it('allows an omitted query with errors_only=true', () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [{ sheetName: 'Sheet1', address: 'C3', value: '#REF!', formula: '=A1/B1' }],
      truncated: false,
      incompleteSheets: [],
    })
    const result = execSync(call('find_cells', { errors_only: true }), fakeDeps({ findCells }))
    expect(result.isError).toBeUndefined()
    expect(findCells).toHaveBeenCalledWith(
      expect.objectContaining({ errorsOnly: true, lookIn: 'both', maxResults: 50 }),
    )
    expect(result.output).toContain('Sheet1!C3: #REF! (=A1/B1)')
  })

  it('normalizes look_in, clamps max_results, and forwards sheetId', () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [],
      truncated: false,
      incompleteSheets: [],
    })
    execSync(
      call('find_cells', {
        query: 'total',
        look_in: 'formulas',
        max_results: 9999,
        sheetId: 'sheet-2',
      }),
      fakeDeps({ findCells }),
    )
    expect(findCells).toHaveBeenCalledWith(
      expect.objectContaining({
        query: 'total',
        lookIn: 'formulas',
        maxResults: 200,
        sheetId: 'sheet-2',
      }),
    )
  })

  it('reports matches with truncation and indexing notes', () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [
        { sheetName: 'Sheet1', address: 'A1', value: 'Total' },
        { sheetName: 'Summary', address: 'B2', value: 120, formula: '=SUM(A:A)' },
      ],
      truncated: true,
      incompleteSheets: ['Data'],
    })
    const result = execSync(call('find_cells', { query: 'total' }), fakeDeps({ findCells }))
    expect(result.output).toContain('2 matching cell(s)')
    expect(result.output).toContain('stopped at the cap')
    expect(result.output).toContain('Sheet1!A1: Total')
    expect(result.output).toContain('Summary!B2: 120 (=SUM(A:A))')
    expect(result.output).toContain('indexing has not finished on Data')
  })

  it('reports no matches without an error', () => {
    const result = execSync(call('find_cells', { query: 'missing' }), fakeDeps())
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('No matching cells found')
  })

  it('surfaces truncation even when nothing matched in the scanned region', () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [],
      truncated: true,
      incompleteSheets: [],
    })
    const result = execSync(call('find_cells', { query: 'missing' }), fakeDeps({ findCells }))
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('stopped at the scan budget')
    expect(result.output).toContain('do NOT conclude there are no matches')
  })

  it('propagates search errors (unknown sheet, bad regex)', () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [],
      truncated: false,
      incompleteSheets: [],
      error: 'Invalid regex: bad pattern',
    })
    const result = execSync(
      call('find_cells', { query: '(', regex: true }),
      fakeDeps({ findCells }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Invalid regex')
  })
})

describe('executeWorkbookTool: select_range', () => {
  it('rejects an unparsable range', () => {
    const result = execSync(call('select_range', { range: 'nope!!' }), fakeDeps())
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Cannot parse range')
  })

  it('selects a range and reports the normalized label', () => {
    const selectRange = vi.fn().mockReturnValue({ ok: true, sheetName: 'Summary' })
    const result = execSync(
      call('select_range', { range: '$b$2:c4', sheetId: 'sheet-2' }),
      fakeDeps({ selectRange }),
    )
    expect(selectRange).toHaveBeenCalledWith(
      'sheet-2',
      expect.objectContaining({ startRow: 1, startColumn: 1, endRow: 3, endColumn: 2 }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('Selected Summary!B2:C4')
    expect(result.mutated).toBe(false)
  })

  it('collapses a single cell to one address', () => {
    const result = execSync(call('select_range', { range: 'B2' }), fakeDeps())
    expect(result.output).toContain('Sheet1!B2 ')
    expect(result.output).not.toContain('B2:B2')
  })

  it('propagates selection failures', () => {
    const selectRange = vi.fn().mockReturnValue({ ok: false, error: 'Unknown sheet: ghost' })
    const result = execSync(
      call('select_range', { range: 'A1', sheetId: 'ghost' }),
      fakeDeps({ selectRange }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: ghost')
  })
})

describe('executeWorkbookTool: trace_precedents', () => {
  it('rejects a range instead of a single cell', () => {
    const result = execSync(call('trace_precedents', { address: 'A1:B2' }), fakeDeps())
    expect(result.isError).toBe(true)
    expect(result.output).toContain('single cell')
  })

  it('normalizes $-refs and formats refs with error flags', () => {
    const tracePrecedents = vi.fn().mockReturnValue({
      formula: '=B1/B2',
      value: '#DIV/0!',
      refs: [
        {
          label: 'B1',
          cellCount: 1,
          samples: [{ address: 'B1', value: 10 }],
          hasError: false,
        },
        {
          label: 'B2',
          cellCount: 1,
          samples: [{ address: 'B2', value: 0 }],
          hasError: false,
        },
      ],
      usesNames: false,
    })
    const result = execSync(
      call('trace_precedents', { address: '$c$10' }),
      fakeDeps({ tracePrecedents }),
    )
    expect(tracePrecedents).toHaveBeenCalledWith(undefined, 'C10')
    expect(result.output).toContain('C10 = #DIV/0! (=B1/B2)')
    expect(result.output).toContain('- B1 (1 cell(s)): B1=10')
  })

  it('marks error-bearing and external refs and defined-name usage', () => {
    const tracePrecedents = vi.fn().mockReturnValue({
      formula: '=SUM(Data!A1:A9)+[Ext.xlsx]S1!B2+Total',
      value: '#REF!',
      refs: [
        {
          label: 'Data!A1:A9',
          cellCount: 9,
          samples: [{ address: 'A1', value: '#REF!', formula: '=Gone!A1' }],
          hasError: true,
        },
        { label: '[Ext.xlsx]S1!…', cellCount: 0, samples: [], hasError: false, external: true },
      ],
      usesNames: true,
    })
    const result = execSync(
      call('trace_precedents', { address: 'C1' }),
      fakeDeps({ tracePrecedents }),
    )
    expect(result.output).toContain('⚠️ contains error values')
    expect(result.output).toContain('external/unresolved reference')
    expect(result.output).toContain('defined names')
  })

  it('explains non-formula cells', () => {
    const tracePrecedents = vi.fn().mockReturnValue({ refs: [], value: 42 })
    const result = execSync(
      call('trace_precedents', { address: 'A1' }),
      fakeDeps({ tracePrecedents }),
    )
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('not a formula cell; value: 42')
  })
})

describe('executeWorkbookTool: trace_dependents', () => {
  it('lists dependents with truncation and indexing notes', () => {
    const traceDependents = vi.fn().mockReturnValue({
      dependents: [
        { sheetName: 'Sheet1', address: 'C1', formula: '=B1*2', value: 20 },
        { sheetName: 'Summary', address: 'A1', formula: '=Sheet1!B1', value: 20 },
      ],
      truncated: true,
      incompleteSheets: ['Data'],
    })
    const result = execSync(
      call('trace_dependents', { address: 'b1', sheetId: 'sheet-1' }),
      fakeDeps({ traceDependents }),
    )
    expect(traceDependents).toHaveBeenCalledWith('sheet-1', 'B1')
    expect(result.output).toContain('2+ formula cell(s) read B1')
    expect(result.output).toContain('- Summary!A1 = 20 (=Sheet1!B1)')
    expect(result.output).toContain('result cap')
    expect(result.output).toContain('indexing has not finished on Data')
  })

  it('reports zero dependents with the defined-name caveat', () => {
    const result = execSync(call('trace_dependents', { address: 'Z9' }), fakeDeps())
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('No formulas read Z9')
    expect(result.output).toContain('defined name')
  })

  it('propagates errors', () => {
    const traceDependents = vi.fn().mockReturnValue({
      dependents: [],
      truncated: false,
      incompleteSheets: [],
      error: 'Unknown sheet: ghost',
    })
    const result = execSync(
      call('trace_dependents', { address: 'A1', sheetId: 'ghost' }),
      fakeDeps({ traceDependents }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: ghost')
  })
})

describe('executeWorkbookTool: unknown tool', () => {
  it('fails closed on an unrecognized tool name', () => {
    const result = execSync(call('delete_everything', {}), fakeDeps())
    expect(result.isError).toBe(true)
  })
})

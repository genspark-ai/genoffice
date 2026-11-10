/**
 * MCP worksheet addressing: the tools accept the CLI's NAME vocabulary
 * (`sheet` / `sourceSheet` / `targetSheet`), while the DSL underneath keys
 * sheets by id. Names are what survive a close-and-reopen (ids are minted per
 * session), so the bridge resolves them in the renderer before validation.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { installSheetsMcpBridge, type McpSheetHandlers } from '../src/renderer/mcp-bridge'
import { normalizeSheetRefs, primaryCellOf, primarySheetId } from '../src/renderer/mcp-sheet-refs'

const SHEETS = [
  { id: 'sheet-1', name: 'Summary' },
  { id: 'jg_nIT-Asw-tL3tvR3f4q', name: 'Data 2024' },
]

describe('normalizeSheetRefs', () => {
  it('turns a name into the id the DSL expects', () => {
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: 'Summary', address: 'A1' }], SHEETS)
    expect(result.ok).toBe(true)
    expect(result).toMatchObject({ ops: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1' }] })
  })

  it('matches sheet names case-insensitively and ignores surrounding space', () => {
    const result = normalizeSheetRefs(
      [{ op: 'set_cell', sheet: '  data 2024 ', address: 'B2' }],
      SHEETS,
    )
    expect(result).toMatchObject({
      ops: [{ sheetId: 'jg_nIT-Asw-tL3tvR3f4q', address: 'B2' }],
    })
  })

  it('resolves sourceSheet and targetSheet, including one level down (chart series)', () => {
    const result = normalizeSheetRefs(
      [
        {
          op: 'copy_range',
          sheet: 'Summary',
          sourceSheet: 'Data 2024',
          source: 'A1:B2',
          target: 'C1',
        },
        {
          op: 'add_chart',
          sheet: 'Summary',
          series: [{ valuesRange: 'B2:B9', sheet: 'Data 2024' }],
        },
      ],
      SHEETS,
    )
    expect(result).toMatchObject({
      ops: [
        {
          op: 'copy_range',
          sheetId: 'sheet-1',
          sourceSheetId: 'jg_nIT-Asw-tL3tvR3f4q',
          source: 'A1:B2',
          target: 'C1',
        },
        { op: 'add_chart', sheetId: 'sheet-1', series: [{ sheetId: 'jg_nIT-Asw-tL3tvR3f4q' }] },
      ],
    })
  })

  it('accepts an id sent in the name slot (an agent reusing an earlier read)', () => {
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: 'sheet-1', address: 'A1' }], SHEETS)
    expect(result).toMatchObject({ ops: [{ sheetId: 'sheet-1' }] })
  })

  it('rejects an unknown name before anything applies, listing the real sheets', () => {
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: 'Sheet1', address: 'A1' }], SHEETS)
    expect(result.ok).toBe(false)
    expect(result).toMatchObject({ error: expect.stringContaining('"Sheet1"') })
    if (!result.ok) {
      expect(result.error).toContain('Summary')
      expect(result.error).toContain('Data 2024')
      expect(result.error).toContain('op #0')
    }
  })

  it('leaves an id-addressed batch untouched', () => {
    const ops = [{ op: 'set_formula', sheetId: 'sheet-1', address: 'B1', formula: '=1/3' }]
    const result = normalizeSheetRefs(ops, SHEETS)
    expect(result).toMatchObject({ ops })
  })

  it('reports the sheet list as empty when the workbook has none', () => {
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: 'Nope', address: 'A1' }], [])
    expect(result).toMatchObject({ ok: false, error: expect.stringContaining('none') })
  })

  it('ignores an empty-string ref like an absent one', () => {
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: '', address: 'A1' }], SHEETS)
    expect(result.ok).toBe(true)
    if (!result.ok) throw new Error('expected ok')
    expect(result.ops).toHaveLength(1)
    expect(result.ops[0]).not.toHaveProperty('sheet')
    expect(result.ops[0]).toMatchObject({ op: 'set_cell', address: 'A1' })
  })

  it('strips an empty-string name but keeps an explicit id', () => {
    const result = normalizeSheetRefs(
      [{ op: 'set_cell', sheet: '', sheetId: 'sheet-1', address: 'A1' }],
      SHEETS,
    )
    expect(result).toMatchObject({ ops: [{ sheetId: 'sheet-1', address: 'A1' }] })
    if (result.ok) expect(result.ops[0]).not.toHaveProperty('sheet')
  })

  it('resolves duplicate names deterministically: first tab wins', () => {
    const dupes = [
      { id: 'sheet-1', name: 'Summary' },
      { id: 'sheet-9', name: 'summary' },
    ]
    const result = normalizeSheetRefs([{ op: 'set_cell', sheet: 'SUMMARY', address: 'A1' }], dupes)
    expect(result).toMatchObject({ ops: [{ sheetId: 'sheet-1' }] })
  })
})

describe('primarySheetId / primaryCellOf', () => {
  it('reports the first sheet and cell the batch touches', () => {
    const ops = normalizeSheetRefs(
      [
        { op: 'set_cell', sheet: 'Data 2024', address: 'C7', value: 1 },
        { op: 'set_cell', sheet: 'Summary', address: 'A1', value: 2 },
      ],
      SHEETS,
    )
    if (!ops.ok) throw new Error(ops.error)
    expect(primarySheetId(ops.ops)).toBe('jg_nIT-Asw-tL3tvR3f4q')
    expect(primaryCellOf(ops.ops)).toBe('C7')
  })

  it('takes the top-left of a range', () => {
    expect(primaryCellOf([{ op: 'set_range', sheetId: 'sheet-1', range: 'B2:D9' }])).toBe('B2')
  })

  it('has nothing to report for a sheet-level op', () => {
    expect(primarySheetId([{ op: 'add_sheet', name: 'New' }])).toBeUndefined()
    expect(primaryCellOf([{ op: 'add_sheet', name: 'New' }])).toBeUndefined()
  })

  it('finds the primary sheet inside nested series data', () => {
    expect(
      primarySheetId([
        {
          op: 'edit_chart',
          chartPath: 'added-chart-1',
          seriesData: [{ index: 0, sheetId: 'jg_nIT-Asw-tL3tvR3f4q', valuesRange: 'A1:A5' }],
        },
      ]),
    ).toBe('jg_nIT-Asw-tL3tvR3f4q')
  })

  it('prefers the top-level sheet but falls back to nested series', () => {
    expect(
      primarySheetId([
        {
          op: 'add_chart',
          sheetId: 'sheet-1',
          series: [{ sheetId: 'jg_nIT-Asw-tL3tvR3f4q' }],
        },
      ]),
    ).toBe('sheet-1')
  })

  it('reports the copy target cell for a source-sheet-only op', () => {
    expect(
      primaryCellOf([
        {
          op: 'copy_range',
          sourceSheetId: 'jg_nIT-Asw-tL3tvR3f4q',
          source: 'A1:B2',
          target: 'C3',
        },
      ]),
    ).toBe('C3')
  })

  it('reports the target over the source for copy_range', () => {
    expect(
      primaryCellOf([{ op: 'copy_range', sheetId: 'sheet-1', source: 'A1:B2', target: 'C3:D4' }]),
    ).toBe('C3')
  })

  it('focuses the top of a whole-column range and skips named ranges', () => {
    expect(primaryCellOf([{ op: 'clear_range', sheetId: 'sheet-1', range: 'A:C' }])).toBe('A1')
    expect(
      primaryCellOf([{ op: 'clear_range', sheetId: 'sheet-1', range: 'MyRange' }]),
    ).toBeUndefined()
  })
})

describe('mcp bridge sheet addressing', () => {
  let commandListener:
    | ((message: { requestId: number; command: string; payload: Record<string, unknown> }) => void)
    | undefined
  const reportMcpResult = vi.fn()

  beforeEach(() => {
    commandListener = undefined
    reportMcpResult.mockReset()
    ;(globalThis as unknown as { window: unknown }).window = {
      desktopApi: {
        onMcpCommand: (
          listener: (message: {
            requestId: number
            command: string
            payload: Record<string, unknown>
          }) => void,
        ) => {
          commandListener = listener
          return () => {}
        },
        reportMcpResult,
        signalMcpReady: () => {},
      },
    }
  })

  function handlersWith(overrides: Partial<McpSheetHandlers> = {}): McpSheetHandlers {
    return {
      hasWorkbook: () => true,
      context: () => ({}),
      readCells: () => ({}),
      sheets: () => SHEETS,
      focusSheet: vi.fn(),
      applyOps: vi.fn().mockResolvedValue({ ok: true }),
      saveTo: vi.fn().mockResolvedValue({ ok: true }),
      ...overrides,
    }
  }

  async function send(
    command: string,
    payload: Record<string, unknown>,
    handlers = handlersWith(),
  ): Promise<{
    handlers: McpSheetHandlers
    reply: { ok: boolean; result?: unknown; error?: string }
  }> {
    const uninstall = installSheetsMcpBridge(handlers)
    commandListener!({ requestId: 1, command, payload })
    uninstall()
    await vi.waitFor(() => expect(reportMcpResult).toHaveBeenCalledTimes(1))
    return {
      handlers,
      reply: reportMcpResult.mock.calls[0]![0] as {
        ok: boolean
        result?: unknown
        error?: string
      },
    }
  }

  it('applies a name-addressed batch and focuses the addressed sheet', async () => {
    const applyOpsFn = vi.fn().mockResolvedValue({ ok: true })
    const focusSheet = vi.fn()
    const { reply } = await send(
      'apply_ops',
      {
        ops: [
          { op: 'set_cell', sheet: 'Data 2024', address: 'C7', value: 5 },
          { op: 'set_formula', sheet: 'Summary', address: 'B1', formula: '=1/3' },
        ],
      },
      handlersWith({ applyOps: applyOpsFn, focusSheet }),
    )
    expect(reply.ok, reply.error).toBe(true)
    // the DSL receives ids, in batch order
    expect(applyOpsFn).toHaveBeenCalledWith(
      [
        { op: 'set_cell', sheetId: 'jg_nIT-Asw-tL3tvR3f4q', address: 'C7', value: 5 },
        { op: 'set_formula', sheetId: 'sheet-1', address: 'B1', formula: '=1/3' },
      ],
      false,
    )
    // focus follows the FIRST op, so the user watches the sheet that changes first
    expect(focusSheet).toHaveBeenCalledWith('jg_nIT-Asw-tL3tvR3f4q', 'C7')
  })

  it('does not apply when a sheet name is unknown', async () => {
    const applyOpsFn = vi.fn()
    const { reply } = await send(
      'apply_ops',
      { ops: [{ op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 1 }] },
      handlersWith({ applyOps: applyOpsFn }),
    )
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('Sheet1')
    expect(applyOpsFn).not.toHaveBeenCalled()
  })

  it('reads from a sheet addressed by name', async () => {
    const readCells = vi.fn().mockReturnValue({ A1: { value: 1 } })
    const { reply } = await send(
      'read_sheet',
      { sheet: 'Data 2024', addresses: ['A1'] },
      handlersWith({ readCells }),
    )
    expect(reply.ok, reply.error).toBe(true)
    expect(readCells).toHaveBeenCalledWith(['A1'], 'jg_nIT-Asw-tL3tvR3f4q')
    expect(reply.result).toEqual({ cells: { A1: { value: 1 } } })
  })

  it('refuses a read naming a sheet that is not open', async () => {
    const readCells = vi.fn()
    const { reply } = await send(
      'read_sheet',
      { sheet: 'Gone', addresses: ['A1'] },
      handlersWith({ readCells }),
    )
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('Gone')
    expect(readCells).not.toHaveBeenCalled()
  })

  it('still reads by id when no name is given', async () => {
    const readCells = vi.fn().mockReturnValue({})
    const { reply } = await send(
      'read_sheet',
      { sheetId: 'sheet-1', addresses: ['A1'] },
      handlersWith({ readCells }),
    )
    expect(reply.ok).toBe(true)
    expect(readCells).toHaveBeenCalledWith(['A1'], 'sheet-1')
  })
})

/**
 * The MCP apply_ops boundary validates raw ops against the workbook DSL schema
 * (the same gate the built-in AI path applies). Without it, an op missing its
 * sheetId surfaces as a cryptic "Unknown sheet: undefined" from the planner
 * instead of naming the op and pointing at read_sheet.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { installSheetsMcpBridge, type McpSheetHandlers } from '../src/renderer/mcp-bridge'

type CommandListener = (message: {
  requestId: number
  command: string
  payload: Record<string, unknown>
}) => void

let commandListener: CommandListener | undefined
const reportMcpResult = vi.fn()

beforeEach(() => {
  commandListener = undefined
  reportMcpResult.mockReset()
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: {
      onMcpCommand: (listener: CommandListener) => {
        commandListener = listener
        return () => {}
      },
      reportMcpResult,
      signalMcpReady: () => {},
    },
  }
})

function bridgeWith(overrides: Partial<McpSheetHandlers> = {}): McpSheetHandlers {
  return {
    hasWorkbook: () => true,
    context: () => ({}),
    readCells: () => ({}),
    applyOps: vi.fn().mockResolvedValue({ ok: true }),
    saveTo: vi.fn().mockResolvedValue({ ok: true }),
    ...overrides,
  }
}

async function applyOps(payload: Record<string, unknown>): Promise<void> {
  const uninstall = installSheetsMcpBridge(bridgeWith())
  commandListener!({ requestId: 1, command: 'apply_ops', payload })
  uninstall()
  await vi.waitFor(() => expect(reportMcpResult).toHaveBeenCalledTimes(1))
}

describe('mcp bridge apply_ops validation', () => {
  it('rejects an op missing sheetId with the op named and a read_sheet hint', async () => {
    await applyOps({ ops: [{ op: 'set_cell', address: 'A1', value: 'x' }] })
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean; error?: string }
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('op #0 (set_cell)')
    expect(reply.error).toContain('sheetId')
    expect(reply.error).toContain('read_sheet')
  })

  it('passes schema-valid ops through to the apply handler', async () => {
    const applyOpsFn = vi.fn().mockResolvedValue({ ok: true, applied: true })
    const uninstall = installSheetsMcpBridge(bridgeWith({ applyOps: applyOpsFn }))
    commandListener!({
      requestId: 2,
      command: 'apply_ops',
      payload: { ops: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'x' }] },
    })
    uninstall()
    await vi.waitFor(() => expect(reportMcpResult).toHaveBeenCalledTimes(1))
    expect(applyOpsFn).toHaveBeenCalledTimes(1)
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean }
    expect(reply.ok).toBe(true)
  })
})

import { z } from 'zod'
import type { McpCommandMessage } from '../shared/desktop-api'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/**
 * Renderer half of the MCP → sheets-grid bridge.
 *
 * The shell main process pushes one command at a time (`sheets:mcp-command`);
 * each command runs against the live Univer workbook through the same executors
 * the built-in AI uses (readers from ai/workbook-readers.ts, the op planner and
 * apply path from op-executor.ts, the save pipeline from save-actions.ts), so
 * external edits land in the edit journal and undo history exactly like
 * in-app/AI ones. Results are reported back correlated by requestId
 * (`sheets:mcp-result`), and once the workbook is mounted the bridge announces
 * `sheets:mcp-ready` — the shell waits for it before accepting a session.
 *
 * No-op when the preload lacks the bridge channels (standalone/dev builds
 * without MCP wiring).
 */

const READY_POLL_MS = 50
/** after the fast window the poll backs off instead of stopping (see below) */
const READY_SLOW_POLL_MS = 1_000
const READY_FAST_WINDOW_MS = 20_000

export interface McpSheetHandlers {
  /** workbook mounted and editable? (drives the ready announce) */
  hasWorkbook: () => boolean
  /** workbook overview: sheets with ids/names/extents, active sheet */
  context: () => unknown
  /** values + formulas for the requested A1 addresses */
  readCells: (addresses: string[], sheetId?: string) => unknown
  /** apply a workbook DSL batch (planFromOps + applyChangePlan) */
  applyOps: (ops: WorkbookOperation[], dryRun: boolean) => Promise<unknown>
  /** save to an explicit absolute path (dialog-free) */
  saveTo: (
    path: string,
    overwrite: boolean,
  ) => Promise<{ ok: boolean; path?: string; error?: string }>
}

export function installSheetsMcpBridge(handlers: McpSheetHandlers): () => void {
  const api = window.desktopApi
  if (typeof api?.onMcpCommand !== 'function') return () => {}

  /** one command at a time, in arrival order — applies and saves must not interleave */
  let queue: Promise<void> = Promise.resolve()
  let disposed = false

  // Announce readiness as soon as the workbook is mounted, then stop polling.
  //
  // The poll must never give up permanently: a cold dev start (vite serving the
  // Univer bundle, the sidecar opening the workbook) can take longer than the
  // fast window, and once the timer has stopped the tab stays unaddressable
  // over MCP for the rest of its life — the grid looks perfectly normal in the
  // UI while every sheet command times out in the shell. Back off instead.
  let readyTimer: ReturnType<typeof setTimeout> | null = null
  const mountedAt = performance.now()
  const announceWhenMounted = (): void => {
    readyTimer = null
    if (disposed) return
    if (handlers.hasWorkbook()) {
      api.signalMcpReady()
      return
    }
    const slow = performance.now() - mountedAt > READY_FAST_WINDOW_MS
    readyTimer = setTimeout(announceWhenMounted, slow ? READY_SLOW_POLL_MS : READY_POLL_MS)
  }
  readyTimer = setTimeout(announceWhenMounted, READY_POLL_MS)

  const off = api.onMcpCommand((message: McpCommandMessage) => {
    queue = queue.then(() => runCommand(message)).catch(() => undefined)
  })

  return () => {
    disposed = true
    if (readyTimer !== null) clearTimeout(readyTimer)
    off()
  }

  /** First schema issue, with the offending op named so the client can fix it. */
  function describeOpError(ops: unknown[], error: z.ZodError): string {
    const issue = error.issues[0]
    if (!issue) return 'invalid ops'
    const index = typeof issue.path[0] === 'number' ? issue.path[0] : -1
    const raw = index >= 0 ? ops[index] : undefined
    const opName =
      raw && typeof raw === 'object' && 'op' in raw
        ? String((raw as { op: unknown }).op)
        : 'unknown'
    const field = issue.path.slice(1).join('.')
    const hint = issue.path.includes('sheetId')
      ? ' — call read_sheet first and use a sheetId from its output'
      : ''
    return `op #${index} (${opName}) is invalid${field ? ` (${field})` : ''}: ${issue.message}${hint}`
  }

  async function runCommand(message: McpCommandMessage): Promise<void> {
    const reply = (ok: boolean, result?: unknown, error?: string): void => {
      api.reportMcpResult({
        requestId: message.requestId,
        ok,
        ...(result !== undefined ? { result } : {}),
        ...(error !== undefined ? { error } : {}),
      })
    }
    try {
      const payload = (message.payload ?? {}) as Record<string, unknown>
      switch (message.command) {
        case 'read_sheet': {
          const addresses = Array.isArray(payload.addresses)
            ? payload.addresses.filter((a): a is string => typeof a === 'string')
            : []
          const sheetId = typeof payload.sheetId === 'string' ? payload.sheetId : undefined
          if (addresses.length > 0) {
            reply(true, { cells: handlers.readCells(addresses, sheetId) })
          } else {
            reply(true, { context: handlers.context() })
          }
          return
        }
        case 'apply_ops': {
          const ops = Array.isArray(payload.ops) ? payload.ops : []
          if (ops.length === 0) {
            reply(false, undefined, 'ops must be a non-empty array')
            return
          }
          // Same validation the built-in AI path applies (ai/tools.ts): without
          // it a missing sheetId surfaces later as a cryptic
          // "Unknown sheet: undefined" from the planner instead of naming the op.
          const parsed = z.array(workbookOperationSchema).safeParse(ops)
          if (!parsed.success) {
            reply(false, undefined, describeOpError(ops, parsed.error))
            return
          }
          reply(true, await handlers.applyOps(parsed.data, payload.dryRun === true))
          return
        }
        case 'save_sheet': {
          const path = typeof payload.path === 'string' ? payload.path : ''
          const overwrite = payload.overwrite === true
          if (!path) {
            reply(false, undefined, 'save_sheet needs an absolute path')
            return
          }
          const saved = await handlers.saveTo(path, overwrite)
          // a refused save must surface as an error so the agent keeps its session and can retry
          if (!saved.ok) {
            reply(false, undefined, saved.error ?? 'the spreadsheet could not be saved')
            return
          }
          reply(true, saved)
          return
        }
        default:
          reply(false, undefined, `unknown command: ${String(message.command)}`)
      }
    } catch (error: unknown) {
      reply(false, undefined, error instanceof Error ? error.message : String(error))
    }
  }
}

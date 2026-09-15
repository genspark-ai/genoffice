import { ipcMain, webContents } from 'electron'
import type { SheetsControl } from './tools/sheets-tools'

/**
 * Shell-main half of the MCP → sheets-grid bridge.
 *
 * A sheets workbook lives in the **renderer** (Univer), so unlike slides this
 * needs a request/response channel into the tab — the same plumbing as the docs
 * bridge. This module waits for a tab's renderer to announce itself
 * (`sheets:mcp-ready`, sent once the workbook is mounted), pushes one command
 * at a time over `sheets:mcp-command`, and resolves the matching
 * `sheets:mcp-result`. The renderer half is
 * `apps/sheets/src/renderer/mcp-bridge.ts`, which reuses the built-in AI's own
 * readers, op planner and save pipeline.
 *
 * Lifecycle: `installSheetsBridge()` registers the reply listeners once at
 * boot; readiness is tracked per webContents id and dropped when the tab goes
 * away.
 */

const COMMAND_TIMEOUT_MS = 120_000

interface PendingCommand {
  resolve: (result: unknown) => void
  reject: (error: Error) => void
  timer: NodeJS.Timeout
}

const readyIds = new Set<number>()
const readyWaiters = new Map<number, Array<() => void>>()
const pending = new Map<string, PendingCommand>()
let requestSeq = 0
let installed = false

function markReady(wcId: number): void {
  readyIds.add(wcId)
  const waiters = readyWaiters.get(wcId)
  if (waiters) {
    readyWaiters.delete(wcId)
    for (const resolve of waiters) resolve()
  }
}

function waitForReady(wcId: number): Promise<void> {
  if (readyIds.has(wcId)) return Promise.resolve()
  return new Promise((resolve, reject) => {
    const waiters = readyWaiters.get(wcId) ?? []
    const timer = setTimeout(() => {
      const list = readyWaiters.get(wcId)
      if (list) {
        const at = list.indexOf(onReady)
        if (at >= 0) list.splice(at, 1)
        if (list.length === 0) readyWaiters.delete(wcId)
      }
      reject(new Error('the spreadsheet did not become ready in time'))
    }, 30_000)
    const onReady = (): void => {
      clearTimeout(timer)
      resolve()
    }
    waiters.push(onReady)
    readyWaiters.set(wcId, waiters)
  })
}

/** Register the renderer reply channels. Safe to call more than once. */
export function installSheetsBridge(): void {
  if (installed) return
  installed = true
  ipcMain.on('sheets:mcp-ready', (event) => {
    const wcId = event.sender.id
    const first = !readyIds.has(wcId)
    markReady(wcId)
    if (!first) return
    // webContents ids are never reused, so drop readiness when the tab goes away
    event.sender.once('destroyed', () => {
      readyIds.delete(wcId)
      readyWaiters.delete(wcId)
    })
  })
  ipcMain.on('sheets:mcp-result', (event, result: unknown) => {
    void event
    const payload = result as {
      requestId?: unknown
      ok?: unknown
      result?: unknown
      error?: unknown
    }
    if (!payload || typeof payload.requestId !== 'string') return
    const entry = pending.get(payload.requestId)
    if (!entry) return
    pending.delete(payload.requestId)
    clearTimeout(entry.timer)
    if (payload.ok === true) entry.resolve(payload.result)
    else
      entry.reject(new Error(typeof payload.error === 'string' ? payload.error : 'command failed'))
  })
}

export interface SheetsBridgeDeps {
  /** open a fresh blank sheets tab (a real backing file, like the app's own
   *  "new spreadsheet"); returns its webContents id */
  openBlankTab: () => Promise<number>
}

export function createSheetsControl(deps: SheetsBridgeDeps): SheetsControl {
  const runCommand = async (
    wcId: number,
    command: 'apply_ops' | 'read_sheet' | 'save_sheet',
    payload: unknown,
  ): Promise<unknown> => {
    const wc = webContents.fromId(wcId)
    if (!wc || wc.isDestroyed()) throw new Error('the target spreadsheet is no longer open')
    await waitForReady(wcId)
    const requestId = `mcp-${++requestSeq}`
    const result = new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId)
        reject(new Error(`the spreadsheet command timed out after ${COMMAND_TIMEOUT_MS}ms`))
      }, COMMAND_TIMEOUT_MS)
      pending.set(requestId, { resolve, reject, timer })
    })
    wc.send('sheets:mcp-command', { requestId, command, payload })
    return result
  }

  return {
    openBlankTab: async () => {
      const wcId = await deps.openBlankTab()
      await waitForReady(wcId)
      return wcId
    },
    runCommand,
  }
}

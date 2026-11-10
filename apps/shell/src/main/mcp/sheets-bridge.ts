import { existsSync } from 'node:fs'
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

// Kept under the MCP client's own tool timeout (30s), because that is the
// budget the caller actually sees: waiting longer here makes the client report
// a bare "timed out" and the reason below never reaches the agent.
const READY_TIMEOUT_MS = 24_000
const COMMAND_TIMEOUT_MS = 120_000

interface PendingCommand {
  wcId: number
  resolve: (result: unknown) => void
  reject: (error: Error) => void
  timer: NodeJS.Timeout
}

interface ReadyWaiter {
  resolve: () => void
  reject: (error: Error) => void
}

const readyIds = new Set<number>()
const readyWaiters = new Map<number, ReadyWaiter[]>()
/** tabs whose 'destroyed' hook is attached (once per webContents id) */
const watchedIds = new Set<number>()
const pending = new Map<string, PendingCommand>()
let requestSeq = 0
let installed = false

function takeWaiters(wcId: number): ReadyWaiter[] {
  const waiters = readyWaiters.get(wcId) ?? []
  readyWaiters.delete(wcId)
  return waiters
}

function markReady(wcId: number): void {
  readyIds.add(wcId)
  for (const waiter of takeWaiters(wcId)) waiter.resolve()
}

/** a closed tab can never become ready: fail its waiters instead of letting them time out */
function watchDestroyed(wcId: number): void {
  if (watchedIds.has(wcId)) return
  const wc = webContents.fromId(wcId)
  if (!wc || wc.isDestroyed()) return
  watchedIds.add(wcId)
  wc.once('destroyed', () => {
    watchedIds.delete(wcId)
    readyIds.delete(wcId)
    const closed = new Error('the spreadsheet tab was closed')
    for (const waiter of takeWaiters(wcId)) waiter.reject(closed)
    // in-flight commands can never be answered either
    for (const [requestId, entry] of pending) {
      if (entry.wcId !== wcId) continue
      pending.delete(requestId)
      clearTimeout(entry.timer)
      entry.reject(closed)
    }
  })
}

function waitForReady(wcId: number): Promise<void> {
  if (readyIds.has(wcId)) return Promise.resolve()
  const wc = webContents.fromId(wcId)
  if (!wc || wc.isDestroyed()) return Promise.reject(new Error('the spreadsheet tab was closed'))
  watchDestroyed(wcId)
  return new Promise((resolve, reject) => {
    const waiters = readyWaiters.get(wcId) ?? []
    const timer = setTimeout(() => {
      const list = readyWaiters.get(wcId)
      if (list) {
        const at = list.indexOf(waiter)
        if (at >= 0) list.splice(at, 1)
        if (list.length === 0) readyWaiters.delete(wcId)
      }
      reject(new Error(`the spreadsheet did not become ready within ${READY_TIMEOUT_MS}ms`))
    }, READY_TIMEOUT_MS)
    const waiter: ReadyWaiter = {
      resolve: () => {
        clearTimeout(timer)
        resolve()
      },
      reject: (error) => {
        clearTimeout(timer)
        reject(error)
      },
    }
    waiters.push(waiter)
    readyWaiters.set(wcId, waiters)
  })
}

/** Register the renderer reply channels. Safe to call more than once. */
export function installSheetsBridge(): void {
  if (installed) return
  installed = true
  ipcMain.on('sheets:mcp-ready', (event) => {
    markReady(event.sender.id)
    // webContents ids are never reused, so readiness is dropped when the tab goes away
    watchDestroyed(event.sender.id)
  })
  ipcMain.on('sheets:mcp-result', (event, result: unknown) => {
    const payload = result as {
      requestId?: unknown
      ok?: unknown
      result?: unknown
      error?: unknown
    }
    if (!payload || typeof payload.requestId !== 'string') return
    const entry = pending.get(payload.requestId)
    if (!entry) return
    // requestIds are guessable, so any sheets tab could otherwise answer another
    // tab's command: the reply is only valid from the tab the command targeted
    if (entry.wcId !== event.sender.id) return
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
  /** grant the tab's renderer write access to the resolved save target (the sheets save handler checks it) */
  authorizeSave: (wcId: number, filePath: string) => void
  /**
   * Close a blank tab whose session never became ready and delete the empty
   * workbook file created for it. `create_session` must not leave a tab and an
   * orphaned .xlsx behind when the wait fails, and only the caller that created
   * the file knows its path.
   */
  abandonBlankTab?: (wcId: number) => void
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
    if (command === 'save_sheet') {
      const { path: requested, overwrite } = (payload ?? {}) as {
        path?: unknown
        overwrite?: unknown
      }
      if (typeof requested !== 'string') throw new Error('save_sheet requires a path')
      // same normalisation as the sheets save handler, so the authorized path is the written one
      const target = /\.xlsx$/i.test(requested) ? requested : `${requested}.xlsx`
      if (existsSync(target) && overwrite !== true) {
        throw new Error(`file already exists: ${target} (pass overwrite:true to replace it)`)
      }
      deps.authorizeSave(wcId, target)
    }
    const requestId = `mcp-${++requestSeq}`
    const result = new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId)
        reject(new Error(`the spreadsheet command timed out after ${COMMAND_TIMEOUT_MS}ms`))
      }, COMMAND_TIMEOUT_MS)
      pending.set(requestId, { wcId, resolve, reject, timer })
    })
    wc.send('sheets:mcp-command', { requestId, command, payload })
    return result
  }

  return {
    openBlankTab: async () => {
      const wcId = await deps.openBlankTab()
      try {
        await waitForReady(wcId)
      } catch (error) {
        // the tab and its backing file would otherwise outlive the failed
        // create_session as an orphan the user never asked for
        deps.abandonBlankTab?.(wcId)
        throw error
      }
      return wcId
    },
    runCommand,
  }
}

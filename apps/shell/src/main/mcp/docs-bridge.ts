import { existsSync } from 'node:fs'
import { ipcMain, webContents } from 'electron'
import type { DocsControl, McpEditorCommandName } from './tools/document-tools'

/**
 * Shell-main half of the MCP → docs-editor bridge.
 *
 * The MCP tools want to drive a *visible* docs editor. This module owns the
 * request/response plumbing: it waits for a tab's renderer to announce itself
 * (`docs:mcp-ready`), pushes one command at a time over `docs:mcp-command`, and
 * resolves the matching `docs:mcp-result`. The renderer half is
 * `apps/docs/src/renderer/mcp-bridge.ts`, which runs the commands against the
 * live Tiptap editor using the built-in agent's own executors.
 *
 * Lifecycle: `installDocsBridge()` registers the two reply listeners once at
 * boot; `docsControl()` exposes the handle the MCP tool layer needs. Reopening
 * a tab destroys its webContents, so readiness is tracked per id and stale ids
 * are dropped on lookup.
 */

// a fresh docs tab boots its renderer and fonts; measured 0.6-8s, occasionally
// far longer. Capped under the MCP client's own tool timeout (30s) so the
// caller gets this reason instead of a bare "timed out" from the client.
const READY_TIMEOUT_MS = 24_000
const COMMAND_TIMEOUT_MS = 120_000

interface PendingCommand {
  wcId: number
  resolve: (result: unknown) => void
  reject: (error: Error) => void
  timer: NodeJS.Timeout
}

const readyIds = new Set<number>()
interface ReadyWaiter {
  resolve: () => void
  reject: (error: Error) => void
}
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
    const closed = new Error('the document tab was closed')
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
  if (!wc || wc.isDestroyed()) return Promise.reject(new Error('the document tab was closed'))
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
      reject(new Error(`the document did not become ready within ${READY_TIMEOUT_MS}ms`))
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
export function installDocsBridge(): void {
  if (installed) return
  installed = true
  ipcMain.on('docs:mcp-ready', (event) => {
    const wcId = event.sender.id
    markReady(wcId)
    // webContents ids are never reused, so readiness is dropped when the tab goes away
    watchDestroyed(wcId)
  })
  ipcMain.on('docs:mcp-result', (event, result: unknown) => {
    const payload = result as {
      requestId?: unknown
      ok?: unknown
      result?: unknown
      error?: unknown
    }
    if (!payload || typeof payload.requestId !== 'string') return
    const entry = pending.get(payload.requestId)
    if (!entry) return
    // requestIds are guessable, so any docs tab could otherwise answer another
    // tab's command: the reply is only valid from the tab the command targeted
    if (entry.wcId !== event.sender.id) return
    pending.delete(payload.requestId)
    clearTimeout(entry.timer)
    if (payload.ok === true) entry.resolve(payload.result)
    else
      entry.reject(new Error(typeof payload.error === 'string' ? payload.error : 'command failed'))
  })
}

export interface DocsBridgeDeps {
  /** open a fresh blank docs tab; returns its webContents id */
  openBlankTab: () => number
  /** grant the tab's renderer write access to the resolved save target (docs:save-to checks it) */
  authorizeSave: (wcId: number, filePath: string) => void
  /** close a blank tab whose session never became ready, so a failed create_session leaves no orphan */
  abandonBlankTab?: (wcId: number) => void
}

export function createDocsControl(deps: DocsBridgeDeps): DocsControl {
  const runCommand = async (
    wcId: number,
    command: McpEditorCommandName,
    payload: unknown,
  ): Promise<unknown> => {
    const wc = webContents.fromId(wcId)
    if (!wc || wc.isDestroyed()) throw new Error('the target document is no longer open')
    await waitForReady(wcId)
    if (command === 'save_document') {
      const { path: target, overwrite } = (payload ?? {}) as { path?: unknown; overwrite?: unknown }
      if (typeof target !== 'string') throw new Error('save_document requires a path')
      // checked here so the agent gets the reason; the renderer's save only reports success/failure
      if (existsSync(target) && overwrite !== true) {
        throw new Error(`file already exists: ${target} (pass overwrite:true to replace it)`)
      }
      deps.authorizeSave(wcId, target)
    }
    const requestId = `mcp-${++requestSeq}`
    const result = new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId)
        reject(new Error(`the document command timed out after ${COMMAND_TIMEOUT_MS}ms`))
      }, COMMAND_TIMEOUT_MS)
      pending.set(requestId, { wcId, resolve, reject, timer })
    })
    wc.send('docs:mcp-command', { requestId, command, payload })
    return result
  }

  return {
    // resolve only once the tab accepts commands, so the agent's first content
    // call never races the renderer boot
    openBlankTab: async () => {
      const wcId = deps.openBlankTab()
      try {
        await waitForReady(wcId)
      } catch (error) {
        deps.abandonBlankTab?.(wcId)
        throw error
      }
      return wcId
    },
    runCommand,
  }
}

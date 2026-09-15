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

const READY_TIMEOUT_MS = 20_000
const COMMAND_TIMEOUT_MS = 120_000

interface PendingCommand {
  /** the tab this command was sent to; only that tab may answer it */
  wcId: number
  resolve: (result: unknown) => void
  reject: (error: Error) => void
  timer: NodeJS.Timeout
}

const readyIds = new Set<number>()
const readyWaiters = new Map<number, Array<() => void>>()
const pending = new Map<string, PendingCommand>()
let requestSeq = 0
let installed = false

/** drop every command still waiting on this tab, so a closed tab fails fast */
function abortPendingFor(wcId: number): void {
  for (const [requestId, entry] of pending) {
    if (entry.wcId !== wcId) continue
    pending.delete(requestId)
    clearTimeout(entry.timer)
    entry.reject(new Error('the target document was closed while the command was running'))
  }
}

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
      reject(new Error(`the document did not become ready within ${READY_TIMEOUT_MS}ms`))
    }, READY_TIMEOUT_MS)
    const onReady = (): void => {
      clearTimeout(timer)
      resolve()
    }
    waiters.push(onReady)
    readyWaiters.set(wcId, waiters)
  })
}

/** Register the renderer reply channels. Safe to call more than once. */
export function installDocsBridge(): void {
  if (installed) return
  installed = true
  ipcMain.on('docs:mcp-ready', (event) => {
    const wcId = event.sender.id
    const first = !readyIds.has(wcId)
    markReady(wcId)
    if (!first) return
    // webContents ids are never reused, so drop readiness when the tab goes away
    event.sender.once('destroyed', () => {
      readyIds.delete(wcId)
      readyWaiters.delete(wcId)
      abortPendingFor(wcId)
    })
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
    openBlankTab: async () => deps.openBlankTab(),
    runCommand,
  }
}

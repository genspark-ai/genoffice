// Modified for ForaOffice from GenOffice v0.12.852.
import { existsSync } from 'node:fs'
import { ipcMain, webContents } from 'electron'
import type { DocsControl, McpEditorCommandName } from './tools/document-tools'
import { McpToolExecutionError, type McpInvocation } from './mcp-server'

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
  callerId?: string
  resolve: (result: unknown) => void
  reject: (error: Error) => void
  cleanup: () => void
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
const callerTargets = new Map<string, Set<number>>()
let requestSeq = 0
let installed = false

function takePending(requestId: string): PendingCommand | undefined {
  const entry = pending.get(requestId)
  if (!entry) return
  pending.delete(requestId)
  entry.cleanup()
  return entry
}

function cancelRenderer(wcId: number, requestId: string, callerId?: string): void {
  const wc = webContents.fromId(wcId)
  if (wc && !wc.isDestroyed())
    wc.send('docs:mcp-command', {
      requestId: `mcp-cancel-${++requestSeq}`,
      callerId,
      command: 'cancel_operation',
      payload: { requestId },
    })
}

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
    for (const [callerId, targets] of callerTargets) {
      targets.delete(wcId)
      if (!targets.size) callerTargets.delete(callerId)
    }
    const closed = new Error('the document tab was closed')
    for (const waiter of takeWaiters(wcId)) waiter.reject(closed)
    // in-flight commands can never be answered either
    for (const [requestId, entry] of pending) {
      if (entry.wcId !== wcId) continue
      takePending(requestId)
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
    takePending(payload.requestId)
    if (payload.ok === true) entry.resolve(payload.result)
    else {
      const message = typeof payload.error === 'string' ? payload.error : 'command failed'
      entry.reject(
        payload.result && typeof payload.result === 'object' && !Array.isArray(payload.result)
          ? new McpToolExecutionError(message, payload.result as Record<string, unknown>)
          : new Error(message),
      )
    }
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
    invocation?: McpInvocation,
  ): Promise<unknown> => {
    const wc = webContents.fromId(wcId)
    if (!wc || wc.isDestroyed()) throw new Error('the target document is no longer open')
    await waitForReady(wcId)
    if (wc.isDestroyed()) throw new Error('The document tab was closed')
    if (invocation?.signal.aborted) throw new Error('The document operation was canceled')
    if (invocation && command !== 'list_tools') {
      const targets = callerTargets.get(invocation.callerId) ?? new Set<number>()
      targets.add(wcId)
      callerTargets.set(invocation.callerId, targets)
    }
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
        takePending(requestId)
        cancelRenderer(wcId, requestId, invocation?.callerId)
        reject(new Error(`the document command timed out after ${COMMAND_TIMEOUT_MS}ms`))
      }, COMMAND_TIMEOUT_MS)
      const onAbort = () => {
        const entry = takePending(requestId)
        if (!entry) return
        cancelRenderer(wcId, requestId, invocation?.callerId)
        reject(new Error('The document operation was canceled; inspect its state before retrying'))
      }
      pending.set(requestId, {
        wcId,
        callerId: invocation?.callerId,
        resolve,
        reject,
        cleanup: () => {
          clearTimeout(timer)
          invocation?.signal.removeEventListener('abort', onAbort)
        },
      })
      invocation?.signal.addEventListener('abort', onAbort, { once: true })
    })
    wc.send('docs:mcp-command', {
      requestId,
      callerId: invocation?.callerId,
      bindingId: invocation?.bindingId,
      command,
      payload,
    })
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
    closeCaller(callerId) {
      for (const wcId of callerTargets.get(callerId) ?? []) {
        const wc = webContents.fromId(wcId)
        if (wc && !wc.isDestroyed())
          wc.send('docs:mcp-command', {
            requestId: `mcp-close-${++requestSeq}`,
            callerId,
            command: 'close_caller',
            payload: {},
          })
      }
      callerTargets.delete(callerId)
      for (const [requestId, entry] of pending) {
        if (entry.callerId !== callerId) continue
        takePending(requestId)
        entry.reject(new Error('The MCP caller disconnected'))
      }
    },
  }
}

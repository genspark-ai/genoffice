/// <reference lib="webworker" />
/**
 * Sandbox for user scripts (issue #815).
 *
 * Runs inside a Web Worker: no `window`, no DOM, no Electron preload, no file
 * system. Everything the script can touch is either a local value, the API below,
 * or a handle the host resolves against the Univer facade. That is the whole point
 * — a script must not be able to open, save or delete the user's files.
 *
 * Apps Script is synchronous; this API is not, because every call is a message to
 * the host. Scripts therefore `await` (documented in the editor's help text).
 */
import { isHandleRef, type HandleRef, type HostMessage, type WorkerMessage } from './script-rpc'

// Defense in depth: the worker response carries its own strict CSP (a dedicated
// worker never sees the page's CSP, so serving the worker from its own protocol
// is the real gate), but shadow the obvious network surfaces too. import() cannot
// be shadowed — which is exactly why the CSP, not this list, is the boundary.
for (const name of [
  'fetch',
  'XMLHttpRequest',
  'WebSocket',
  'EventSource',
  'importScripts',
  'Worker',
  'SharedWorker',
]) {
  try {
    Reflect.deleteProperty(self, name)
  } catch {
    /* non-configurable global: the CSP still denies the traffic */
  }
}

const pending = new Map<number, (message: HostMessage) => void>()
let nextRequestId = 1

function post(message: WorkerMessage): void {
  ;(self as unknown as Worker).postMessage(message)
}

function ask(target: number | 'app', method: string, args: unknown[]): Promise<unknown> {
  const id = nextRequestId++
  return new Promise((resolve, reject) => {
    pending.set(id, (message) => {
      pending.delete(id)
      if (message.kind === 'result') resolve(message.value)
      else if (message.kind === 'fail') reject(new Error(message.message))
    })
    post({ kind: 'rpc', id, target, method, args })
  })
}

function wrap(value: unknown): unknown {
  // __h === 0 marks an absent object (e.g. getSheetByName('missing')): the script
  // sees null instead of a handle whose every call would throw.
  if (isHandleRef(value) && value.__h === 0) return null
  if (typeof value !== 'object' || value === null) return value
  switch ((value as HandleRef).__k) {
    case 'workbook':
      return asSpreadsheet(value as HandleRef)
    case 'sheet':
      return asSheet(value as HandleRef)
    case 'range':
      return asRange(value as HandleRef)
    default:
      return value
  }
}

const asSpreadsheet = (h: HandleRef) => ({
  getName: () => ask(h.__h, 'getName', []) as Promise<string>,
  getActiveSheet: () => ask(h.__h, 'getActiveSheet', []).then(wrap),
  getSheetByName: (name: string) => ask(h.__h, 'getSheetByName', [name]).then(wrap),
  getSheets: () => ask(h.__h, 'getSheets', []).then((v) => (v as unknown[]).map(wrap)),
  getSheetNames: () => ask(h.__h, 'getSheetNames', []) as Promise<string[]>,
})

const asSheet = (h: HandleRef) => ({
  getName: () => ask(h.__h, 'getName', []) as Promise<string>,
  getLastRow: () => ask(h.__h, 'getLastRow', []) as Promise<number>,
  getLastColumn: () => ask(h.__h, 'getLastColumn', []) as Promise<number>,
  getMaxRows: () => ask(h.__h, 'getMaxRows', []) as Promise<number>,
  getMaxColumns: () => ask(h.__h, 'getMaxColumns', []) as Promise<number>,
  getRange: (a1OrRow: string | number, column?: number, numRows?: number, numColumns?: number) =>
    ask(
      h.__h,
      'getRange',
      [a1OrRow, column, numRows, numColumns].filter((v) => v !== undefined),
    ).then(wrap),
  getDataRange: () => ask(h.__h, 'getDataRange', []).then(wrap),
})

const asRange = (h: HandleRef) => ({
  getA1Notation: () => ask(h.__h, 'getA1Notation', []) as Promise<string>,
  getRow: () => ask(h.__h, 'getRow', []) as Promise<number>,
  getColumn: () => ask(h.__h, 'getColumn', []) as Promise<number>,
  getNumRows: () => ask(h.__h, 'getNumRows', []) as Promise<number>,
  getNumColumns: () => ask(h.__h, 'getNumColumns', []) as Promise<number>,
  getValue: () => ask(h.__h, 'getValue', []),
  getValues: () => ask(h.__h, 'getValues', []),
  setValue: (value: string | number | boolean | null) => ask(h.__h, 'setValue', [value]),
  setValues: (values: unknown[][]) => ask(h.__h, 'setValues', [values]),
  clear: () => ask(h.__h, 'clear', []),
})

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ask('app', 'getActiveSpreadsheet', []).then(wrap),
  getActiveSheet: () => ask('app', 'getActiveSheet', []).then(wrap),
}

function format(value: unknown): string {
  if (typeof value === 'string') return value
  if (value instanceof Error) return value.message
  try {
    return JSON.stringify(value)
  } catch {
    return String(value)
  }
}

const Logger = {
  log: (...parts: unknown[]) => post({ kind: 'log', text: parts.map(format).join(' ') }),
}

const Utilities = {
  sleep: (ms: number) =>
    new Promise<void>((resolve) => setTimeout(resolve, Math.max(0, Math.min(ms, 10_000)))),
}

async function run(code: string): Promise<void> {
  try {
    // Async IIFE so scripts can await; the API objects are the only bindings.
    const fn = new Function(
      'SpreadsheetApp',
      'Logger',
      'Utilities',
      `"use strict"; return (async () => {\n${code}\n})()`,
    ) as (
      app: typeof SpreadsheetApp,
      logger: typeof Logger,
      utilities: typeof Utilities,
    ) => Promise<unknown>
    await fn(SpreadsheetApp, Logger, Utilities)
    post({ kind: 'done' })
  } catch (err) {
    post({
      kind: 'error',
      message: err instanceof Error ? (err.stack ?? err.message) : String(err),
    })
  }
}

self.onmessage = (event: MessageEvent<HostMessage>) => {
  const message = event.data
  if (message.kind === 'run') {
    void run(message.code)
    return
  }
  if (message.kind === 'result' || message.kind === 'fail') {
    pending.get(message.id)?.(message)
  }
}

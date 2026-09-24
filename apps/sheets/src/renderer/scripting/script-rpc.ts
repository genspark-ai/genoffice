/**
 * Scripting RPC for GenOffice Sheets (issue #815).
 *
 * Google Apps Script exposes a synchronous API to user code; a script that can
 * reach `window` (and therefore Electron's preload bridge) would be able to
 * open, save or delete files, so scripts here run inside a Web Worker that gets
 * nothing but the API below. The worker asks the host (the renderer, which owns
 * the Univer facade) to do the work through this tiny request/response protocol.
 *
 * Objects the script holds (a spreadsheet, a sheet, a range) are represented as
 * numeric handles: the host keeps the real facade objects, the worker holds ids.
 */

export const SCRIPT_RPC_TIMEOUT_MS = 30_000

/** Methods the host accepts, per handle kind. `app` is the entry point. */
export type HandleKind = 'app' | 'workbook' | 'sheet' | 'range'

export interface HostRequest {
  kind: 'rpc'
  id: number
  /** numeric handle, or the string 'app' for the entry object */
  target: number | 'app'
  method: string
  args: unknown[]
}

export type WorkerMessage =
  | HostRequest
  | { kind: 'log'; text: string }
  | { kind: 'done' }
  | { kind: 'error'; message: string }

export type HostMessage =
  | { kind: 'run'; code: string }
  | { kind: 'result'; id: number; value: unknown }
  | { kind: 'fail'; id: number; message: string }

/** A handle returned to the script: the number is opaque, the kind helps the worker pick its wrapper. */
export interface HandleRef {
  __h: number
  __k: HandleKind
}

export function isHandleRef(value: unknown): value is HandleRef {
  return typeof value === 'object' && value !== null && typeof (value as HandleRef).__h === 'number'
}

/**
 * Resolve one RPC call against the Univer facade.
 *
 * The facade is duck-typed: Univer's surface differs between versions and we
 * must never let a missing method kill the host, so every call is checked and
 * reported as a script-visible error instead.
 */
export function dispatchRpc(
  host: ScriptHost,
  request: HostRequest,
): { ok: true; value: unknown } | { ok: false; message: string } {
  try {
    return { ok: true, value: host.call(request.target, request.method, request.args) }
  } catch (err) {
    return { ok: false, message: err instanceof Error ? err.message : String(err) }
  }
}

export interface ScriptHost {
  call(target: number | 'app', method: string, args: unknown[]): unknown
}

/** Method signature check used by the host implementation. */
export function pickFunction(
  value: unknown,
  method: string,
  on: string,
): ((...args: unknown[]) => unknown) | null {
  if (!value) return null
  const fn = (value as Record<string, unknown>)[method]
  return typeof fn === 'function' ? (fn as (...args: unknown[]) => unknown).bind(value) : null
}

export function missingMethod(on: string, method: string): never {
  throw new Error(`${on}.${method}() is not available in this build`)
}

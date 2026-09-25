/**
 * Host side of the scripting API (issue #815): answers the worker's RPC calls by
 * driving the Univer facade the renderer already owns.
 *
 * The facade is duck-typed on purpose — Univer's surface varies by version, and a
 * missing method must surface as an error inside the script, never as a crash in
 * the host. Values crossing to the worker are reduced to primitives so nothing
 * that cannot be structured-cloned (rich text, cell objects) leaks across.
 */
import {
  missingMethod,
  pickFunction,
  type HandleKind,
  type HandleRef,
  type ScriptHost,
} from './script-rpc'

export type CellPrimitive = string | number | boolean | null

/** Reduce a facade cell value to a primitive: `''` for an empty cell, like Apps Script. */
export function toCellValue(value: unknown): CellPrimitive {
  if (value === null || value === undefined) return null
  if (typeof value === 'number' || typeof value === 'boolean') return value
  if (typeof value === 'string') return value
  if (Array.isArray(value)) return toCellValue(value[0])
  if (typeof value === 'object') {
    const rec = value as Record<string, unknown>
    if ('v' in rec) return toCellValue(rec.v)
    if ('p' in rec) {
      const dataStream = (rec.p as { dataStream?: unknown } | undefined)?.dataStream
      if (typeof dataStream === 'string') return dataStream.replace(/\r\n$/, '')
      return null
    }
  }
  return null
}

function toValueGrid(grid: unknown): CellPrimitive[][] {
  if (!Array.isArray(grid)) throw new Error('getValues() did not return a grid')
  return (grid as unknown[]).map((row) => {
    if (!Array.isArray(row)) throw new Error('getValues() row is not an array')
    return (row as unknown[]).map(toCellValue)
  })
}

function toInputGrid(grid: unknown): unknown[][] {
  if (!Array.isArray(grid)) throw new Error('setValues() expects a 2D array')
  if (grid.length > 10_000) throw new Error('setValues() grid too large (max 10,000 rows)')
  return (grid as unknown[]).map((row) => {
    if (!Array.isArray(row)) throw new Error('setValues() expects a 2D array')
    return (row as unknown[]).map((cell) => {
      if (cell === null || cell === undefined) return null
      if (typeof cell === 'number' || typeof cell === 'boolean' || typeof cell === 'string')
        return cell
      if (typeof cell === 'object') return cell
      throw new Error('setValues() cells must be string, number, boolean or null')
    })
  })
}

export class FacadeScriptHost implements ScriptHost {
  private readonly handles = new Map<number, { kind: HandleKind; value: unknown }>()
  private nextHandle = 1

  constructor(private readonly api: unknown) {}

  call(target: number | 'app', method: string, args: unknown[]): unknown {
    if (target === 'app') return this.callApp(method, args)
    const entry = this.handles.get(target)
    if (!entry) throw new Error('This object is no longer valid (the script was stopped?)')
    switch (entry.kind) {
      case 'workbook':
        return this.callWorkbook(entry.value, method, args)
      case 'sheet':
        return this.callSheet(entry.value, method, args)
      case 'range':
        return this.callRange(entry.value, method, args)
      default:
        throw new Error(`Unknown handle kind: ${entry.kind}`)
    }
  }

  private ref(kind: HandleKind, value: unknown): HandleRef {
    if (value === null || value === undefined) return { __h: 0, __k: kind }
    const id = this.nextHandle++
    this.handles.set(id, { kind, value })
    return { __h: id, __k: kind }
  }

  private callApp(method: string, args: unknown[]): unknown {
    if (method === 'getActiveSpreadsheet') {
      const fn = pickFunction(this.api, 'getActiveWorkbook')
      if (!fn) missingMethod('SpreadsheetApp', 'getActiveSpreadsheet')
      return this.ref('workbook', fn!(...args))
    }
    if (method === 'getActiveSheet') {
      const workbook = pickFunction(this.api, 'getActiveWorkbook')?.()
      if (!workbook) return { __h: 0, __k: 'sheet' }
      const sheet = pickFunction(workbook, 'getActiveSheet')?.()
      return this.ref('sheet', sheet)
    }
    missingMethod('SpreadsheetApp', method)
  }

  private callWorkbook(workbook: unknown, method: string, args: unknown[]): unknown {
    const fn = pickFunction(workbook, method)
    if (!fn) missingMethod('Spreadsheet', method)
    const value = fn!(...args)
    switch (method) {
      case 'getActiveSheet':
      case 'getSheetByName':
      case 'getSheetBySheetId':
        return this.ref('sheet', value)
      case 'getSheets':
        return (Array.isArray(value) ? value : []).map((sheet) => this.ref('sheet', sheet))
      case 'getSheetNames':
        return Array.isArray(value) ? value.map(String) : []
      default:
        return typeof value === 'object' ? String(value) : (value ?? null)
    }
  }

  private callSheet(sheet: unknown, method: string, args: unknown[]): unknown {
    const fn = pickFunction(sheet, method)
    if (!fn) missingMethod('Sheet', method)
    const value = fn!(...args)
    switch (method) {
      case 'getRange':
      case 'getDataRange':
        return this.ref('range', value)
      default:
        return typeof value === 'object' ? String(value) : (value ?? null)
    }
  }

  private callRange(range: unknown, method: string, args: unknown[]): unknown {
    const fn = pickFunction(range, method)
    if (!fn) missingMethod('Range', method)
    switch (method) {
      case 'setValues':
        return !!fn!(toInputGrid(args[0]))
      case 'setValue': {
        const cell = args[0]
        if (!(cell === null || ['string', 'number', 'boolean'].includes(typeof cell)))
          throw new Error('setValue() expects a string, number, boolean or null')
        return !!fn!(cell)
      }
      case 'getValues':
        return toValueGrid(fn!())
      case 'getValue':
        return toCellValue(fn!())
      default: {
        const value = fn!(...args)
        return typeof value === 'object' ? String(value) : (value ?? null)
      }
    }
  }
}

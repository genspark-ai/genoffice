// Scripting host (issue #815): the facade is duck-typed, handles stand in for
// workbook/sheet/range objects, and values are reduced to primitives before they
// cross to the worker.
import { describe, expect, it } from 'vitest'
import { FacadeScriptHost, toCellValue } from '../src/renderer/scripting/script-host'
import type { HandleRef } from '../src/renderer/scripting/script-rpc'

const range = {
  getA1Notation: () => 'A1:B2',
  getValue: () => ({ v: 'hello' }),
  getValues: () => [
    [{ v: 'a' }, { v: 1 }],
    [{ v: true }, undefined],
  ],
  setValue: () => true,
  setValues: () => true,
}

const sheet = {
  getName: () => 'Sheet1',
  getLastRow: () => 3,
  getLastColumn: () => 2,
  getRange: (..._args: unknown[]) => range,
  getDataRange: () => range,
}

const api = {
  getActiveWorkbook: () => ({
    getName: () => 'Book1',
    getActiveSheet: () => sheet,
    getSheetByName: (name: string) => (name === 'Sheet1' ? sheet : null),
    getSheets: () => [sheet],
  }),
}

describe('toCellValue', () => {
  it('unwraps Univer cell objects and rich text', () => {
    expect(toCellValue({ v: 'text' })).toBe('text')
    expect(toCellValue({ p: { dataStream: 'line\r\n' } })).toBe('line')
    expect(toCellValue(undefined)).toBeNull()
    expect(toCellValue(7)).toBe(7)
  })
})

describe('FacadeScriptHost', () => {
  it('walks app → workbook → sheet → range through handles', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    expect(workbook.__k).toBe('workbook')

    const active = host.call(workbook.__h, 'getActiveSheet', []) as HandleRef
    expect(active.__k).toBe('sheet')

    const grid = host.call(active.__h, 'getRange', ['A1:B2']) as HandleRef
    expect(grid.__k).toBe('range')
    expect(host.call(grid.__h, 'getValues', [])).toEqual([
      ['a', 1],
      [true, null],
    ])
  })

  it('returns a sheet handle for getSheetByName and a list for getSheets', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    const found = host.call(workbook.__h, 'getSheetByName', ['Sheet1']) as HandleRef
    expect(found.__k).toBe('sheet')
    expect(host.call(found.__h, 'getName', [])).toBe('Sheet1')

    const sheets = host.call(workbook.__h, 'getSheets', []) as HandleRef[]
    expect(sheets).toHaveLength(1)
    expect(host.call(sheets[0]!.__h, 'getLastRow', [])).toBe(3)
  })

  it('reports a missing facade method as a script error instead of crashing', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    expect(() => host.call(workbook.__h, 'notAThing', [])).toThrow(/not available/)
  })

  it('enforces the method allowlist host-side, per handle kind', () => {
    const host = new FacadeScriptHost(api)
    // 'dispose' exists on real Univer objects — the allowlist, not the facade, is
    // what stops a crafted rpc from reaching it.
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    expect(() => host.call('app', 'dispose', [])).toThrow(/not available/)
    expect(() => host.call(workbook.__h, 'dispose', [])).toThrow(/not available/)
    const active = host.call(workbook.__h, 'getActiveSheet', []) as HandleRef
    expect(() => host.call(active.__h, 'setRowHeight', [100])).toThrow(/not available/)
  })

  it('getSheetByName on a missing sheet yields an absent handle (0), not a broken one', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    const missing = host.call(workbook.__h, 'getSheetByName', ['nope']) as HandleRef
    expect(missing.__h).toBe(0)
  })

  it('reset() invalidates every handle from the previous run', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    host.reset()
    expect(() => host.call(workbook.__h, 'getName', [])).toThrow(/no longer valid/)
  })

  it('rejects input that is not a 2D array', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    const active = host.call(workbook.__h, 'getActiveSheet', []) as HandleRef
    const grid = host.call(active.__h, 'getDataRange', []) as HandleRef
    expect(() => host.call(grid.__h, 'setValues', ['nope'])).toThrow(/2D array/)
  })

  it('does not hand the facade object itself to the script', () => {
    const host = new FacadeScriptHost(api)
    const workbook = host.call('app', 'getActiveSpreadsheet', []) as HandleRef
    // a handle is a plain number + kind: nothing structured-cloneable reaches the worker
    expect(Object.keys(workbook).sort()).toEqual(['__h', '__k'])
  })
})

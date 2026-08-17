import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { installCachedValueFallbackInterceptor } from '../src/renderer/formula-cached-fallback'

interface Interceptor {
  priority: number
  handler: (
    cell: Record<string, unknown> | undefined,
    location: { subUnitId: string; row: number; col: number },
    next: (cell: Record<string, unknown> | undefined) => unknown,
  ) => unknown
}

function captureInterceptor(lazyWorkbookRef: { current: unknown }): Interceptor {
  let captured: Interceptor | undefined
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: () => ({
          intercept: (_point: unknown, interceptor: Interceptor) => {
            captured = interceptor
            return { dispose: () => undefined }
          },
        }),
      }),
    },
  }
  installCachedValueFallbackInterceptor(runtime as never, lazyWorkbookRef as { current: null })
  if (!captured) throw new Error('interceptor not registered')
  return captured
}

function makeState(cached: Record<string, string | number | boolean>) {
  return {
    editJournal: { structuralOps: new Map(), cells: new Map() },
    cachedFormulaValues: new Map([['sheet-1', new Map(Object.entries(cached))]]),
  }
}

const at = (row: number, col: number) => ({ subUnitId: 'sheet-1', row, col })
const passthrough = (cell: Record<string, unknown> | undefined) => cell

describe('installCachedValueFallbackInterceptor', () => {
  it('replaces an engine error with the file-cached value', () => {
    const ref = { current: makeState({ '0:0': 51.6373 }) }
    const { handler } = captureInterceptor(ref)
    expect(handler({ v: '#NAME?', f: '=X[Y]' }, at(0, 0), passthrough)).toEqual({
      v: 51.6373,
      f: '=X[Y]',
      t: CellValueType.NUMBER,
    })
  })

  it('types cached strings and booleans', () => {
    const ref = { current: makeState({ '0:0': 'All States', '0:1': true }) }
    const { handler } = captureInterceptor(ref)
    expect(handler({ v: '#VALUE!' }, at(0, 0), passthrough)).toMatchObject({
      v: 'All States',
      t: CellValueType.STRING,
    })
    expect(handler({ v: '#N/A' }, at(0, 1), passthrough)).toMatchObject({
      v: true,
      t: CellValueType.BOOLEAN,
    })
  })

  it('leaves non-error values, uncached cells, and identical errors alone', () => {
    const ref = { current: makeState({ '0:0': 42, '1:0': '#DIV/0!' }) }
    const { handler } = captureInterceptor(ref)
    const value = { v: 7, f: '=A1' }
    expect(handler(value, at(0, 0), passthrough)).toBe(value)
    const uncached = { v: '#NAME?' }
    expect(handler(uncached, at(5, 5), passthrough)).toBe(uncached)
    // The file itself cached this error; nothing better to show.
    const sameError = { v: '#DIV/0!' }
    expect(handler(sameError, at(1, 0), passthrough)).toBe(sameError)
  })

  it('defers to the engine once the user owns the cell or shifted the sheet', () => {
    const edited = makeState({ '0:0': 42 })
    edited.editJournal.cells.set('sheet-1', new Map([['0:0', { hasValue: true }]]))
    const { handler } = captureInterceptor({ current: edited })
    const cell = { v: '#NAME?' }
    expect(handler(cell, at(0, 0), passthrough)).toBe(cell)

    const shifted = makeState({ '0:0': 42 })
    shifted.editJournal.structuralOps.set('sheet-1', [{ kind: 'remove-rows' }])
    const { handler: shiftedHandler } = captureInterceptor({ current: shifted })
    expect(shiftedHandler(cell, at(0, 0), passthrough)).toBe(cell)
  })

  it('survives layout-only structural ops (sizing, hiding)', () => {
    const state = makeState({ '0:0': 42 })
    state.editJournal.structuralOps.set('sheet-1', [
      { kind: 'set-row-size' },
      { kind: 'set-cols-hidden' },
    ])
    const { handler } = captureInterceptor({ current: state })
    expect(handler({ v: '#NAME?' }, at(0, 0), passthrough)).toMatchObject({ v: 42 })
  })

  it('keeps computed errors on an edited sheet but still falls back for #NAME?', () => {
    const state = makeState({ '0:0': 42, '0:1': 7 })
    state.editJournal.cells.set('sheet-1', new Map([['9:9', { hasValue: true }]]))
    const { handler } = captureInterceptor({ current: state })
    // #DIV/0! may be the true result of the user's new inputs.
    const divide = { v: '#DIV/0!' }
    expect(handler(divide, at(0, 0), passthrough)).toBe(divide)
    // #NAME? cannot be computed under any inputs; the cache stays better.
    expect(handler({ v: '#NAME?' }, at(0, 1), passthrough)).toMatchObject({ v: 7 })
  })

  it('ignores style-only journal entries', () => {
    const state = makeState({ '0:0': 42, '0:1': 7 })
    // Bold/number-format writes journal with hasValue: false.
    state.editJournal.cells.set(
      'sheet-1',
      new Map([
        ['9:9', { hasValue: false }],
        ['0:1', { hasValue: false }],
      ]),
    )
    const { handler } = captureInterceptor({ current: state })
    expect(handler({ v: '#DIV/0!' }, at(0, 0), passthrough)).toMatchObject({ v: 42 })
    expect(handler({ v: '#VALUE!' }, at(0, 1), passthrough)).toMatchObject({ v: 7 })
  })
})

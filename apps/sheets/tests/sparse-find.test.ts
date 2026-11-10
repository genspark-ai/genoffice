import { describe, expect, it, vi } from 'vitest'
import type { IRange } from '@univerjs/core'
import { FindDirection, type IFindQuery } from '@univerjs/find-replace'

import {
  collectRuns,
  installSparseFind,
  patchFindModel,
  sliceRange,
} from '../src/renderer/sparse-find'

const range = (
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): IRange => ({
  startRow,
  startColumn,
  endRow,
  endColumn,
})

describe('collectRuns', () => {
  const matrix = {
    0: { 0: { v: 'a' }, 10: { v: 'k' } },
    2: { 2: { v: 'c' } },
    5: undefined,
    7: { 3: null },
  }

  it('spans each row over its populated columns only', () => {
    expect(collectRuns(matrix, [], false)).toEqual([
      { index: 0, start: 0, end: 10 },
      { index: 2, start: 2, end: 2 },
    ])
  })

  it('spans columns over their populated rows in column direction', () => {
    expect(collectRuns(matrix, [], true)).toEqual([
      { index: 0, start: 0, end: 0 },
      { index: 2, start: 2, end: 2 },
      { index: 10, start: 0, end: 0 },
    ])
  })

  it('counts array formula spills as populated', () => {
    expect(collectRuns(matrix, [range(3, 4, 4, 6)], false)).toEqual([
      { index: 0, start: 0, end: 10 },
      { index: 2, start: 2, end: 2 },
      { index: 3, start: 4, end: 6 },
      { index: 4, start: 4, end: 6 },
    ])
  })
})

describe('sliceRange', () => {
  const runs = [
    { index: 0, start: 0, end: 10 },
    { index: 2, start: 2, end: 2 },
    { index: 40, start: 0, end: 3 },
  ]

  it('clips runs to the scanned range and drops rows outside it', () => {
    expect(sliceRange(range(0, 1, 31, 16383), runs, false)).toEqual([
      range(0, 1, 0, 10),
      range(2, 2, 2, 2),
    ])
  })

  it('yields nothing when the range misses every run', () => {
    expect(sliceRange(range(3, 0, 39, 5), runs, false)).toEqual([])
  })

  it('slices per column in column direction', () => {
    expect(sliceRange(range(0, 0, 5, 50), runs, true)).toEqual([
      range(0, 0, 5, 0),
      range(2, 2, 2, 2),
      range(0, 40, 3, 40),
    ])
  })
})

function worksheet(matrix: Record<number, Record<number, unknown>>) {
  return {
    getSheetId: () => 's1',
    getCellMatrix: () => ({ getMatrix: () => matrix }),
  }
}

const query = (findDirection = FindDirection.ROW): IFindQuery =>
  ({ findString: 'x', findDirection }) as unknown as IFindQuery

describe('patchFindModel', () => {
  it('walks only the populated slices and keeps the results flat and ordered', () => {
    const scanned: IRange[] = []
    const original = vi.fn(function (_ws: unknown, _q: unknown, r: IRange) {
      scanned.push(r)
      return { results: [`${r.startRow}:${r.startColumn}-${r.endColumn}`] }
    })
    const proto = { _findInRange: original } as unknown as Parameters<typeof patchFindModel>[0]
    expect(patchFindModel(proto, () => [])).toBe(true)
    const ws = worksheet({ 0: { 0: { v: 1 }, 10: { v: 2 } }, 31: { 4: { v: 3 } } })
    const result = proto._findInRange!.call({}, ws as never, query(), range(0, 0, 31, 16383), 'u1')
    expect(scanned).toEqual([range(0, 0, 0, 10), range(31, 4, 31, 4)])
    expect(result.results).toEqual(['0:0-10', '31:4-4'])
  })

  it('patches a prototype once', () => {
    const proto = { _findInRange: vi.fn(() => ({ results: [] })) } as unknown as Parameters<
      typeof patchFindModel
    >[0]
    patchFindModel(proto, () => [])
    const patched = proto._findInRange
    patchFindModel(proto, () => [])
    expect(proto._findInRange).toBe(patched)
  })

  it('falls back to the dense walk when the matrix is unreadable', () => {
    const original = vi.fn(() => ({ results: ['dense'] }))
    const proto = { _findInRange: original } as unknown as Parameters<typeof patchFindModel>[0]
    patchFindModel(proto, () => [])
    const broken = { getSheetId: () => 's1', getCellMatrix: () => null }
    const result = proto._findInRange!.call({}, broken as never, query(), range(0, 0, 9, 9), 'u1')
    expect(result.results).toEqual(['dense'])
    expect(original).toHaveBeenCalledWith(broken, query(), range(0, 0, 9, 9), 'u1', undefined)
  })

  it('ignores objects that are not find models', () => {
    expect(patchFindModel({}, () => [])).toBe(false)
  })
})

describe('installSparseFind', () => {
  it('patches the model class the provider instantiates through its injector', () => {
    class FakeFindModel {
      _findInRange(_ws: unknown, _q: unknown, r: IRange) {
        return { results: [r] }
      }
    }
    const created: unknown[] = []
    const injector = {
      createInstance: vi.fn((ctor: new () => unknown) => {
        const instance = new ctor()
        created.push(instance)
        return instance
      }),
      get: vi.fn(() => 'service'),
    }
    const provider = { _injector: injector }
    expect(installSparseFind(provider, () => [])).toBe(true)

    const proxied = (provider as { _injector: typeof injector })._injector
    const model = proxied.createInstance(FakeFindModel) as FakeFindModel
    expect(created).toHaveLength(1)
    const ws = worksheet({ 3: { 1: { v: 'x' } } })
    expect(model._findInRange(ws as never, query(), range(0, 0, 99, 99))).toEqual({
      results: [range(3, 1, 3, 1)],
    })
    expect(proxied.get()).toBe('service')
    expect(installSparseFind(provider, () => [])).toBe(true)
  })

  it('leaves providers without an injector alone', () => {
    expect(installSparseFind({}, () => [])).toBe(false)
    expect(installSparseFind({ _injector: {} }, () => [])).toBe(false)
  })
})

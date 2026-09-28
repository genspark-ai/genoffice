import { describe, expect, it } from 'vitest'

import {
  DEFAULT_STATUS_BAR_FUNCS,
  filterStatusBarState,
  readStatusBarFuncs,
  STATUS_BAR_STATS_STORAGE_KEY,
  toggleStatusBarFunc,
  writeStatusBarFuncs,
} from '../src/renderer/status-bar-stats'

function memoryStorage(initial: Record<string, string> = {}) {
  const map = new Map(Object.entries(initial))
  return {
    getItem: (key: string) => map.get(key) ?? null,
    setItem: (key: string, value: string) => void map.set(key, value),
  }
}

describe('status bar statistics preference', () => {
  it('defaults to Excel ticks (Average, Count, Sum) without a stored value', () => {
    expect(readStatusBarFuncs(memoryStorage())).toEqual(['AVERAGE', 'COUNTA', 'SUM'])
    expect(readStatusBarFuncs(null)).toEqual([...DEFAULT_STATUS_BAR_FUNCS])
  })

  it('ignores corrupt or unknown stored entries', () => {
    expect(readStatusBarFuncs(memoryStorage({ [STATUS_BAR_STATS_STORAGE_KEY]: '{' }))).toEqual([
      ...DEFAULT_STATUS_BAR_FUNCS,
    ])
    expect(
      readStatusBarFuncs(
        memoryStorage({ [STATUS_BAR_STATS_STORAGE_KEY]: '["SUM","MEDIAN","MIN"]' }),
      ),
    ).toEqual(['MIN', 'SUM'])
    expect(readStatusBarFuncs(memoryStorage({ [STATUS_BAR_STATS_STORAGE_KEY]: '[]' }))).toEqual([])
  })

  it('toggles a function on and off, keeping canonical order', () => {
    const on = toggleStatusBarFunc(['AVERAGE', 'SUM'], 'MIN')
    expect(on).toEqual(['AVERAGE', 'MIN', 'SUM'])
    expect(toggleStatusBarFunc(on, 'AVERAGE')).toEqual(['MIN', 'SUM'])
    expect(toggleStatusBarFunc(['SUM'], 'SUM')).toEqual([])
  })

  it('round-trips through storage', () => {
    const storage = memoryStorage()
    writeStatusBarFuncs(storage, ['COUNT', 'MAX'])
    expect(readStatusBarFuncs(storage)).toEqual(['COUNT', 'MAX'])
  })

  it('filters upstream values to the ticked functions and keeps the pattern', () => {
    const state = {
      values: [
        { func: 'COUNTA', value: 3 },
        { func: 'SUM', value: 6 },
        { func: 'MIN', value: 1 },
      ],
      pattern: '0.00',
    }
    expect(filterStatusBarState(state, ['SUM'])).toEqual({
      values: [{ func: 'SUM', value: 6 }],
      pattern: '0.00',
    })
    expect(filterStatusBarState(null, ['SUM'])).toBeNull()
  })
})

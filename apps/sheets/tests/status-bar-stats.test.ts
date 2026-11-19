import { describe, expect, it } from 'vitest'

import {
  applyStatusBarVisibilityRules,
  DEFAULT_STATUS_BAR_FUNCS,
  filterStatusBarState,
  installStatusBarStatsFilter,
  readStatusBarFuncs,
  STATUS_BAR_STATS_STORAGE_KEY,
  toggleStatusBarFunc,
  writeStatusBarFuncs,
} from '../src/renderer/status-bar-stats'
import type { UniverRuntime } from '../src/renderer/univer-state'

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

interface TestState {
  values: { func: string; value: number }[]
  pattern: string | null
}

/** Mirrors Univer StatusBarService: COUNT/COUNTA are read from the values list. */
function univerLikeService() {
  const findValue = (state: TestState, func: string): number =>
    state.values.find((item) => item.func === func)?.value ?? 0
  const numericRule = (state: TestState): boolean =>
    findValue(state, 'COUNTA') > 1 && findValue(state, 'COUNT') > 0
  const functions = [
    { func: 'MAX', filter: numericRule },
    { func: 'MIN', filter: numericRule },
    { func: 'SUM', filter: numericRule },
    { func: 'COUNTA', filter: (state: TestState) => findValue(state, 'COUNTA') > 1 },
    { func: 'COUNT', filter: numericRule },
    { func: 'AVERAGE', filter: numericRule },
  ]
  const published: (TestState | null)[] = []
  const service = {
    getFunctions: () => functions,
    setState(state: TestState | null) {
      if (!state) {
        published.push(null)
        return
      }
      published.push({
        values: state.values.filter((item) => {
          const entry = functions.find((candidate) => candidate.func === item.func)
          return entry !== undefined && (entry.filter === undefined || entry.filter(state))
        }),
        pattern: state.pattern,
      })
    },
  }
  const runtime = {
    univer: { __getInjector: () => ({ get: () => service }) },
  } as unknown as UniverRuntime
  return { service, published, runtime, functions }
}

const FULL_STATE: TestState = {
  values: [
    { func: 'MAX', value: 8 },
    { func: 'MIN', value: 1 },
    { func: 'SUM', value: 15 },
    { func: 'COUNTA', value: 5 },
    { func: 'COUNT', value: 4 },
    { func: 'AVERAGE', value: 3.75 },
  ],
  pattern: null,
}

describe('status bar visibility vs preference filter', () => {
  it('keeps SUM and AVERAGE visible when COUNT is unticked (default prefs)', () => {
    const { service, published, runtime } = univerLikeService()
    installStatusBarStatsFilter(runtime, () => [...DEFAULT_STATUS_BAR_FUNCS])
    service.setState(FULL_STATE)
    expect(published.at(-1)).toEqual({
      values: [
        { func: 'SUM', value: 15 },
        { func: 'COUNTA', value: 5 },
        { func: 'AVERAGE', value: 3.75 },
      ],
      pattern: null,
    })
  })

  it('still hides numeric statistics when the selection has no numbers', () => {
    const { service, published, runtime } = univerLikeService()
    installStatusBarStatsFilter(runtime, () => [...DEFAULT_STATUS_BAR_FUNCS])
    service.setState({
      values: [
        { func: 'SUM', value: 0 },
        { func: 'COUNTA', value: 3 },
        { func: 'COUNT', value: 0 },
        { func: 'AVERAGE', value: Number.NaN },
      ],
      pattern: null,
    })
    expect(published.at(-1)).toEqual({ values: [{ func: 'COUNTA', value: 3 }], pattern: null })
  })

  it('hides everything on a single-cell selection, Univer-style', () => {
    const { service, published, runtime } = univerLikeService()
    installStatusBarStatsFilter(runtime, () => [...DEFAULT_STATUS_BAR_FUNCS])
    service.setState({
      values: [
        { func: 'SUM', value: 7 },
        { func: 'COUNTA', value: 1 },
        { func: 'COUNT', value: 1 },
      ],
      pattern: null,
    })
    expect(published.at(-1)).toEqual({ values: [], pattern: null })
  })

  it('refresh() re-applies the full upstream state after a preference change', () => {
    const { service, published, runtime } = univerLikeService()
    let enabled: string[] = [...DEFAULT_STATUS_BAR_FUNCS]
    const filter = installStatusBarStatsFilter(runtime, () => enabled as never)
    service.setState(FULL_STATE)
    enabled = ['MIN', 'MAX']
    filter.refresh()
    expect(published.at(-1)).toEqual({
      values: [
        { func: 'MAX', value: 8 },
        { func: 'MIN', value: 1 },
      ],
      pattern: null,
    })
  })

  it('dispose() restores the service rules and setState', () => {
    const { service, published, runtime, functions } = univerLikeService()
    const originalFilters = functions.map((entry) => entry.filter)
    const filter = installStatusBarStatsFilter(runtime, () => [...DEFAULT_STATUS_BAR_FUNCS])
    expect(functions.every((entry) => entry.filter === undefined)).toBe(true)
    filter.dispose()
    expect(functions.map((entry) => entry.filter)).toEqual(originalFilters)
    service.setState(FULL_STATE)
    expect(published.at(-1)).toEqual(FULL_STATE)
  })
})

describe('applyStatusBarVisibilityRules', () => {
  it('passes entries with no rule and evaluates rules against the full state', () => {
    const rules = new Map<string, ((state: TestState) => boolean) | undefined>([
      ['SUM', (state) => state.values.some((item) => item.func === 'COUNT' && item.value > 0)],
      ['COUNTA', undefined],
    ])
    const state: TestState = {
      values: [
        { func: 'SUM', value: 5 },
        { func: 'COUNTA', value: 2 },
        { func: 'COUNT', value: 2 },
      ],
      pattern: '0.0',
    }
    expect(applyStatusBarVisibilityRules(state, rules as never)).toEqual(state)
    expect(applyStatusBarVisibilityRules(null, rules as never)).toBeNull()
  })
})

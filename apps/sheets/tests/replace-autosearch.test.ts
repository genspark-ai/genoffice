import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { FindBy, FindDirection, FindScope, type IFindReplaceState } from '@univerjs/find-replace'
import { BehaviorSubject } from 'rxjs'

import { installReplaceAutoSearch, planAutoSearch } from '../src/renderer/replace-autosearch'

function state(overrides: Partial<IFindReplaceState> = {}): IFindReplaceState {
  return {
    revealed: true,
    replaceRevealed: true,
    findString: '',
    inputtingFindString: '',
    replaceString: '',
    matchesPosition: 0,
    matchesCount: 0,
    findCompleted: false,
    caseSensitive: false,
    matchesTheWholeCell: false,
    findBy: FindBy.VALUE,
    findScope: FindScope.SUBUNIT,
    findDirection: FindDirection.ROW,
    ...overrides,
  }
}

describe('planAutoSearch', () => {
  it('commits typed input that has not been searched yet', () => {
    expect(planAutoSearch(null, state({ inputtingFindString: '1350' }))).toBe('commit-and-search')
  })

  it('ignores find mode, hidden dialog and blank input', () => {
    expect(planAutoSearch(null, state({ replaceRevealed: false, inputtingFindString: '1' }))).toBe(
      null,
    )
    expect(planAutoSearch(null, state({ revealed: false, inputtingFindString: '1' }))).toBe(null)
    expect(planAutoSearch(null, state({ inputtingFindString: '   ' }))).toBe(null)
  })

  it('stays quiet once the input matches the searched string', () => {
    const searched = state({ findString: '1350', inputtingFindString: '1350', matchesCount: 2 })
    expect(planAutoSearch(searched, searched)).toBe(null)
    expect(planAutoSearch(null, searched)).toBe(null)
  })

  it('re-searches when an option toggles after the session was stopped', () => {
    const before = state({ findString: '1350', inputtingFindString: '1350' })
    const after = { ...before, matchesTheWholeCell: true }
    expect(planAutoSearch(before, after)).toBe('search')
  })

  it('does not re-search after Replace All stops the session', () => {
    const before = state({ findString: '1350', inputtingFindString: '1350', matchesCount: 2 })
    const after = { ...before, matchesCount: 0, findCompleted: false }
    expect(planAutoSearch(before, after)).toBe(null)
  })
})

describe('installReplaceAutoSearch', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  it('debounces typing, then commits and searches once', () => {
    const state$ = new BehaviorSubject(state())
    const service = { state$, changeFindString: vi.fn(), find: vi.fn() }
    const handle = installReplaceAutoSearch(service, 300)

    state$.next(state({ inputtingFindString: '1' }))
    vi.advanceTimersByTime(100)
    state$.next(state({ inputtingFindString: '13' }))
    vi.advanceTimersByTime(100)
    state$.next(state({ inputtingFindString: '1350' }))
    vi.advanceTimersByTime(299)
    expect(service.find).not.toHaveBeenCalled()

    vi.advanceTimersByTime(1)
    expect(service.changeFindString).toHaveBeenCalledWith('1350')
    expect(service.changeFindString).toHaveBeenCalledTimes(1)
    expect(service.find).toHaveBeenCalledTimes(1)

    // Univer echoes the committed string back; no second search.
    state$.next(state({ findString: '1350', inputtingFindString: '1350', matchesCount: 3 }))
    vi.advanceTimersByTime(300)
    expect(service.find).toHaveBeenCalledTimes(1)

    handle.dispose()
    state$.next(state({ findString: '1350', inputtingFindString: '1350x' }))
    vi.advanceTimersByTime(300)
    expect(service.find).toHaveBeenCalledTimes(1)
  })

  it('is not starved by match-count churn from a running scan', () => {
    const state$ = new BehaviorSubject(state())
    const service = { state$, changeFindString: vi.fn(), find: vi.fn() }
    installReplaceAutoSearch(service, 300)

    state$.next(state({ inputtingFindString: '1350' }))
    for (let i = 1; i <= 5; i++) {
      vi.advanceTimersByTime(100)
      state$.next(state({ inputtingFindString: '1350', matchesCount: i, matchesPosition: 1 }))
    }
    expect(service.find).toHaveBeenCalledTimes(1)
    expect(service.changeFindString).toHaveBeenCalledWith('1350')
  })

  it('searches without recommitting when only an option changed', () => {
    const searched = state({ findString: '1350', inputtingFindString: '1350', matchesCount: 3 })
    const state$ = new BehaviorSubject(searched)
    const service = { state$, changeFindString: vi.fn(), find: vi.fn() }
    installReplaceAutoSearch(service, 300)
    vi.advanceTimersByTime(300)

    state$.next({ ...searched, caseSensitive: true, matchesCount: 0 })
    vi.advanceTimersByTime(300)
    expect(service.changeFindString).not.toHaveBeenCalled()
    expect(service.find).toHaveBeenCalledTimes(1)
  })
})

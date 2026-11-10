import { beforeEach, describe, expect, it, vi } from 'vitest'

import { startFullLoad } from '../src/renderer/full-load'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import { preloadEntireWorkbook } from '../src/renderer/univer-sync'

vi.mock('../src/renderer/univer-sync', () => ({ preloadEntireWorkbook: vi.fn() }))

const runtime = {} as UniverRuntime

function lazyState(preloadComplete = false): LazyWorkbookState {
  return { flags: { preloadComplete, preloadRunning: false } } as LazyWorkbookState
}

/// A preload the test finishes by hand, flagged the way the real one is.
function manualPreload(state: LazyWorkbookState) {
  let finish!: () => void
  let fail!: (reason: unknown) => void
  vi.mocked(preloadEntireWorkbook).mockImplementationOnce(() => {
    state.flags.preloadRunning = true
    return new Promise<void>((resolve, reject) => {
      finish = () => {
        state.flags.preloadRunning = false
        state.flags.preloadComplete = true
        resolve()
      }
      fail = (reason) => {
        state.flags.preloadRunning = false
        reject(reason)
      }
    })
  })
  return { finish: () => finish(), fail: (reason: unknown) => fail(reason) }
}

describe('startFullLoad', () => {
  beforeEach(() => {
    vi.mocked(preloadEntireWorkbook).mockReset()
  })

  it('runs one preload for concurrent requests and hands every caller the same load', async () => {
    const state = lazyState()
    const ref = { current: state }
    const preload = manualPreload(state)

    const first = startFullLoad(runtime, ref, () => undefined)
    const second = startFullLoad(runtime, ref, () => undefined)

    expect(second).toBe(first)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
    preload.finish()
    await Promise.all([first, second])
    expect(state.flags.preloadComplete).toBe(true)
  })

  it('does nothing when the workbook is already fully loaded or none is open', async () => {
    await startFullLoad(runtime, { current: lazyState(true) }, () => undefined)
    await startFullLoad(runtime, { current: null }, () => undefined)
    expect(preloadEntireWorkbook).not.toHaveBeenCalled()
  })

  it('lets a new load start after a failed one', async () => {
    const state = lazyState()
    const ref = { current: state }
    const failing = manualPreload(state)
    const first = startFullLoad(runtime, ref, () => undefined)
    failing.fail(new Error('disk'))
    await expect(first).rejects.toThrow('disk')

    const retry = manualPreload(state)
    const second = startFullLoad(runtime, ref, () => undefined)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(2)
    retry.finish()
    await second
  })

  it('keeps loads of different workbooks apart', async () => {
    const stateA = lazyState()
    const stateB = lazyState()
    const preloadA = manualPreload(stateA)
    const loadA = startFullLoad(runtime, { current: stateA }, () => undefined)
    const preloadB = manualPreload(stateB)
    const loadB = startFullLoad(runtime, { current: stateB }, () => undefined)

    expect(loadB).not.toBe(loadA)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(2)
    preloadA.finish()
    preloadB.finish()
    await Promise.all([loadA, loadB])
  })
})

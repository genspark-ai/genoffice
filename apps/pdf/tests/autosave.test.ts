import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AUTOSAVE_INTERVAL_MS, startAutosave } from '../src/renderer/useAutosave'

describe('autosave scheduling', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  it('saves on each 30s tick while there are pending changes', () => {
    const save = vi.fn()
    const stop = startAutosave(() => true, save)
    vi.advanceTimersByTime(AUTOSAVE_INTERVAL_MS)
    expect(save).toHaveBeenCalledTimes(1)
    vi.advanceTimersByTime(AUTOSAVE_INTERVAL_MS)
    expect(save).toHaveBeenCalledTimes(2)
    stop()
  })

  it('skips ticks while shouldSave reports nothing to persist', () => {
    const save = vi.fn()
    let dirty = false
    const stop = startAutosave(() => dirty, save)
    vi.advanceTimersByTime(AUTOSAVE_INTERVAL_MS * 2)
    expect(save).not.toHaveBeenCalled()
    dirty = true
    vi.advanceTimersByTime(AUTOSAVE_INTERVAL_MS)
    expect(save).toHaveBeenCalledTimes(1)
    stop()
  })

  it('saves on window blur only when there are pending changes', () => {
    const save = vi.fn()
    let dirty = false
    const stop = startAutosave(() => dirty, save)
    window.dispatchEvent(new Event('blur'))
    expect(save).not.toHaveBeenCalled()
    dirty = true
    window.dispatchEvent(new Event('blur'))
    expect(save).toHaveBeenCalledTimes(1)
    stop()
  })

  it('stop detaches both the interval and the blur listener', () => {
    const save = vi.fn()
    const stop = startAutosave(() => true, save)
    stop()
    vi.advanceTimersByTime(AUTOSAVE_INTERVAL_MS * 3)
    window.dispatchEvent(new Event('blur'))
    expect(save).not.toHaveBeenCalled()
  })
})

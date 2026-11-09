import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createDeferredDelete } from '../src/renderer/src/deferred-delete'

describe('createDeferredDelete', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  const setup = (commit = vi.fn(async () => undefined)) => {
    const changes: string[][] = []
    const dd = createDeferredDelete({
      delayMs: 6000,
      commit,
      onChange: (s) => changes.push([...s].sort()),
    })
    return { dd, commit, changes }
  }

  it('hides paths immediately and commits only after the delay', async () => {
    const { dd, commit, changes } = setup()
    dd.schedule(['a', 'b'])
    expect(changes.at(-1)).toEqual(['a', 'b'])
    await vi.advanceTimersByTimeAsync(5999)
    expect(commit).not.toHaveBeenCalled()
    await vi.advanceTimersByTimeAsync(1)
    expect(commit).toHaveBeenCalledWith(['a', 'b'])
  })

  it('keeps rows hidden until the commit settles (no flash back)', async () => {
    let finish: () => void = () => undefined
    const { dd, changes } = setup(vi.fn(() => new Promise<void>((r) => (finish = r))))
    dd.schedule(['a'])
    await vi.advanceTimersByTimeAsync(6000)
    expect(dd.pending().has('a')).toBe(true)
    finish()
    await vi.advanceTimersByTimeAsync(0)
    expect(dd.pending().size).toBe(0)
    expect(changes.at(-1)).toEqual([])
  })

  it('undo cancels the commit and restores the rows', async () => {
    const { dd, commit } = setup()
    const id = dd.schedule(['a'])
    expect(dd.undo(id)).toEqual(['a'])
    await vi.advanceTimersByTimeAsync(10000)
    expect(commit).not.toHaveBeenCalled()
    expect(dd.pending().size).toBe(0)
  })

  it('cannot undo once the commit has started', async () => {
    const { dd } = setup(vi.fn(() => new Promise<void>(() => undefined)))
    const id = dd.schedule(['a'])
    await vi.advanceTimersByTimeAsync(6000)
    expect(dd.undo(id)).toBeNull()
  })

  it('tracks overlapping batches independently', async () => {
    const { dd, commit } = setup()
    const first = dd.schedule(['a'])
    await vi.advanceTimersByTimeAsync(2000)
    dd.schedule(['b'])
    dd.undo(first)
    expect([...dd.pending()]).toEqual(['b'])
    await vi.advanceTimersByTimeAsync(6000)
    expect(commit).toHaveBeenCalledTimes(1)
    expect(commit).toHaveBeenCalledWith(['b'])
  })

  it('flush commits everything now', async () => {
    const { dd, commit } = setup()
    dd.schedule(['a'])
    dd.schedule(['b'])
    dd.flush()
    await vi.advanceTimersByTimeAsync(0)
    expect(commit).toHaveBeenCalledTimes(2)
  })
})

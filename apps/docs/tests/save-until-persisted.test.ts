import { describe, expect, it, vi } from 'vitest'
import { createSaveSerializer, saveUntilPersisted } from '../src/renderer/save-until-persisted'

/** deps whose save() reports "incomplete" for the first `incompletePasses` calls */
const depsWith = (incompletePasses: number, hasPath = true) => {
  let calls = 0
  const save = vi.fn(async () => {
    calls++
    return true
  })
  return {
    save,
    wasIncomplete: () => calls <= incompletePasses,
    hasPath: () => hasPath,
  }
}

describe('saveUntilPersisted', () => {
  it('reports success after a single fully persisted save', async () => {
    const deps = depsWith(0)
    await expect(saveUntilPersisted(deps)).resolves.toBe(true)
    expect(deps.save).toHaveBeenCalledTimes(1)
  })

  it('saves again when the editor changed mid-save, then reports success', async () => {
    const deps = depsWith(1)
    await expect(saveUntilPersisted(deps)).resolves.toBe(true)
    expect(deps.save).toHaveBeenCalledTimes(2)
  })

  // the regression: an incomplete save must never be reported as complete, or the
  // close guard closes the window over edits that never reached disk
  it('reports failure when the document never catches up', async () => {
    const deps = depsWith(Number.POSITIVE_INFINITY)
    await expect(saveUntilPersisted(deps)).resolves.toBe(false)
    expect(deps.save).toHaveBeenCalledTimes(3)
  })

  it('honours a custom pass budget', async () => {
    const deps = depsWith(Number.POSITIVE_INFINITY)
    await expect(saveUntilPersisted({ ...deps, maxPasses: 5 })).resolves.toBe(false)
    expect(deps.save).toHaveBeenCalledTimes(5)
  })

  it('propagates a failed save without retrying', async () => {
    const save = vi.fn(async () => false)
    const deps = { save, wasIncomplete: () => true, hasPath: () => true }
    await expect(saveUntilPersisted(deps)).resolves.toBe(false)
    expect(save).toHaveBeenCalledTimes(1)
  })

  it('does not retry a pathless document (Save As would prompt again)', async () => {
    const deps = depsWith(Number.POSITIVE_INFINITY, false)
    await expect(saveUntilPersisted(deps)).resolves.toBe(true)
    expect(deps.save).toHaveBeenCalledTimes(1)
  })
})

/** a save pass the test resolves manually */
const manualPass = () => {
  let resolve!: (ok: boolean) => void
  const promise = new Promise<boolean>((r) => {
    resolve = r
  })
  return { promise, resolve }
}

describe('createSaveSerializer', () => {
  it('runs immediately when nothing is in flight', async () => {
    const run = createSaveSerializer()
    await expect(
      run(
        async () => true,
        () => true,
      ),
    ).resolves.toBe(true)
  })

  // the regression: a save arriving mid-flight returned false immediately, so
  // "Save and close" during a blur autosave silently gave up
  it('waits for the in-flight save and reuses its clean result', async () => {
    const run = createSaveSerializer()
    const first = manualPass()
    const p1 = run(
      () => first.promise,
      () => true,
    )
    const second = vi.fn(async () => true)
    const p2 = run(second, () => true)
    first.resolve(true)
    await expect(p1).resolves.toBe(true)
    await expect(p2).resolves.toBe(true)
    expect(second).not.toHaveBeenCalled()
  })

  it('runs a follow-up pass when the finished save left dirty state behind', async () => {
    const run = createSaveSerializer()
    const first = manualPass()
    let dirty = true
    const p1 = run(
      () => first.promise,
      () => !dirty,
    )
    const second = vi.fn(async () => {
      dirty = false
      return true
    })
    const p2 = run(second, () => !dirty)
    first.resolve(true)
    await expect(p1).resolves.toBe(true)
    await expect(p2).resolves.toBe(true)
    expect(second).toHaveBeenCalledTimes(1)
  })

  it('retries after a failed in-flight save and reports the real outcome', async () => {
    const run = createSaveSerializer()
    const first = manualPass()
    const p1 = run(
      () => first.promise,
      () => true,
    )
    const second = vi.fn(async () => false)
    const p2 = run(second, () => true)
    first.resolve(false)
    await expect(p1).resolves.toBe(false)
    await expect(p2).resolves.toBe(false)
    expect(second).toHaveBeenCalledTimes(1)
  })

  it('serializes three overlapping callers onto one flight when all are clean', async () => {
    const run = createSaveSerializer()
    const first = manualPass()
    const runs: number[] = []
    const p1 = run(
      () => (runs.push(1), first.promise),
      () => true,
    )
    const p2 = run(
      async () => (runs.push(2), true),
      () => true,
    )
    const p3 = run(
      async () => (runs.push(3), true),
      () => true,
    )
    first.resolve(true)
    await expect(Promise.all([p1, p2, p3])).resolves.toEqual([true, true, true])
    expect(runs).toEqual([1])
  })

  // the regression: several waiters resumed together after the in-flight pass
  // settled, all saw the cleared slot, and started overlapping save passes
  it('waiters that each need their own pass run strictly one after another', async () => {
    const run = createSaveSerializer()
    const first = manualPass()
    let active = 0
    let maxActive = 0
    const tracked = (label: number, order: number[]) => async () => {
      order.push(label)
      active++
      maxActive = Math.max(maxActive, active)
      await new Promise((r) => setTimeout(r, 0))
      active--
      return true
    }
    const order: number[] = []
    const dirtyForever = () => false
    const p1 = run(() => first.promise, dirtyForever)
    const p2 = run(tracked(2, order), dirtyForever)
    const p3 = run(tracked(3, order), dirtyForever)
    first.resolve(true)
    await expect(Promise.all([p1, p2, p3])).resolves.toEqual([true, true, true])
    expect(order).toEqual([2, 3])
    expect(maxActive).toBe(1)
  })

  it('treats a rejected pass as a failure and lets the waiter retry', async () => {
    const run = createSaveSerializer()
    const p1 = run(
      () => Promise.reject(new Error('boom')),
      () => true,
    )
    const second = vi.fn(async () => true)
    const p2 = run(second, () => true)
    await expect(p1).rejects.toThrow('boom')
    await expect(p2).resolves.toBe(true)
    expect(second).toHaveBeenCalledTimes(1)
  })
})

import { describe, expect, it, vi } from 'vitest'
import { runGuardedPrint, type PrintGuard } from '../src/renderer/print-guard'
import type { PrintResult } from '../src/shared/ipc'

/** A print that stays in flight until the test releases it, so a second request
 * can be attempted while the first is still open at the system dialog. */
function deferred<T>() {
  let release!: (value: T) => void
  const promise = new Promise<T>((resolve) => {
    release = resolve
  })
  return { promise, release }
}

const guard = (): PrintGuard => ({ current: false })

describe('print in-flight guard', () => {
  it('a second request while one is in flight does not start a second print', async () => {
    const g = guard()
    const first = deferred<PrintResult>()
    const print = vi.fn(() => first.promise)
    const onFailure = vi.fn()

    const a = runGuardedPrint(g, print, onFailure)
    expect(g.current).toBe(true)

    // The system dialog is modal: without the gate this opens a second window
    // and a second dialog for the same document.
    expect(await runGuardedPrint(g, print, onFailure)).toBe(false)
    expect(print).toHaveBeenCalledTimes(1)

    first.release({ ok: true })
    expect(await a).toBe(true)
    expect(print).toHaveBeenCalledTimes(1)
  })

  it('clears the flag after a success', async () => {
    const g = guard()
    const onFailure = vi.fn()
    expect(await runGuardedPrint(g, async () => ({ ok: true }), onFailure)).toBe(true)
    expect(g.current).toBe(false)
    expect(onFailure).not.toHaveBeenCalled()
    // Free again: the next keystroke is not swallowed by a leaked flag.
    expect(await runGuardedPrint(g, async () => ({ ok: true }), onFailure)).toBe(true)
  })

  it('clears the flag after a failure and reports the reason', async () => {
    const g = guard()
    const onFailure = vi.fn()
    const result = await runGuardedPrint(
      g,
      async () => ({ ok: false, error: 'Printer not available' }),
      onFailure,
    )
    expect(result).toBe(false)
    expect(onFailure).toHaveBeenCalledWith('Printer not available')
    expect(g.current).toBe(false)
    expect(await runGuardedPrint(g, async () => ({ ok: true }), onFailure)).toBe(true)
  })

  it('clears the flag after a cancellation and stays silent about it', async () => {
    const g = guard()
    const onFailure = vi.fn()
    // The user closed the dialog: not printed, but nothing to report.
    const result = await runGuardedPrint(g, async () => ({ ok: true, canceled: true }), onFailure)
    expect(result).toBe(false)
    expect(onFailure).not.toHaveBeenCalled()
    expect(g.current).toBe(false)
    expect(await runGuardedPrint(g, async () => ({ ok: true }), onFailure)).toBe(true)
  })

  it('clears the flag when printing throws', async () => {
    const g = guard()
    const onFailure = vi.fn()
    const result = await runGuardedPrint(
      g,
      async () => {
        throw new Error('ipc channel closed')
      },
      onFailure,
    )
    expect(result).toBe(false)
    expect(onFailure).toHaveBeenCalledWith('ipc channel closed')
    expect(g.current).toBe(false)
    expect(await runGuardedPrint(g, async () => ({ ok: true }), onFailure)).toBe(true)
  })
})

import { describe, expect, it } from 'vitest'
import { abortOnDestroyed, type DestroyableSender } from '../src/abort-on-destroyed'

function fakeSender(destroyed = false) {
  const listeners = new Set<() => void>()
  const sender: DestroyableSender & { destroy(): void; listenerCount(): number } = {
    isDestroyed: () => destroyed,
    once: (_e, l) => {
      listeners.add(l)
    },
    removeListener: (_e, l) => {
      listeners.delete(l)
    },
    destroy() {
      destroyed = true
      for (const l of [...listeners]) {
        listeners.delete(l)
        l()
      }
    },
    listenerCount: () => listeners.size,
  }
  return sender
}

describe('abortOnDestroyed', () => {
  it('aborts the controller when the sender is destroyed mid-stream', () => {
    const sender = fakeSender()
    const controller = new AbortController()
    abortOnDestroyed(sender, controller)
    expect(controller.signal.aborted).toBe(false)
    sender.destroy()
    expect(controller.signal.aborted).toBe(true)
  })

  it('removes the listener when the stream settles first', () => {
    const sender = fakeSender()
    const controller = new AbortController()
    const dispose = abortOnDestroyed(sender, controller)
    expect(sender.listenerCount()).toBe(1)
    dispose()
    expect(sender.listenerCount()).toBe(0)
    sender.destroy()
    expect(controller.signal.aborted).toBe(false)
  })

  it('aborts immediately for an already-destroyed sender', () => {
    const controller = new AbortController()
    abortOnDestroyed(fakeSender(true), controller)
    expect(controller.signal.aborted).toBe(true)
  })
})

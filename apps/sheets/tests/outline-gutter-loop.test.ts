import { describe, expect, it } from 'vitest'

import { paintLoop } from '../src/renderer/outline-gutter'

function fakeFrames() {
  const pending = new Map<number, () => void>()
  let next = 1
  return {
    requestFrame: (callback: () => void) => {
      const id = next
      next += 1
      pending.set(id, callback)
      return id
    },
    cancelFrame: (handle: number) => {
      pending.delete(handle)
    },
    flush() {
      const callbacks = [...pending.values()]
      pending.clear()
      for (const callback of callbacks) callback()
    },
    get pending() {
      return pending.size
    },
  }
}

describe('paintLoop: the gutter paints on demand and sleeps when still', () => {
  it('runs one frame per schedule and stops once the geometry is stable', () => {
    const frames = fakeFrames()
    let paints = 0
    const loop = paintLoop(frames.requestFrame, frames.cancelFrame, () => {
      paints += 1
      return paints < 2 // moving on the first paint, stable on the second
    })
    loop.schedule()
    expect(frames.pending).toBe(1)
    frames.flush()
    // still moving: the loop rode along for one more frame
    expect(frames.pending).toBe(1)
    frames.flush()
    // stable: the loop went to sleep
    expect(frames.pending).toBe(0)
    frames.flush()
    expect(paints).toBe(2)
  })

  it('coalesces a burst of triggers into a single frame', () => {
    const frames = fakeFrames()
    let paints = 0
    const loop = paintLoop(frames.requestFrame, frames.cancelFrame, () => {
      paints += 1
      return false
    })
    loop.schedule()
    loop.schedule()
    loop.schedule()
    expect(frames.pending).toBe(1)
    frames.flush()
    expect(paints).toBe(1)
    expect(frames.pending).toBe(0)
  })

  it('a trigger after sleep schedules exactly one more frame', () => {
    const frames = fakeFrames()
    const loop = paintLoop(frames.requestFrame, frames.cancelFrame, () => false)
    loop.schedule()
    frames.flush()
    expect(frames.pending).toBe(0)
    loop.schedule()
    expect(frames.pending).toBe(1)
    frames.flush()
    expect(frames.pending).toBe(0)
  })

  it('dispose cancels the pending frame and ignores later schedules', () => {
    const frames = fakeFrames()
    let paints = 0
    const loop = paintLoop(frames.requestFrame, frames.cancelFrame, () => {
      paints += 1
      return false
    })
    loop.schedule()
    loop.dispose()
    expect(frames.pending).toBe(0)
    loop.schedule()
    expect(frames.pending).toBe(0)
    frames.flush()
    expect(paints).toBe(0)
  })
})

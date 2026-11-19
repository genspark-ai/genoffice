/** @vitest-environment jsdom */
import { describe, expect, it, vi } from 'vitest'
import { SelectionControl } from '@univerjs/sheets-ui'
import {
  CopyAnimationIdle,
  installCopyAnimationIdle,
  type AntLineControl,
} from '../src/renderer/copy-animation-idle'

/**
 * The guard on Univer's marching-ants border.
 *
 * The animation is a recursive frame chain: the selection control asks for a
 * frame, moves the dash offset, and asks for the next. That is why the guard
 * is asked once per frame — a frame it does not let through is the last frame
 * of the chain, which is how the border stops without Univer knowing.
 */
function fakeControl(): AntLineControl & { start: ReturnType<typeof vi.fn> } {
  const start = vi.fn()
  return {
    start,
    _startAntLineAnimation: start,
    _stopAntLineAnimation: vi.fn(),
  }
}

describe('CopyAnimationIdle', () => {
  it('lets the frame through while the reader is there', () => {
    const control = fakeControl()
    const guard = new CopyAnimationIdle({ isVisible: () => true, now: () => 0 })
    const step = vi.fn()
    guard.frame(control, step)
    expect(step).toHaveBeenCalledTimes(1)
  })

  it('does not let a frame through while the document is off screen', () => {
    const control = fakeControl()
    const guard = new CopyAnimationIdle({ isVisible: () => false, now: () => 0 })
    const step = vi.fn()
    guard.frame(control, step)
    expect(step).not.toHaveBeenCalled()
  })

  it('ends the frame chain once the reader has been idle for a while', () => {
    const control = fakeControl()
    let clock = 0
    const guard = new CopyAnimationIdle({ idleMs: 100, isVisible: () => true, now: () => clock })
    const step = vi.fn()
    guard.frame(control, step)
    expect(step).toHaveBeenCalledTimes(1)
    clock = 200
    // the next frame of the chain finds the reader gone and does not continue it
    guard.frame(control, step)
    expect(step).toHaveBeenCalledTimes(1)
  })

  it('starts a border that stopped moving again when the reader comes back', () => {
    const control = fakeControl()
    let clock = 0
    const guard = new CopyAnimationIdle({ idleMs: 100, isVisible: () => true, now: () => clock })
    const step = vi.fn()
    guard.frame(control, step)
    clock = 200
    guard.frame(control, step)
    guard.resume()
    expect(control._startAntLineAnimation).toHaveBeenCalledTimes(1)
  })

  it('does not start a second chain for a border that is already moving', () => {
    const control = fakeControl()
    const guard = new CopyAnimationIdle({ isVisible: () => true, now: () => 0 })
    guard.frame(control, vi.fn())
    guard.resume()
    expect(control._startAntLineAnimation).not.toHaveBeenCalled()
  })

  it('cancels the chain of a border that was moving when the document went off screen', () => {
    const control = fakeControl()
    const guard = new CopyAnimationIdle({ isVisible: () => true, now: () => 0 })
    guard.frame(control, vi.fn())
    guard.pause()
    expect(control._stopAntLineAnimation).toHaveBeenCalledTimes(1)
  })

  it('starts the border again after it was paused', () => {
    const control = fakeControl()
    const guard = new CopyAnimationIdle({ isVisible: () => true, now: () => 0 })
    guard.frame(control, vi.fn())
    guard.pause()
    guard.resume()
    expect(control._startAntLineAnimation).toHaveBeenCalledTimes(1)
  })

  it('forgets a control that has been disposed', () => {
    const control = fakeControl()
    control._disposed = true
    const guard = new CopyAnimationIdle({ isVisible: () => true, now: () => 0 })
    guard.frame(control, vi.fn())
    guard.pause()
    expect(control._stopAntLineAnimation).not.toHaveBeenCalled()
  })

  it('picks the border up again after a glance at another window', () => {
    const control = fakeControl()
    let clock = 0
    const guard = new CopyAnimationIdle({ idleMs: 100, isVisible: () => true, now: () => clock })
    guard.frame(control, vi.fn())
    clock = 10
    guard.away()
    expect(control._stopAntLineAnimation).toHaveBeenCalledTimes(1)
    guard.back()
    expect(control._startAntLineAnimation).toHaveBeenCalledTimes(1)
  })

  it('leaves the border still after the machine was asleep', () => {
    // The reported case: a copy mode that outlived the reader overnight. The
    // border must not start moving again by itself when the machine wakes up —
    // that is the loop that filled the renderer with gigabytes.
    const control = fakeControl()
    let clock = 0
    const guard = new CopyAnimationIdle({ idleMs: 100, isVisible: () => true, now: () => clock })
    guard.frame(control, vi.fn())
    clock = 10
    guard.away()
    clock = 10 + 8 * 60 * 60 * 1000
    guard.back()
    expect(control._startAntLineAnimation).not.toHaveBeenCalled()
  })
})

describe('installCopyAnimationIdle', () => {
  it('guards the animation method and puts it back when uninstalled', () => {
    const proto = SelectionControl.prototype as unknown as {
      _startAntLineAnimation: () => void
    }
    const original = proto._startAntLineAnimation
    const step = vi.fn()
    proto._startAntLineAnimation = step
    try {
      const uninstall = installCopyAnimationIdle({ isVisible: () => false })
      // the guard sits in front of the animation: an off-screen document does
      // not get a frame, so Univer's own step never runs
      proto._startAntLineAnimation.call(fakeControl())
      expect(step).not.toHaveBeenCalled()
      uninstall()
      expect(proto._startAntLineAnimation).toBe(step)
    } finally {
      proto._startAntLineAnimation = original
    }
  })

  it('lets the animation run while the document is on screen', () => {
    const proto = SelectionControl.prototype as unknown as {
      _startAntLineAnimation: () => void
    }
    const original = proto._startAntLineAnimation
    const step = vi.fn()
    proto._startAntLineAnimation = step
    try {
      const uninstall = installCopyAnimationIdle({ isVisible: () => true })
      proto._startAntLineAnimation.call(fakeControl())
      expect(step).toHaveBeenCalledTimes(1)
      uninstall()
    } finally {
      proto._startAntLineAnimation = original
    }
  })

  it('does not wrap the animation in itself when installed twice', () => {
    // The install runs from an effect, so it is reached again on every
    // re-render. A second wrap would put the guard in front of the guard,
    // and every frame would run through both of them for the rest of the
    // session.
    const proto = SelectionControl.prototype as unknown as {
      _startAntLineAnimation: () => void
    }
    const original = proto._startAntLineAnimation
    const step = vi.fn()
    proto._startAntLineAnimation = step
    try {
      const uninstall = installCopyAnimationIdle({ isVisible: () => true })
      const guarded = proto._startAntLineAnimation
      installCopyAnimationIdle({ isVisible: () => true })
      expect(proto._startAntLineAnimation).toBe(guarded)
      uninstall()
    } finally {
      proto._startAntLineAnimation = original
    }
  })
})

// Stage wheel paging: discrete mouse-wheel notches flip one page each, trackpad
// swipes accumulate to one flip per gesture with the momentum tail swallowed,
// and a direction reversal or quiet gap starts a fresh gesture.
import { describe, expect, it } from 'vitest'
import { createWheelPager } from '../src/renderer/wheel-page-flip'

describe('createWheelPager', () => {
  it('flips next on a single mouse-wheel notch down', () => {
    const pager = createWheelPager()
    expect(pager.feed(100, 0)).toBe(1)
  })

  it('flips previous on a wheel notch up', () => {
    const pager = createWheelPager()
    expect(pager.feed(-100, 0)).toBe(-1)
  })

  it('flips once per notch while the wheel keeps spinning', () => {
    const pager = createWheelPager()
    expect(pager.feed(100, 0)).toBe(1)
    expect(pager.feed(100, 80)).toBe(1)
    expect(pager.feed(100, 160)).toBe(1)
  })

  it('accumulates small trackpad deltas into a single flip', () => {
    const pager = createWheelPager()
    expect(pager.feed(10, 0)).toBe(0)
    expect(pager.feed(20, 16)).toBe(0)
    expect(pager.feed(40, 32)).toBe(1)
  })

  it('swallows the momentum tail after a trackpad flip', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    pager.feed(20, 16)
    expect(pager.feed(40, 32)).toBe(1)
    // macOS momentum keeps emitting large decaying deltas — no extra flips
    expect(pager.feed(120, 48)).toBe(0)
    expect(pager.feed(80, 64)).toBe(0)
    expect(pager.feed(30, 80)).toBe(0)
  })

  it('starts a new gesture after a quiet gap', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    pager.feed(20, 16)
    expect(pager.feed(40, 32)).toBe(1)
    pager.feed(80, 48) // momentum, swallowed
    expect(pager.feed(15, 500)).toBe(0) // fresh swipe after the tail went quiet
    expect(pager.feed(60, 516)).toBe(1)
  })

  it('a direction reversal flips immediately without waiting for the gap', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    expect(pager.feed(60, 16)).toBe(1)
    pager.feed(90, 32) // momentum, swallowed
    expect(pager.feed(-100, 60)).toBe(-1)
  })

  it('ignores tiny drifts separated by pauses', () => {
    const pager = createWheelPager()
    expect(pager.feed(3, 0)).toBe(0)
    expect(pager.feed(4, 300)).toBe(0)
    expect(pager.feed(3, 600)).toBe(0)
  })

  it('flips again on a second swipe even while the first tail is still running', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    pager.feed(20, 16)
    expect(pager.feed(40, 32)).toBe(1)
    // decaying momentum tail — every gap stays under 200ms, so a pure
    // gap-based reset never fires
    const tail = [
      [100, 48],
      [70, 64],
      [50, 80],
      [35, 96],
      [25, 112],
      [18, 128],
      [12, 144],
      [8, 160],
      [5, 300],
      [4, 440],
    ]
    for (const [d, t] of tail) expect(pager.feed(d, t)).toBe(0)
    // second deliberate swipe 60ms after the last tail event: velocity ramps up
    pager.feed(12, 500)
    pager.feed(28, 516)
    const flips = [pager.feed(45, 532), pager.feed(50, 548)]
    expect(flips).toContain(1)
  })

  it('coalesced momentum events (double delta, double interval) do not re-flip', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    pager.feed(20, 16)
    expect(pager.feed(40, 32)).toBe(1)
    // tail past the refractory window, then one coalesced event: 50px over
    // 32ms is two ~25px frames merged — per-ms velocity stays flat
    for (const [d, t] of [
      [90, 48],
      [80, 64],
      [70, 96],
      [60, 128],
      [45, 160],
      [38, 192],
      [34, 240],
      [28, 256],
      [50, 288],
      [20, 304],
      [15, 320],
    ]) {
      expect(pager.feed(d, t)).toBe(0)
    }
  })

  it('a hard flick whose first event is already notch-sized flips exactly once', () => {
    const pager = createWheelPager()
    // violent swipe (or a coalesced first frame): opens at 70px, then a
    // continuous ~16ms stream with several more notch-sized events
    expect(pager.feed(70, 0)).toBe(1)
    for (const [d, t] of [
      [90, 16],
      [110, 32],
      [80, 48],
      [50, 64],
      [25, 80],
      [12, 96],
      [5, 112],
    ]) {
      expect(pager.feed(d, t)).toBe(0)
    }
  })

  it('a coalesced pulse right after a flip (page-render jank) does not re-flip', () => {
    const pager = createWheelPager()
    pager.feed(10, 0)
    pager.feed(25, 16)
    expect(pager.feed(30, 32)).toBe(1)
    // the flip re-renders the slide; Chromium coalesces the next momentum
    // frames into one notch-sized event spaced past DISCRETE_MS
    expect(pager.feed(90, 90)).toBe(0)
    expect(pager.feed(40, 106)).toBe(0)
    expect(pager.feed(20, 122)).toBe(0)
  })

  it('a single stray rapid event does not stall a steady wheel spin', () => {
    const pager = createWheelPager()
    expect(pager.feed(100, 0)).toBe(1)
    expect(pager.feed(100, 80)).toBe(1)
    expect(pager.feed(100, 160)).toBe(1)
    // isolated jitter event between notches (high-res wheel remnant)
    expect(pager.feed(30, 176)).toBe(0)
    // the spin continues — paging must not stop until a GAP_MS pause
    expect(pager.feed(100, 256)).toBe(1)
    expect(pager.feed(100, 336)).toBe(1)
  })

  it('ignores zero-delta events (pure horizontal scroll)', () => {
    const pager = createWheelPager()
    expect(pager.feed(0, 0)).toBe(0)
    // and they do not break an in-progress accumulation
    pager.feed(30, 16)
    pager.feed(0, 24)
    expect(pager.feed(40, 32)).toBe(1)
  })
})

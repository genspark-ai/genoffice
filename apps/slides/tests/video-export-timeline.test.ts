import { describe, expect, it } from 'vitest'
import {
  ANIM_HOLD_MS,
  STEP_GAP_MS,
  TRANSITION_MS,
  animCursorAt,
  buildTimeline,
  resolveTransition,
  segmentAt,
  totalDurationMs,
} from '../src/renderer/video-export/timeline'
import type { AnimStep } from '../src/renderer/animation-play'

const step = (totalMs: number, auto = false): AnimStep => ({ auto, items: [], totalMs })

const base = { advanceMs: null, narrationMs: 0, steps: [] as AnimStep[] }

describe('buildTimeline', () => {
  it('gives a plain slide the default time and no transition', () => {
    const [seg] = buildTimeline([{ index: 0, transition: 'none', ...base }], {
      defaultSlideMs: 5000,
      useTimings: true,
    })
    expect(seg).toMatchObject({ startMs: 0, durationMs: 5000, transition: null, transitionMs: 0 })
  })

  it('adds the transition in front of the dwell time and chains starts', () => {
    const segs = buildTimeline(
      [
        { index: 0, transition: 'none', ...base },
        { index: 1, transition: 'wipe', ...base },
      ],
      { defaultSlideMs: 3000, useTimings: true },
    )
    expect(segs[1]).toMatchObject({
      startMs: 3000,
      transition: 'wipe',
      transitionMs: TRANSITION_MS.wipe,
      durationMs: 3000 + TRANSITION_MS.wipe,
    })
    expect(totalDurationMs(segs)).toBe(6000 + TRANSITION_MS.wipe)
  })

  it('honors a saved advance time only when timings are used', () => {
    const slide = { index: 0, transition: 'none' as const, ...base, advanceMs: 12000 }
    expect(buildTimeline([slide], { defaultSlideMs: 5000, useTimings: true })[0]!.durationMs).toBe(
      12000,
    )
    expect(buildTimeline([slide], { defaultSlideMs: 5000, useTimings: false })[0]!.durationMs).toBe(
      5000,
    )
  })

  it('keeps the slide up for its narration', () => {
    const slide = { index: 0, transition: 'fade' as const, ...base, narrationMs: 9000 }
    expect(buildTimeline([slide], { defaultSlideMs: 5000, useTimings: true })[0]!.durationMs).toBe(
      9000 + TRANSITION_MS.fade,
    )
  })

  it('plays click steps back to back with a gap and holds the final state', () => {
    const steps = [step(1000), step(500)]
    const [seg] = buildTimeline([{ index: 0, transition: 'none', ...base, steps }], {
      defaultSlideMs: 1000,
      useTimings: true,
    })
    expect(seg!.stepStartsMs).toEqual([STEP_GAP_MS, STEP_GAP_MS + 1000 + STEP_GAP_MS])
    expect(seg!.durationMs).toBe(STEP_GAP_MS + 1000 + STEP_GAP_MS + 500 + ANIM_HOLD_MS)
  })

  it('starts an automatic first step right after the transition', () => {
    const [seg] = buildTimeline(
      [{ index: 0, transition: 'push', ...base, steps: [step(800, true)] }],
      { defaultSlideMs: 5000, useTimings: true },
    )
    expect(seg!.stepStartsMs).toEqual([TRANSITION_MS.push])
  })
})

describe('resolveTransition', () => {
  it('is deterministic for random', () => {
    expect(resolveTransition('random', 3)).toBe(resolveTransition('random', 3))
    expect(resolveTransition('random', 0)).not.toBe('random')
    expect(resolveTransition('none', 0)).toBeNull()
  })
})

describe('segmentAt / animCursorAt', () => {
  const segs = buildTimeline(
    [
      { index: 0, transition: 'none', ...base, steps: [step(1000), step(500)] },
      { index: 1, transition: 'fade', ...base },
    ],
    { defaultSlideMs: 2000, useTimings: true },
  )

  it('finds the segment on screen, the last one past the end', () => {
    expect(segmentAt(segs, 0)!.index).toBe(0)
    expect(segmentAt(segs, segs[1]!.startMs)!.index).toBe(1)
    expect(segmentAt(segs, 1e9)!.index).toBe(1)
  })

  it('walks the animation cursor through rest and play', () => {
    const seg = segs[0]!
    expect(animCursorAt(seg, 0)).toEqual({ played: 0, activeMs: null })
    expect(animCursorAt(seg, STEP_GAP_MS + 250)).toEqual({ played: 0, activeMs: 250 })
    expect(animCursorAt(seg, STEP_GAP_MS + 1000 + 10)).toEqual({ played: 1, activeMs: null })
    expect(animCursorAt(seg, seg.durationMs)).toEqual({ played: 2, activeMs: null })
  })
})

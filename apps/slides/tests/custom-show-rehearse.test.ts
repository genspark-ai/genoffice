/**
 * Pure-logic tests for custom shows + rehearsal timing:
 * - computePlayOrder: hidden-page and out-of-range filtering for default/custom order
 * - Rehearsal time accumulation: page turns/revisits accumulate per-page dwell milliseconds, converted to seconds at the end
 * - Engine advTm patch: write/read/clear the auto-advance time; changing the transition effect doesn't lose advTm
 */
import { describe, expect, it } from 'vitest'
import {
  computePlayOrder,
  finishRehearse,
  formatClock,
  startRehearse,
  switchRehearsePage,
} from '../src/renderer/slideshow-utils'
import {
  patchSlideAdvanceTimeXml,
  patchSlideTransitionXml,
  readSlideAdvanceTimeXml,
} from '@genoffice/pptx-engine'

describe('computePlayOrder (playback order)', () => {
  const slides = [{}, { hidden: true }, {}, {}, { hidden: true }]

  it('default order skips hidden slides', () => {
    expect(computePlayOrder(slides, 0)).toEqual([0, 2, 3])
  })

  it('starting from a hidden slide keeps that slide', () => {
    expect(computePlayOrder(slides, 1)).toEqual([0, 1, 2, 3])
  })

  it('custom show plays in the specified order', () => {
    expect(computePlayOrder(slides, 3, [3, 0, 2])).toEqual([3, 0, 2])
  })

  it('custom show filters out hidden slides and out-of-range indices', () => {
    expect(computePlayOrder(slides, 3, [3, 1, 99, 0])).toEqual([3, 0])
  })

  it('falls back to the starting slide when the result is empty', () => {
    expect(computePlayOrder([{ hidden: true }], 0, [99])).toEqual([0])
  })

  it('returns no order for an out-of-range start slide', () => {
    expect(computePlayOrder([], 0)).toEqual([])
    expect(computePlayOrder([{}, {}], 5)).toEqual([])
    expect(computePlayOrder([{}, {}], -1)).toEqual([])
    expect(computePlayOrder([{}, {}], 1)).toEqual([0, 1])
  })
})

describe('rehearsal timer accumulation', () => {
  it('page turns accumulate the previous slide dwell time; revisiting keeps accumulating', () => {
    let t = startRehearse(3, 0, 1000)
    t = switchRehearsePage(t, 1, 3500) // page 0 dwelled 2.5s
    expect(t.perPageMs).toEqual([2500, 0, 0])
    t = switchRehearsePage(t, 0, 4000) // page 1 dwelled 0.5s, back to page 0
    expect(t.perPageMs).toEqual([2500, 500, 0])
    // End: page 0 accumulates 2s more -> 4.5s rounds to 5s; page 1's 0.5s records at least 1s; unvisited page 2 is 0
    expect(finishRehearse(t, 6000)).toEqual([5, 1, 0])
  })

  it('clock going backwards never produces negatives', () => {
    const t = startRehearse(1, 0, 5000)
    expect(finishRehearse(t, 4000)).toEqual([0])
  })

  it('formatClock outputs m:ss', () => {
    expect(formatClock(0)).toBe('0:00')
    expect(formatClock(65_000)).toBe('1:05')
    expect(formatClock(600_000)).toBe('10:00')
  })
})

describe('advTm (auto-advance time) pptx patch', () => {
  const SUFFIX = '</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping/></p:clrMapOvr></p:sld>'

  it('with no transition effect, creates an empty transition with only advTm, readable back', () => {
    const out = patchSlideAdvanceTimeXml(SUFFIX, 3000)
    expect(out).toContain('</p:clrMapOvr><p:transition advTm="3000"/>')
    expect(readSlideAdvanceTimeXml(out)).toBe(3000)
  })

  it('with an existing transition, adds advTm on its opening tag; ms=null removes it', () => {
    const withFade = patchSlideTransitionXml(SUFFIX, 'fade')
    const out = patchSlideAdvanceTimeXml(withFade, 5000)
    expect(out).toContain('<p:transition advTm="5000"><p:fade/></p:transition>')
    const cleared = patchSlideAdvanceTimeXml(out, null)
    expect(readSlideAdvanceTimeXml(cleared)).toBeNull()
    expect(cleared).toContain('<p:transition><p:fade/></p:transition>')
  })

  it('rewriting overwrites the old value instead of stacking attributes', () => {
    const out = patchSlideAdvanceTimeXml(patchSlideAdvanceTimeXml(SUFFIX, 3000), 8000)
    expect(out.match(/advTm=/g)?.length).toBe(1)
    expect(readSlideAdvanceTimeXml(out)).toBe(8000)
  })

  it('changing the transition keeps advTm; setting it to none keeps it too', () => {
    const timed = patchSlideAdvanceTimeXml(patchSlideTransitionXml(SUFFIX, 'fade'), 4000)
    const pushed = patchSlideTransitionXml(timed, 'push')
    expect(pushed).toContain('<p:push dir="u"/>')
    expect(readSlideAdvanceTimeXml(pushed)).toBe(4000)
    const none = patchSlideTransitionXml(timed, 'none')
    expect(none).toContain('<p:transition advTm="4000"/>')
  })

  it('writes advTm to both Choice/Fallback transitions inside morph AlternateContent', () => {
    const morph = patchSlideTransitionXml(SUFFIX, 'morph')
    const out = patchSlideAdvanceTimeXml(morph, 6000)
    expect(out.match(/advTm="6000"/g)?.length).toBe(2)
    expect(readSlideAdvanceTimeXml(out)).toBe(6000)
  })
})

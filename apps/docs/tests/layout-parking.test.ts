/**
 * Layout parking (route-B prototype, #2781): the gate that keeps whole-
 * document style recalc/layout/paint off the typing path. The decisions
 * must be exact from both sides — parking a visible block hides live text,
 * failing to park an offscreen block keeps the O(document) cost the
 * prototype exists to remove.
 */
import { describe, expect, it } from 'vitest'
import {
  applyParking,
  PARKED_CLASS,
  PARK_SUPPRESSED_CLASS,
  parkBand,
  parkDecisions,
  parkingSuppressed,
  resumeParking,
  suppressParking,
} from '../src/renderer/editor/layout-parking'

const band = { top: 0, bottom: 1000 }

describe('parkDecisions', () => {
  it('parks everything outside the band', () => {
    expect(
      parkDecisions(
        [
          { top: -900, height: 400 },
          { top: -50, height: 300 },
          { top: 200, height: 400 },
          { top: 1500, height: 300 },
        ],
        band,
      ),
    ).toEqual([true, false, false, true])
  })

  it('keeps a box straddling the band edge unparked', () => {
    expect(parkDecisions([{ top: 990, height: 40 }], band)).toEqual([false])
    expect(parkDecisions([{ top: -20, height: 40 }], band)).toEqual([false])
  })

  it('parks a box that ends exactly at the band top', () => {
    expect(parkDecisions([{ top: -400, height: 400 }], band)).toEqual([true])
  })

  it('answers nothing for no boxes', () => {
    expect(parkDecisions([], band)).toEqual([])
  })
})

describe('parkBand', () => {
  it('grows the scroller viewport by one viewport of headroom each way', () => {
    const scroller = document.createElement('div')
    scroller.getBoundingClientRect = () => ({ top: 100, bottom: 600, height: 500 }) as DOMRect
    expect(parkBand(scroller)).toEqual({ top: -400, bottom: 1100 })
  })
})

describe('applyParking', () => {
  it('parks offscreen blocks with their remembered height and reveals the rest', () => {
    const parked = document.createElement('p')
    const shown = document.createElement('p')
    applyParking(
      [parked, shown],
      [
        { top: -900, height: 240 },
        { top: 10, height: 60 },
      ],
      [true, false],
    )
    expect(parked.classList.contains(PARKED_CLASS)).toBe(true)
    expect(parked.style.getPropertyValue('--park-h')).toBe('240px')
    expect(shown.classList.contains(PARKED_CLASS)).toBe(false)
  })

  it('reveals a block that scrolled back into the band and drops the stamp', () => {
    const el = document.createElement('p')
    applyParking([el], [{ top: -900, height: 240 }], [true])
    expect(el.classList.contains(PARKED_CLASS)).toBe(true)
    applyParking([el], [{ top: 10, height: 240 }], [false])
    expect(el.classList.contains(PARKED_CLASS)).toBe(false)
    expect(el.style.getPropertyValue('--park-h')).toBe('')
  })

  it('re-parking an already parked block keeps the class without restamping', () => {
    const el = document.createElement('p')
    el.style.setProperty('--park-h', '180px')
    el.classList.add(PARKED_CLASS)
    applyParking([el], [{ top: -900, height: 999 }], [true])
    expect(el.classList.contains(PARKED_CLASS)).toBe(true)
    expect(el.style.getPropertyValue('--park-h')).toBe('180px')
  })

  it('skips null elements instead of throwing', () => {
    expect(() =>
      applyParking([null, document.createElement('p')], [{ top: 0, height: 10 }], [true, true]),
    ).not.toThrow()
  })
})

describe('suppression: the pass lifecycle', () => {
  it('toggles the pm-level class and reports it', () => {
    const pm = document.createElement('div')
    expect(parkingSuppressed(pm)).toBe(false)
    suppressParking(pm)
    expect(parkingSuppressed(pm)).toBe(true)
    resumeParking(pm)
    expect(parkingSuppressed(pm)).toBe(false)
    expect(pm.classList.contains(PARK_SUPPRESSED_CLASS)).toBe(false)
  })
})

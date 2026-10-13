import { describe, expect, it } from 'vitest'

import {
  DOCK_BAND_SLACK,
  TAB_STRIP_HEIGHT,
  TEAR_OFF_SLACK,
  dockBand,
  insertionIndexForX,
  isBeyondBand,
  pointInRect,
  tabStripHeight,
} from '../src/shared/tab-drag-geometry'

/**
 * Pure geometry behind tearing a tab off the strip and docking a detached
 * window back (src/shared/tab-drag-geometry.ts): shared by the strip renderer
 * and the main process, so both sides agree on where the thresholds are.
 */

describe('isBeyondBand', () => {
  const top = 0
  const bottom = TAB_STRIP_HEIGHT

  it('stays inside while the pointer wanders within the slack', () => {
    expect(isBeyondBand(20, top, bottom, TEAR_OFF_SLACK)).toBe(false)
    expect(isBeyondBand(bottom + TEAR_OFF_SLACK, top, bottom, TEAR_OFF_SLACK)).toBe(false)
    expect(isBeyondBand(top - TEAR_OFF_SLACK, top, bottom, TEAR_OFF_SLACK)).toBe(false)
  })

  it('tears off one pixel past the slack, above or below', () => {
    expect(isBeyondBand(bottom + TEAR_OFF_SLACK + 1, top, bottom, TEAR_OFF_SLACK)).toBe(true)
    expect(isBeyondBand(top - TEAR_OFF_SLACK - 1, top, bottom, TEAR_OFF_SLACK)).toBe(true)
  })

  it('with zero slack the band edges themselves are still inside', () => {
    expect(isBeyondBand(top, top, bottom, 0)).toBe(false)
    expect(isBeyondBand(bottom, top, bottom, 0)).toBe(false)
    expect(isBeyondBand(bottom + 1, top, bottom, 0)).toBe(true)
  })
})

describe('dockBand', () => {
  it('is the strip at the top of the shell content, grown by the slack', () => {
    const band = dockBand({ x: 100, y: 200, width: 1000, height: 700 })
    expect(band).toEqual({
      x: 100,
      y: 200 - DOCK_BAND_SLACK,
      width: 1000,
      height: TAB_STRIP_HEIGHT + 2 * DOCK_BAND_SLACK,
    })
  })

  it('accepts a custom slack', () => {
    const band = dockBand({ x: 0, y: 0, width: 10, height: 10 }, 0)
    expect(band).toEqual({ x: 0, y: 0, width: 10, height: TAB_STRIP_HEIGHT })
  })
})

describe('pointInRect', () => {
  const rect = { x: 10, y: 20, width: 100, height: 40 }

  it('includes the top-left edge and excludes the bottom-right edge', () => {
    expect(pointInRect(10, 20, rect)).toBe(true)
    expect(pointInRect(109, 59, rect)).toBe(true)
    expect(pointInRect(110, 30, rect)).toBe(false)
    expect(pointInRect(50, 60, rect)).toBe(false)
  })

  it('rejects points outside on any side', () => {
    expect(pointInRect(9, 30, rect)).toBe(false)
    expect(pointInRect(50, 19, rect)).toBe(false)
  })
})

describe('insertionIndexForX', () => {
  // Home (pinned) + three document tabs, 100px each
  const rects = [
    { left: 0, width: 60 },
    { left: 60, width: 100 },
    { left: 160, width: 100 },
    { left: 260, width: 100 },
  ]

  it('never yields slot 0: left of everything still lands after Home', () => {
    expect(insertionIndexForX(rects, -50)).toBe(1)
    expect(insertionIndexForX(rects, 30)).toBe(1)
  })

  it('drops before the tab whose midpoint the pointer has not passed', () => {
    expect(insertionIndexForX(rects, 100)).toBe(1)
    expect(insertionIndexForX(rects, 110)).toBe(2)
    expect(insertionIndexForX(rects, 209)).toBe(2)
    expect(insertionIndexForX(rects, 210)).toBe(3)
  })

  it('appends past the last midpoint', () => {
    expect(insertionIndexForX(rects, 311)).toBe(4)
    expect(insertionIndexForX(rects, 10_000)).toBe(4)
  })

  it('appends at slot 1 when only Home is open', () => {
    expect(insertionIndexForX([{ left: 0, width: 60 }], 500)).toBe(1)
    expect(insertionIndexForX([], 500)).toBe(1)
  })
})

describe('tabStripHeight', () => {
  it('is the unscaled constant by default', () => {
    expect(tabStripHeight()).toBe(TAB_STRIP_HEIGHT)
    expect(tabStripHeight(1)).toBe(TAB_STRIP_HEIGHT)
  })

  it('grows with the interface scale', () => {
    // The strip is `zoom`ed by `--ui-scale` (apps/shell/src/main/ui-scale.ts),
    // and the main process positions every editor view at this number. Using
    // the unscaled 40 would leave a 150% strip painting 20px down into the
    // document — the document under a tab strip, with nothing on screen saying
    // why.
    expect(tabStripHeight(1.5)).toBe(60)
    expect(tabStripHeight(1.25)).toBe(50)
  })

  it('rounds to whole pixels', () => {
    // The strip is a CSS box: a fractional height makes the view bounds below
    // it land on a half pixel, and the seam shows.
    expect(Number.isInteger(tabStripHeight(1.1))).toBe(true)
    expect(tabStripHeight(1.1)).toBe(44)
  })

  it('grows the dock band with it', () => {
    // The band is what a natively dragged window hovers over to dock back, so
    // it has to be the strip's *rendered* height, not its layout height.
    const bounds = { x: 0, y: 100, width: 800, height: 600 }
    expect(dockBand(bounds).height).toBe(TAB_STRIP_HEIGHT + 2 * DOCK_BAND_SLACK)
    expect(dockBand(bounds, DOCK_BAND_SLACK, 1.5).height).toBe(60 + 2 * DOCK_BAND_SLACK)
  })
})

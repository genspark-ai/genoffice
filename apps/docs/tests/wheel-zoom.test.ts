// Ctrl+wheel in Docs: a mouse notch moves ten percentage points (10% + one
// Windows notch is 20%, not the 500% cap), a trackpad pinch stays continuous.
import { describe, expect, it } from 'vitest'
import {
  DOCS_ZOOM_MAX,
  DOCS_ZOOM_MIN,
  nextDocsZoom,
  stepDocsZoom,
} from '../src/renderer/wheel-zoom'

describe('nextDocsZoom', () => {
  it('steps ten points per notch from 10%', () => {
    expect(nextDocsZoom(10, 'zoom-in', -100)).toBe(20)
    expect(nextDocsZoom(60, 'zoom-in', -100)).toBe(70)
    expect(nextDocsZoom(60, 'zoom-out', 100)).toBe(50)
  })

  it('snaps a pinch-landed fraction to whole percents and clamps to 10-500', () => {
    expect(nextDocsZoom(64.37, 'zoom-in', -100)).toBe(74)
    expect(nextDocsZoom(DOCS_ZOOM_MIN, 'zoom-out', 100)).toBe(DOCS_ZOOM_MIN)
    expect(nextDocsZoom(495, 'zoom-in', -100)).toBe(DOCS_ZOOM_MAX)
    expect(nextDocsZoom(15, 'zoom-out', 100)).toBe(DOCS_ZOOM_MIN)
  })

  it('keeps a pinch continuous with the delta-scaled factor', () => {
    expect(nextDocsZoom(100, 'pinch', -3.4)).toBeCloseTo(102.04, 10)
    expect(nextDocsZoom(100, 'pinch', 7.25)).toBeCloseTo(95.65, 10)
    expect(nextDocsZoom(499, 'pinch', -50)).toBe(DOCS_ZOOM_MAX)
  })
})

describe('stepDocsZoom', () => {
  it('moves ten whole points and stops at the Word range ends', () => {
    expect(stepDocsZoom(100, 1)).toBe(110)
    expect(stepDocsZoom(104.6, -1)).toBe(95)
    expect(stepDocsZoom(495, 1)).toBe(DOCS_ZOOM_MAX)
    expect(stepDocsZoom(15, -1)).toBe(DOCS_ZOOM_MIN)
  })
})

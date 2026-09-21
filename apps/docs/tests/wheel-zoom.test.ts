// Ctrl+wheel in Docs: a mouse notch moves ten percentage points (50% + one
// Windows notch is 60%, not the 200% cap), a trackpad pinch stays continuous.
import { describe, expect, it } from 'vitest'
import { DOCS_ZOOM_MAX, DOCS_ZOOM_MIN, nextDocsZoom } from '../src/renderer/wheel-zoom'

describe('nextDocsZoom', () => {
  it('steps ten points per notch from 50%', () => {
    expect(nextDocsZoom(50, 'zoom-in', -100)).toBe(60)
    expect(nextDocsZoom(60, 'zoom-in', -100)).toBe(70)
    expect(nextDocsZoom(60, 'zoom-out', 100)).toBe(50)
  })

  it('snaps a pinch-landed fraction to whole percents and clamps to 50-200', () => {
    expect(nextDocsZoom(64.37, 'zoom-in', -100)).toBe(74)
    expect(nextDocsZoom(DOCS_ZOOM_MIN, 'zoom-out', 100)).toBe(DOCS_ZOOM_MIN)
    expect(nextDocsZoom(195, 'zoom-in', -100)).toBe(DOCS_ZOOM_MAX)
  })

  it('keeps a pinch continuous with the delta-scaled factor', () => {
    expect(nextDocsZoom(100, 'pinch', -3.4)).toBeCloseTo(102.04, 10)
    expect(nextDocsZoom(100, 'pinch', 7.25)).toBeCloseTo(95.65, 10)
    expect(nextDocsZoom(199, 'pinch', -50)).toBe(DOCS_ZOOM_MAX)
  })
})

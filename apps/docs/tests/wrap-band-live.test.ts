import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { textboxBandBottom, textboxBandTop } from '../src/renderer/editor/extensions'

describe('wrapTopAndBottom band vs live box height', () => {
  const box: TextboxDisplay = { paras: [], bandTopPx: 20, bandBottomPx: 70, heightPx: 50 }

  it('matches the parse-time bottom at the parse-time height', () => {
    expect(textboxBandBottom(box)).toBe(70)
  })

  it('follows a changed height (autogrow / committed heightPx)', () => {
    expect(textboxBandBottom(box, 120)).toBe(140)
    expect(textboxBandBottom({ ...box, heightPx: 120 })).toBe(140)
  })

  it('keeps the parse-time bottom for boxes without a known height', () => {
    const autoFit: TextboxDisplay = { paras: [], bandBottomPx: 90 }
    expect(textboxBandBottom(autoFit)).toBe(90)
    expect(textboxBandBottom(autoFit, undefined)).toBe(90)
  })

  it('sizes the band of a quarter-turned photo by its swapped extent', () => {
    const turned = { paras: [], bandTopPx: 118, widthPx: 365, heightPx: 274, rotDeg: 90 }
    // turned about its centre: 45.5px above and below the 274px layout box
    expect(textboxBandBottom(turned)).toBeCloseTo(118 + 274 + 45.5, 5)
    expect(textboxBandTop(turned)).toBeCloseTo(118 - 45.5, 5)
    // autogrow re-derives both edges from the live height
    expect(textboxBandTop(turned, 300)).toBeCloseTo(118 - 32.5, 5)
    expect(textboxBandBottom(turned, 300)).toBeCloseTo(118 + 300 + 32.5, 5)
    expect(textboxBandBottom({ ...turned, rotDeg: 180 })).toBe(118 + 274)
  })
})

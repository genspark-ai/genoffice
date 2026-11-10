// Zoom stepping mirrors PowerPoint: +/- walk the preset ladder and a Ctrl+wheel
// notch moves ten percentage points within the 10-400% range.
import { describe, expect, it } from 'vitest'
import {
  ZOOM_MAX,
  ZOOM_MIN,
  ZOOM_PRESETS,
  clampZoom,
  nextPreset,
  notchStep,
  prevPreset,
} from '../src/renderer/zoom-steps'

describe('zoom presets', () => {
  it('lists the PowerPoint ladder from 10% to 400%', () => {
    expect(ZOOM_PRESETS.map((p) => Math.round(p * 100))).toEqual([
      10, 25, 33, 50, 66, 75, 100, 150, 200, 300, 400,
    ])
    expect(ZOOM_MIN).toBe(0.1)
    expect(ZOOM_MAX).toBe(4)
  })

  it('steps from a preset to its neighbours', () => {
    expect(nextPreset(1)).toBe(1.5)
    expect(prevPreset(1)).toBe(0.75)
    expect(nextPreset(0.5)).toBe(0.66)
    expect(prevPreset(0.66)).toBe(0.5)
  })

  it('steps from an in-between value to the next/previous preset', () => {
    expect(nextPreset(0.6)).toBe(0.66)
    expect(prevPreset(0.6)).toBe(0.5)
    expect(nextPreset(1.234)).toBe(1.5)
    expect(prevPreset(1.234)).toBe(1)
  })

  it('treats float noise around a preset as that preset', () => {
    expect(nextPreset(0.66 + 1e-9)).toBe(0.75)
    expect(prevPreset(0.66 - 1e-9)).toBe(0.5)
  })

  it('stays put at the ends of the ladder', () => {
    expect(nextPreset(4)).toBe(4)
    expect(nextPreset(3.9)).toBe(4)
    expect(prevPreset(0.1)).toBe(0.1)
    expect(prevPreset(0.12)).toBe(0.1)
  })
})

describe('notchStep', () => {
  it('moves exactly ten percentage points per notch: 50% + one notch is 60%, not 300%', () => {
    expect(notchStep(0.5, 1)).toBeCloseTo(0.6, 10)
    expect(notchStep(0.6, 1)).toBeCloseTo(0.7, 10)
    expect(notchStep(0.5, -1)).toBeCloseTo(0.4, 10)
  })

  it('rounds to whole percents so repeated notches do not drift', () => {
    let z = 0.6437
    for (let i = 0; i < 5; i++) z = notchStep(z, 1)
    expect(z).toBe(1.14)
  })

  it('clamps to the 10-400% range', () => {
    expect(notchStep(0.15, -1)).toBe(0.1)
    expect(notchStep(3.95, 1)).toBe(4)
    expect(clampZoom(0)).toBe(0.1)
    expect(clampZoom(9)).toBe(4)
  })
})

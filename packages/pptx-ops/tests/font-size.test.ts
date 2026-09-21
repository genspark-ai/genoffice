/**
 * PowerPoint's font-size ladder (grow/shrink font) and the ±1 pt nudge behind
 * ⌘] / ⌘[: in-between sizes snap to the next rung, above 96 pt the ladder keeps
 * its last gap of 8, below 8 pt both paths move one point at a time down to 1.
 */
import { describe, it, expect } from 'vitest'
import {
  FONT_SIZES,
  FONT_SIZE_PT_MAX,
  applyFontSizeStep,
  nudgeFontSizePt,
  stepFontSizePt,
} from '../src/font-size'

describe('stepFontSizePt', () => {
  it('walks the ladder rung by rung in both directions', () => {
    for (let i = 0; i < FONT_SIZES.length - 1; i++) {
      expect(stepFontSizePt(FONT_SIZES[i]!, 1)).toBe(FONT_SIZES[i + 1])
      expect(stepFontSizePt(FONT_SIZES[i + 1]!, -1)).toBe(FONT_SIZES[i])
    }
  })

  it('snaps a size between two rungs to the neighbouring rung', () => {
    expect(stepFontSizePt(13, 1)).toBe(14)
    expect(stepFontSizePt(13, -1)).toBe(12)
    expect(stepFontSizePt(10.2, 1)).toBe(10.5)
    expect(stepFontSizePt(10.2, -1)).toBe(10)
    expect(stepFontSizePt(45, 1)).toBe(48)
    expect(stepFontSizePt(45, -1)).toBe(44)
  })

  it('steps by 8 above the ladder and lands back on 96 when shrinking', () => {
    expect(stepFontSizePt(96, 1)).toBe(104)
    expect(stepFontSizePt(104, 1)).toBe(112)
    expect(stepFontSizePt(100, 1)).toBe(108)
    expect(stepFontSizePt(104, -1)).toBe(96)
    expect(stepFontSizePt(100, -1)).toBe(96)
    expect(stepFontSizePt(FONT_SIZE_PT_MAX, 1)).toBe(FONT_SIZE_PT_MAX)
  })

  it('moves one point at a time below the ladder, never under 1 pt', () => {
    expect(stepFontSizePt(8, -1)).toBe(7)
    expect(stepFontSizePt(2, -1)).toBe(1)
    expect(stepFontSizePt(1, -1)).toBe(1)
    expect(stepFontSizePt(6, 1)).toBe(7)
    expect(stepFontSizePt(7.5, 1)).toBe(8)
  })
})

describe('nudgeFontSizePt', () => {
  it('adds or removes one point and clamps to the schema range', () => {
    expect(nudgeFontSizePt(18, 1)).toBe(19)
    expect(nudgeFontSizePt(10.5, -1)).toBe(9.5)
    expect(nudgeFontSizePt(1, -1)).toBe(1)
    expect(nudgeFontSizePt(1.5, -1)).toBe(1)
    expect(nudgeFontSizePt(FONT_SIZE_PT_MAX, 1)).toBe(FONT_SIZE_PT_MAX)
  })
})

describe('applyFontSizeStep', () => {
  it('dispatches on the step mode', () => {
    expect(applyFontSizeStep(18, { dir: 1, mode: 'ladder' })).toBe(20)
    expect(applyFontSizeStep(18, { dir: 1, mode: 'point' })).toBe(19)
    expect(applyFontSizeStep(18, { dir: -1, mode: 'ladder' })).toBe(16)
    expect(applyFontSizeStep(18, { dir: -1, mode: 'point' })).toBe(17)
  })
})

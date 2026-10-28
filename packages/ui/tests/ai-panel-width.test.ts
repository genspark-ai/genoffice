import { describe, expect, it } from 'vitest'
import {
  AI_PANEL_WIDTH_DEFAULT,
  AI_PANEL_WIDTH_MAX,
  aiPanelAutoWidth,
  aiPanelClampWidth,
  aiPanelWidthCap,
} from '../src/ai-panel-width'

describe('aiPanelWidthCap', () => {
  it('tracks 60% of the window, rounded, up to the 720px ceiling', () => {
    expect(aiPanelWidthCap(1000)).toBe(600)
    expect(aiPanelWidthCap(900)).toBe(540)
    expect(aiPanelWidthCap(913)).toBe(548) // round(547.8)
    expect(aiPanelWidthCap(1234)).toBe(720) // round(740.4) exceeds the ceiling
    expect(aiPanelWidthCap(1500)).toBe(720)
    expect(aiPanelWidthCap(4000)).toBe(720)
  })

  it('never drops below the default width, even for tiny windows', () => {
    expect(aiPanelWidthCap(100)).toBe(AI_PANEL_WIDTH_DEFAULT)
    expect(aiPanelWidthCap(0)).toBe(AI_PANEL_WIDTH_DEFAULT)
    // the shell's minimum window already clears the floor: 60% of 720 is 432
    expect(aiPanelWidthCap(720)).toBe(432)
  })
})

describe('aiPanelAutoWidth', () => {
  it('floors at the default for narrow windows', () => {
    expect(aiPanelAutoWidth(1280)).toBe(AI_PANEL_WIDTH_DEFAULT)
    expect(aiPanelAutoWidth(1000)).toBe(AI_PANEL_WIDTH_DEFAULT)
  })

  it('grows with the window at roughly 22%', () => {
    expect(aiPanelAutoWidth(1920)).toBe(422) // round(1920 * 0.22)
    expect(aiPanelAutoWidth(2560)).toBe(563) // round(2560 * 0.22)
  })

  it('stops at the 720px ceiling for very wide windows', () => {
    expect(aiPanelAutoWidth(3840)).toBe(AI_PANEL_WIDTH_MAX)
    expect(aiPanelAutoWidth(5120)).toBe(AI_PANEL_WIDTH_MAX)
  })

  it('stays between the default and the cap across a sweep of window sizes', () => {
    for (let w = 0; w <= 4000; w += 137) {
      const cap = aiPanelWidthCap(w)
      expect(aiPanelAutoWidth(w)).toBeGreaterThanOrEqual(AI_PANEL_WIDTH_DEFAULT)
      expect(aiPanelAutoWidth(w)).toBeLessThanOrEqual(cap)
      expect(cap).toBeLessThanOrEqual(AI_PANEL_WIDTH_MAX)
    }
  })
})

describe('aiPanelClampWidth', () => {
  it('lifts dragged widths below the default up to it', () => {
    expect(aiPanelClampWidth(200, 1920)).toBe(AI_PANEL_WIDTH_DEFAULT)
    expect(aiPanelClampWidth(0, 1920)).toBe(AI_PANEL_WIDTH_DEFAULT)
  })

  it('lowers dragged widths above the cap down to it', () => {
    expect(aiPanelClampWidth(800, 1920)).toBe(720)
    expect(aiPanelClampWidth(900, 1000)).toBe(600) // cap is 60% of 1000
  })

  it('keeps in-range dragged widths untouched', () => {
    expect(aiPanelClampWidth(450, 1280)).toBe(450)
    expect(aiPanelClampWidth(600, 1920)).toBe(600)
  })
})

describe('invalid window widths', () => {
  it('fall back to default-based results for NaN, negative and infinite input', () => {
    for (const bad of [Number.NaN, -400, Number.POSITIVE_INFINITY]) {
      expect(aiPanelWidthCap(bad)).toBe(AI_PANEL_WIDTH_DEFAULT)
      expect(aiPanelAutoWidth(bad)).toBe(AI_PANEL_WIDTH_DEFAULT)
      expect(aiPanelClampWidth(500, bad)).toBe(AI_PANEL_WIDTH_DEFAULT)
    }
  })
})

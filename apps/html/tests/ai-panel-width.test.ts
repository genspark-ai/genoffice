import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { aiPanelAutoWidth, aiPanelClampWidth, AI_PANEL_WIDTH_DEFAULT } from '@genoffice/ui'
import {
  displayPanelWidth,
  loadSavedPanelWidth,
  PANEL_WIDTH_KEY,
} from '../src/renderer/ai/panel-width'

beforeEach(() => localStorage.clear())
afterEach(() => localStorage.clear())

describe('loadSavedPanelWidth', () => {
  it('keeps the app localStorage key', () => {
    expect(PANEL_WIDTH_KEY).toBe('html-ai-panel-width')
  })
  it('returns null when nothing was saved or the value is unusable', () => {
    expect(loadSavedPanelWidth()).toBeNull()
    localStorage.setItem(PANEL_WIDTH_KEY, 'nonsense')
    expect(loadSavedPanelWidth()).toBeNull()
    localStorage.setItem(PANEL_WIDTH_KEY, '0')
    expect(loadSavedPanelWidth()).toBeNull()
  })
  it('returns the raw saved value: window clamping happens at display time', () => {
    localStorage.setItem(PANEL_WIDTH_KEY, '300')
    expect(loadSavedPanelWidth()).toBe(300)
    localStorage.setItem(PANEL_WIDTH_KEY, '640')
    expect(loadSavedPanelWidth()).toBe(640)
  })
})

describe('displayPanelWidth', () => {
  it('no dragged preference: shows the auto width, re-derived per window size', () => {
    expect(displayPanelWidth(null, 1024)).toBe(aiPanelAutoWidth(1024))
    expect(displayPanelWidth(null, 1920)).toBe(aiPanelAutoWidth(1920))
    // 22% of a wide window, floored at the default
    expect(displayPanelWidth(null, 1920)).toBe(
      Math.max(AI_PANEL_WIDTH_DEFAULT, Math.round(1920 * 0.22)),
    )
    expect(displayPanelWidth(null, 1024)).toBe(AI_PANEL_WIDTH_DEFAULT)
  })
  it('a dragged preference is clamped into the window, never below 360', () => {
    expect(displayPanelWidth(420, 1920)).toBe(420)
    // legacy value below the floor migrates up
    expect(displayPanelWidth(300, 1920)).toBe(360)
    expect(displayPanelWidth(40, 800)).toBe(360)
    // above the 60% cap it caps
    expect(displayPanelWidth(700, 800)).toBe(aiPanelClampWidth(700, 800))
    expect(displayPanelWidth(700, 800)).toBeLessThanOrEqual(480)
    // transiently tiny viewport: floor still satisfiable
    expect(displayPanelWidth(400, 0)).toBe(360)
  })
})

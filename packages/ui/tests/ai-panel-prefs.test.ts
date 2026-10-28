import { describe, expect, it } from 'vitest'
import {
  DEFAULT_AI_PANEL_PREFS,
  normalizeAiPanelPrefs,
  sameAiPanelPrefs,
} from '../src/ai-panel-prefs'

const CHAT_TOGGLES = ['showStreaming', 'showThinking', 'showToolCalls'] as const

describe('AI panel chat display toggles', () => {
  it('default to true', () => {
    for (const key of CHAT_TOGGLES) expect(DEFAULT_AI_PANEL_PREFS[key]).toBe(true)
  })

  it('fall back to true for empty or garbage input', () => {
    const garbage: unknown[] = [
      undefined,
      null,
      'large',
      7,
      {},
      { showStreaming: 'yes', showThinking: 1, showToolCalls: null },
    ]
    for (const raw of garbage) {
      const prefs = normalizeAiPanelPrefs(raw)
      for (const key of CHAT_TOGGLES) expect(prefs[key]).toBe(true)
    }
    expect(normalizeAiPanelPrefs(undefined)).toEqual(DEFAULT_AI_PANEL_PREFS)
  })

  it('keep an explicit false through normalization', () => {
    expect(normalizeAiPanelPrefs({ showStreaming: false }).showStreaming).toBe(false)
    expect(normalizeAiPanelPrefs({ showThinking: false }).showThinking).toBe(false)
    expect(normalizeAiPanelPrefs({ showToolCalls: false }).showToolCalls).toBe(false)
    // turning one off leaves the others at their defaults
    const oneOff = normalizeAiPanelPrefs({ showToolCalls: false })
    expect(oneOff.showStreaming).toBe(true)
    expect(oneOff.showThinking).toBe(true)
  })

  it('sameAiPanelPrefs detects a flip of any one toggle', () => {
    const base = normalizeAiPanelPrefs({})
    const flips = [
      { ...base, showStreaming: false },
      { ...base, showThinking: false },
      { ...base, showToolCalls: false },
    ]
    for (const flipped of flips) expect(sameAiPanelPrefs(base, flipped)).toBe(false)
    for (const flipped of flips) expect(sameAiPanelPrefs(flipped, { ...flipped })).toBe(true)
    expect(sameAiPanelPrefs(base, normalizeAiPanelPrefs({}))).toBe(true)
  })
})

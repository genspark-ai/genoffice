import { describe, expect, it } from 'vitest'
import {
  DEFAULT_AI_PANEL_PREFS,
  aiPanelFontPx,
  aiPanelZoom,
  normalizeAiPanelPrefs,
  sameAiPanelPrefs,
} from '@genoffice/ui/ai-panel-prefs'

describe('normalizeAiPanelPrefs', () => {
  it('falls back to defaults for missing or malformed input', () => {
    expect(normalizeAiPanelPrefs(undefined)).toEqual(DEFAULT_AI_PANEL_PREFS)
    expect(normalizeAiPanelPrefs('large')).toEqual(DEFAULT_AI_PANEL_PREFS)
    expect(
      normalizeAiPanelPrefs({ fontSize: 'huge', customFontSize: 'big', spellcheck: 'no' }),
    ).toEqual(DEFAULT_AI_PANEL_PREFS)
  })

  it('validates each field independently', () => {
    expect(normalizeAiPanelPrefs({ fontSize: 'xlarge' })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      fontSize: 'xlarge',
    })
    expect(normalizeAiPanelPrefs({ spellcheck: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      spellcheck: false,
    })
    expect(normalizeAiPanelPrefs({ showAiSidebar: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      showAiSidebar: false,
    })
    // absent field means on: settings files written before the toggle existed
    // must not disable the sidebar
    expect(normalizeAiPanelPrefs({}).showAiSidebar).toBe(true)
  })

  it('clamps and rounds the custom size, accepting numeric strings from old settings files', () => {
    expect(normalizeAiPanelPrefs({ fontSize: 'custom', customFontSize: 20 }).customFontSize).toBe(
      20,
    )
    expect(normalizeAiPanelPrefs({ customFontSize: '17.6' }).customFontSize).toBe(18)
    expect(normalizeAiPanelPrefs({ customFontSize: 4 }).customFontSize).toBe(10)
    expect(normalizeAiPanelPrefs({ customFontSize: 500 }).customFontSize).toBe(32)
    expect(normalizeAiPanelPrefs({ customFontSize: Infinity }).customFontSize).toBe(14)
  })
})

describe('aiPanelZoom', () => {
  it('scales presets from the 14px body size and custom sizes proportionally', () => {
    expect(aiPanelZoom(DEFAULT_AI_PANEL_PREFS)).toBe(1)
    expect(aiPanelFontPx({ ...DEFAULT_AI_PANEL_PREFS, fontSize: 'large' })).toBe(16)
    expect(aiPanelFontPx({ ...DEFAULT_AI_PANEL_PREFS, fontSize: 'xlarge' })).toBe(18)
    const custom = { ...DEFAULT_AI_PANEL_PREFS, fontSize: 'custom' as const, customFontSize: 21 }
    expect(aiPanelZoom(custom)).toBeCloseTo(1.5)
    expect(aiPanelFontPx(custom)).toBe(21)
  })
})

describe('AI panel side', () => {
  it('keeps existing users on the left and accepts the right side', () => {
    expect(normalizeAiPanelPrefs({}).side).toBe('left')
    expect(normalizeAiPanelPrefs({ side: 'bottom' }).side).toBe('left')
    expect(normalizeAiPanelPrefs({ side: 'right' }).side).toBe('right')
  })

  it('does not discard a side-only preference change', () => {
    const left = normalizeAiPanelPrefs({ side: 'left' })
    const right = normalizeAiPanelPrefs({ side: 'right' })
    expect(sameAiPanelPrefs(left, right)).toBe(false)
    expect(sameAiPanelPrefs(right, normalizeAiPanelPrefs({ side: 'right' }))).toBe(true)
  })
})

describe('showAiAskTrigger preference', () => {
  it('defaults to on for settings files written before the switch existed', () => {
    expect(DEFAULT_AI_PANEL_PREFS.showAiAskTrigger).toBe(true)
    expect(normalizeAiPanelPrefs({}).showAiAskTrigger).toBe(true)
    expect(normalizeAiPanelPrefs(undefined).showAiAskTrigger).toBe(true)
  })

  it('accepts an explicit off and ignores malformed values', () => {
    expect(normalizeAiPanelPrefs({ showAiAskTrigger: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      showAiAskTrigger: false,
    })
    expect(normalizeAiPanelPrefs({ showAiAskTrigger: 'no' }).showAiAskTrigger).toBe(true)
  })

  it('takes part in equality so toggling is persisted and broadcast', () => {
    const on = normalizeAiPanelPrefs({})
    const off = normalizeAiPanelPrefs({ showAiAskTrigger: false })
    expect(sameAiPanelPrefs(on, off)).toBe(false)
    expect(sameAiPanelPrefs(off, normalizeAiPanelPrefs({ showAiAskTrigger: false }))).toBe(true)
    expect(sameAiPanelPrefs(on, normalizeAiPanelPrefs({ side: 'right' }))).toBe(false)
  })
})

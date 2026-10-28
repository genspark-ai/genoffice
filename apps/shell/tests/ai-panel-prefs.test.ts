import { readFileSync } from 'node:fs'
import { join } from 'node:path'
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
    expect(normalizeAiPanelPrefs({ showStreaming: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      showStreaming: false,
    })
    expect(normalizeAiPanelPrefs({ showThinking: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      showThinking: false,
    })
    expect(normalizeAiPanelPrefs({ showToolCalls: false })).toEqual({
      ...DEFAULT_AI_PANEL_PREFS,
      showToolCalls: false,
    })
    // absent field means on: settings files written before the toggle existed
    // must not disable the sidebar
    expect(normalizeAiPanelPrefs({}).showAiSidebar).toBe(true)
    expect(normalizeAiPanelPrefs({}).showStreaming).toBe(true)
    expect(normalizeAiPanelPrefs({}).showThinking).toBe(true)
    expect(normalizeAiPanelPrefs({}).showToolCalls).toBe(true)
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

/**
 * The shell main process owns the persistence: setAiPanelPrefs writes flat keys
 * into userData/app-settings.json and currentAiPanelPrefs reads them back. The
 * main module cannot be imported without starting Electron, so the mapping is
 * pinned from the source the way main-locale-tables.test.ts reads the locale
 * table — a dropped key here means the toggle silently resets on relaunch.
 */
const mainSource = readFileSync(join(__dirname, '../src/main/index.ts'), 'utf8')
const setPrefsBody = mainSource.match(
  /const setAiPanelPrefs = \(patch: unknown\): AiPanelPrefs => \{[\s\S]*?\n {2}\}/,
)
const readPrefsBody = mainSource.match(
  /function currentAiPanelPrefs\(\): AiPanelPrefs \{[\s\S]*?\n\}/,
)

describe('shell main persists the AI panel prefs round-trip', () => {
  const fields = ['showStreaming', 'showThinking', 'showToolCalls'] as const
  const settingsKey = (field: string): string =>
    `aiPanel${field[0]!.toUpperCase()}${field.slice(1)}`

  it('setAiPanelPrefs forwards each new field with the fallback idiom', () => {
    expect(setPrefsBody).not.toBeNull()
    for (const field of fields) {
      expect(setPrefsBody![0]).toContain(
        `${field}: '${field}' in raw ? raw.${field} : prev.${field}`,
      )
    }
  })

  it('the write maps every field to its aiPanel* settings key', () => {
    const write = setPrefsBody![0].match(/writeAppSettings\(APP_SETTINGS_PATH\(\), \{[\s\S]*?\}\)/)
    expect(write).not.toBeNull()
    for (const field of fields) {
      expect(write![0]).toContain(`${settingsKey(field)}: next.${field}`)
    }
  })

  it('currentAiPanelPrefs reads the same keys back', () => {
    expect(readPrefsBody).not.toBeNull()
    for (const field of fields) {
      expect(readPrefsBody![0]).toContain(`${field}: saved.${settingsKey(field)}`)
    }
  })
})

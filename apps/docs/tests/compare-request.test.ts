import { describe, expect, it } from 'vitest'
import {
  ALL_COMPARE_SETTINGS,
  COMPARE_SETTING_KEYS,
  effectiveSettings,
  settingAvailable,
  targetAvailable,
  type CompareSide,
} from '../src/renderer/compare-request'

const current: CompareSide = { kind: 'current' }
const file = { kind: 'file', name: 'other.docx', author: 'Rae Visor' } as CompareSide

describe('Compare Documents dialog: where changes can be shown', () => {
  it('a new document is always available', () => {
    expect(targetAvailable('new', null, null)).toBe(true)
    expect(targetAvailable('new', file, file, false)).toBe(true)
  })

  it('changes go into the original or the revised document only when it is the open one', () => {
    expect(targetAvailable('original', current, file)).toBe(true)
    expect(targetAvailable('revised', current, file)).toBe(false)
    expect(targetAvailable('revised', file, current)).toBe(true)
    expect(targetAvailable('original', file, current)).toBe(false)
    expect(targetAvailable('original', null, file)).toBe(false)
  })

  it('a read-only document takes no changes', () => {
    expect(targetAvailable('original', current, file, false)).toBe(false)
    expect(targetAvailable('revised', file, current, false)).toBe(false)
  })
})

describe('Compare Documents dialog: which settings a target honours', () => {
  it('a new document honours every setting', () => {
    for (const key of COMPARE_SETTING_KEYS) expect(settingAvailable(key, 'new')).toBe(true)
    expect(effectiveSettings(ALL_COMPARE_SETTINGS, 'new')).toEqual(ALL_COMPARE_SETTINGS)
  })

  it('written into the open document, the text, formatting and table settings stay', () => {
    for (const target of ['original', 'revised'] as const) {
      expect(COMPARE_SETTING_KEYS.filter((key) => settingAvailable(key, target))).toEqual([
        'formatting',
        'caseChanges',
        'whiteSpace',
        'tables',
      ])
      expect(effectiveSettings({ ...ALL_COMPARE_SETTINGS, whiteSpace: false }, target)).toEqual({
        moves: false,
        comments: false,
        formatting: true,
        caseChanges: true,
        whiteSpace: false,
        tables: true,
        headersFooters: false,
        notes: false,
        textBoxes: false,
        fields: false,
      })
    }
  })
})

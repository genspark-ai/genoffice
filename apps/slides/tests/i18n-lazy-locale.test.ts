import { describe, expect, it } from 'vitest'
import { loadLocale, translate } from '../src/renderer/i18n/strings'
import { zh as paneZh } from '../src/renderer/i18n/panes/zh'
import { en as paneEn } from '../src/renderer/i18n/panes/en'
import { en as aiEn } from '../src/renderer/i18n/ai/en'
import { ja as aiJa } from '../src/renderer/i18n/ai/ja'
import { ja as paneJa } from '../src/renderer/i18n/panes/ja'
import { ja as ribbonJa } from '../src/renderer/i18n/ribbon/ja'

describe('lazy locale dictionaries', () => {
  it('zh and en are eager in the registry', () => {
    expect(translate('zh', 'paneCancel')).toBe(paneZh.paneCancel)
    expect(translate('en', 'paneCancel')).toBe(paneEn.paneCancel)
  })

  it('an unloaded locale falls back to en until loadLocale() resolves', async () => {
    expect(translate('ja', 'aiInputPlaceholderGen')).toBe(aiEn.aiInputPlaceholderGen)
    await loadLocale('ja')
    expect(translate('ja', 'aiInputPlaceholderGen')).toBe(aiJa.aiInputPlaceholderGen)
    expect(translate('ja', 'paneCancel')).toBe(paneJa.paneCancel)
    expect(translate('ja', 'ribbonAiCreditNote')).toBe(ribbonJa.ribbonAiCreditNote)
  })

  it('loadLocale is memoized and safe to re-await', async () => {
    await loadLocale('ja')
    await loadLocale('ja')
    expect(translate('ja', 'paneCancel')).toBe(paneJa.paneCancel)
  })
})

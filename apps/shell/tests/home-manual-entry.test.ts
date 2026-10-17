/**
 * The manual has one on-screen entry point, and a sync snapshot once removed it
 * together with all 21 translations — leaving `Help ▸ User guide` (F1) as the
 * only way in, which nobody discovers without already knowing the shortcut.
 *
 * These tests hold the button and its label in place. They are deliberately
 * cheap: the app's test environment is `node`, so what can be checked here is
 * that the wiring exists and resolves in every language, not that the button
 * renders.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { createI18n, LANGS } from '@genoffice/i18n'
import { strings } from '../src/renderer/src/strings'

const here = dirname(fileURLToPath(import.meta.url))
const homeSource = readFileSync(join(here, '../src/renderer/src/Home.tsx'), 'utf8')

describe('the manual entry point', () => {
  it('is on the home screen and calls openHelp', () => {
    // The manual is a tab of its own, not a view of the home screen: no
    // `active` state, and it leaves the sidebar's selection alone.
    expect(homeSource).toContain('window.aiOffice.openHelp()')
    const button = homeSource.slice(homeSource.indexOf('openHelp()') - 400)
    expect(button).toContain('nav-item')
    expect(button).not.toMatch(/className=\{`nav-item\$\{/)
  })

  it('is labelled in every locale', () => {
    for (const lang of LANGS) {
      const label = strings[lang].navUserGuide
      expect(typeof label, `navUserGuide missing for ${lang}`).toBe('string')
      expect(label.trim(), `navUserGuide is blank for ${lang}`).not.toBe('')
    }
  })

  it('translates the label rather than falling back to English', () => {
    // 21 locales, one label each: a snapshot that takes the button but not the
    // translations would leave the button reading "User Guide" everywhere, and
    // one that takes the translations but not the button leaves strings nothing
    // references. Counting catches both.
    const withLabel = LANGS.filter((lang) => strings[lang].navUserGuide !== undefined)
    expect(withLabel).toHaveLength(LANGS.length)

    const translate = createI18n(strings)
    expect(translate('en', 'navUserGuide')).toBe('User Guide')
    expect(translate('ja', 'navUserGuide')).toBe('ユーザーガイド')
  })
})

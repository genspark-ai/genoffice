import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * The status-bar pair is only visible if the app keeps the state fresh and
 * falls back when the caret collapses. `selection-word-count.test.ts` covers
 * the counting rule; these assertions pin the wiring, because an app that
 * never subscribes to `selectionUpdate` (or that always renders the pair)
 * keeps every counting test green while showing nothing useful.
 */

const SRC = join(__dirname, '../src/renderer')
const appSource = (): string => readFileSync(join(SRC, 'App.tsx'), 'utf8')

describe('status-bar selection word count wiring', () => {
  it('subscribes to selectionUpdate and unsubscribes on teardown', () => {
    const src = appSource()
    expect(src).toContain("editor.on('selectionUpdate', refreshSelection)")
    expect(src).toContain("editor.off('selectionUpdate', refreshSelection)")
  })

  it('shows the selected/total pair only while a selection exists', () => {
    const src = appSource()
    expect(src).toContain('appWordCountSelection')
    // the guard: a collapsed caret must fall back to the plain total. Both
    // branches have to sit inside the same expression, so compare positions
    // rather than match one regex against prettier's line breaks.
    const guard = src.indexOf('selectionWordCount > 0')
    const pair = src.indexOf('appWordCountSelection', guard)
    const plain = src.indexOf('appWordCountN', pair)
    expect(guard).toBeGreaterThan(-1)
    expect(pair).toBeGreaterThan(guard)
    expect(plain).toBeGreaterThan(pair)
  })

  it('every locale carries the new key, so no language renders the raw placeholder', () => {
    const locales = [
      'ar',
      'cs',
      'de',
      'en',
      'es',
      'fr',
      'he',
      'hi',
      'id',
      'it',
      'ja',
      'ko',
      'ms',
      'nl',
      'pl',
      'pt',
      'ru',
      'th',
      'vi',
      'zh',
      'zh-TW',
    ]
    for (const loc of locales) {
      const file = join(SRC, 'i18n', 'app', `${loc}.ts`)
      expect(readFileSync(file, 'utf8'), `${loc}.ts`).toContain('appWordCountSelection')
    }
  })
})

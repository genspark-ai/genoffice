/**
 * The document page theme (genoffice#1811): docThemeIsDark resolves the
 * `<html data-doc-theme>` pin first — explicit light/dark from the shell
 * setting — and otherwise falls back to the UI theme exactly as before the
 * setting existed ('follow').
 */
import { afterEach, describe, expect, it } from 'vitest'
import { docThemeIsDark } from '../src/renderer/ui-theme'

function setAttrs(doc: string | null, ui: string | null): void {
  const el = document.documentElement
  const apply = (name: string, value: string | null): void => {
    if (value === null) el.removeAttribute(name)
    else el.setAttribute(name, value)
  }
  apply('data-doc-theme', doc)
  apply('data-theme', ui)
}

const originalMatchMedia = window.matchMedia

afterEach(() => {
  setAttrs(null, null)
  window.matchMedia = originalMatchMedia
})

describe('docThemeIsDark', () => {
  it('an explicit data-doc-theme pin wins over the UI theme', () => {
    setAttrs('light', 'dark')
    expect(docThemeIsDark()).toBe(false)
    setAttrs('dark', 'light')
    expect(docThemeIsDark()).toBe(true)
  })

  it('absent data-doc-theme follows the UI theme attribute', () => {
    setAttrs(null, 'dark')
    expect(docThemeIsDark()).toBe(true)
    setAttrs(null, 'light')
    expect(docThemeIsDark()).toBe(false)
  })

  it('with no attributes at all it falls back to the OS preference', () => {
    setAttrs(null, null)
    window.matchMedia = ((_query: string) => ({
      matches: true,
    })) as unknown as typeof window.matchMedia
    expect(docThemeIsDark()).toBe(true)
    window.matchMedia = ((_query: string) => ({
      matches: false,
    })) as unknown as typeof window.matchMedia
    expect(docThemeIsDark()).toBe(false)
  })
})

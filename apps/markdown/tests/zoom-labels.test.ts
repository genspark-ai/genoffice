import { describe, expect, it } from 'vitest'

import { strings } from '../src/renderer/i18n/strings'

/**
 * Status-bar zoom controls (App.tsx) must speak the UI language like the
 * PDF app does — never hardcoded English aria-labels. zh defines the key
 * set; every locale carries real zoom/zoomIn/zoomOut content.
 */

const locales = Object.keys(strings) as Array<keyof typeof strings>

describe('markdown zoom labels', () => {
  it.each(locales)('locale %s labels the zoom controls', (locale) => {
    const table = strings[locale] as Record<string, unknown>
    for (const key of ['zoom', 'zoomIn', 'zoomOut']) {
      expect(typeof table[key]).toBe('string')
      expect((table[key] as string).trim().length).toBeGreaterThan(0)
    }
  })
})

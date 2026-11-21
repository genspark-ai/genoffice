import { describe, expect, it } from 'vitest'
import { STANDARD_COLORS, THEME_COLORS } from '@genoffice/ui'
import { COLOR_NAME_KEYS } from '../src/renderer/components/ColorField'

describe('colour palette names', () => {
  it('names every swatch in the user language, as Word does', () => {
    const missing = [...THEME_COLORS, ...STANDARD_COLORS].filter(
      (swatch) => !(swatch.name in COLOR_NAME_KEYS),
    )
    expect(missing).toEqual([])
  })
})

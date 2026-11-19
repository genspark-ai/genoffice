import assert from 'node:assert/strict'
import { describe, expect, it } from 'vitest'
import { addPageBackgroundFloat, parsePageBgColor } from '../src/generate/page-settings'

function stubGenerator() {
  return {
    images: {},
    context: { pageWidthDxa: 12240, pageHeightDxa: 15840 },
    pageBackgroundFloat: undefined as unknown,
  }
}

describe('parsePageBgColor', () => {
  it('accepts bare 6-hex colors verbatim', () => {
    expect(parsePageBgColor('FF0000')).toBe('FF0000')
    expect(parsePageBgColor('00ff00')).toBe('00ff00')
    expect(parsePageBgColor('000000')).toBe('000000')
  })

  it('rejects short forms, names, hashes, and non-strings', () => {
    expect(parsePageBgColor('FFF')).toBeNull()
    expect(parsePageBgColor('red')).toBeNull()
    expect(parsePageBgColor('')).toBeNull()
    expect(parsePageBgColor('#FF0000')).toBeNull()
    expect(parsePageBgColor('FF00FG')).toBeNull()
    expect(parsePageBgColor(null)).toBeNull()
    expect(parsePageBgColor(undefined)).toBeNull()
    expect(parsePageBgColor(123456)).toBeNull()
  })
})

describe('addPageBackgroundFloat', () => {
  it('sets a float for a valid color and skips invalid ones', () => {
    const ok = stubGenerator()
    addPageBackgroundFloat(ok, { color: '1A2B3C' })
    assert.ok(ok.pageBackgroundFloat, 'valid color sets the float')

    for (const color of ['FFF', 'red', '', '#1A2B3C']) {
      const gen = stubGenerator()
      addPageBackgroundFloat(gen, { color })
      expect(gen.pageBackgroundFloat).toBeUndefined()
    }
  })

  it('does nothing without a pagebg node', () => {
    const gen = stubGenerator()
    addPageBackgroundFloat(gen, undefined)
    expect(gen.pageBackgroundFloat).toBeUndefined()
  })
})

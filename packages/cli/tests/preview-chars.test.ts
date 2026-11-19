import { describe, expect, it } from 'vitest'
import { MAX_PREVIEW_CHARS, previewChars } from '../src/preview'

const args = (flags: Record<string, string | true>) => ({ positionals: [], flags })

describe('previewChars', () => {
  it('uses the fallback without flags and Infinity with --full', () => {
    expect(previewChars(args({}), 200)).toBe(200)
    expect(previewChars(args({ full: true }), 200)).toBe(Infinity)
  })

  it('accepts positive integers within the cap', () => {
    expect(previewChars(args({ 'max-chars': '50' }), 200)).toBe(50)
    expect(previewChars(args({ 'max-chars': String(MAX_PREVIEW_CHARS) }), 200)).toBe(
      MAX_PREVIEW_CHARS,
    )
  })

  it('rejects non-positive, non-integer, and unbounded values', () => {
    for (const raw of ['0', '-5', '3.5', 'abc', String(MAX_PREVIEW_CHARS + 1), '1e21']) {
      expect(() => previewChars(args({ 'max-chars': raw }), 200), raw).toThrow(/--max-chars/)
    }
  })
})

import { describe, expect, it } from 'vitest'
import { normalizeMarkdownTheme } from '../src/shared/markdown-theme'

describe('normalizeMarkdownTheme', () => {
  it('passes known ids through', () => {
    expect(normalizeMarkdownTheme('nord')).toBe('nord')
    expect(normalizeMarkdownTheme('notion-light')).toBe('notion-light')
    expect(normalizeMarkdownTheme('drake-dark')).toBe('drake-dark')
    expect(normalizeMarkdownTheme('blubook')).toBe('blubook')
    expect(normalizeMarkdownTheme('newsprint')).toBe('newsprint')
    expect(normalizeMarkdownTheme('default')).toBe('default')
  })

  it('falls back to the default paper for missing or unknown values', () => {
    expect(normalizeMarkdownTheme(undefined)).toBe('default')
    expect(normalizeMarkdownTheme(null)).toBe('default')
    expect(normalizeMarkdownTheme('solarized')).toBe('default')
    expect(normalizeMarkdownTheme(42)).toBe('default')
    expect(normalizeMarkdownTheme({})).toBe('default')
  })
})

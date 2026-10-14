import { describe, expect, it } from 'vitest'
import { isMarkdownThemeDark, normalizeMarkdownTheme } from '../src/shared/markdown-theme'

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

describe('isMarkdownThemeDark', () => {
  it('classifies the dark palettes', () => {
    expect(isMarkdownThemeDark('nord')).toBe(true)
    expect(isMarkdownThemeDark('notion-dark')).toBe(true)
    expect(isMarkdownThemeDark('drake-dark')).toBe(true)
    expect(isMarkdownThemeDark('blubook')).toBe(false) // light paper, deep-blue code fences
    expect(isMarkdownThemeDark('notion-light')).toBe(false)
    expect(isMarkdownThemeDark('newsprint')).toBe(false)
    expect(isMarkdownThemeDark('default')).toBe(false)
  })
})

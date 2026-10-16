import { describe, expect, it } from 'vitest'
import {
  APPEARANCE_ACCENTS,
  APPEARANCE_BACKGROUNDS,
  activeAppearance,
  contrastRatio,
  markdownCustomThemeVars,
  validateAppearance,
} from '../src/feature-appearance'
import type { FeatureRecord } from '../src/features-skill'

function record(
  id: string,
  updatedAt: string,
  appearance?: FeatureRecord['appearance'],
): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: appearance ? 'appearance' : '',
    ...(appearance ? { appearance } : {}),
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt,
  }
}

describe('appearance contract', () => {
  it('accepts the names the host knows', () => {
    for (const background of APPEARANCE_BACKGROUNDS) {
      expect(validateAppearance({ background }).ok).toBe(true)
    }
    for (const accent of APPEARANCE_ACCENTS) {
      expect(validateAppearance({ accent }).ok).toBe(true)
    }
    expect(validateAppearance({ background: 'surface-subtle', accent: 'violet' }).ok).toBe(true)
  })

  it('refuses a raw value, naming the field', () => {
    const colour = validateAppearance({ background: '#ff00ff' })
    expect(colour.ok).toBe(false)
    if (!colour.ok) {
      expect(colour.errors[0]!.field).toBe('appearance.background')
      expect(colour.errors[0]!.message).toContain('surface-subtle')
    }
    const hue = validateAppearance({ accent: 'chartreuse' })
    expect(hue.ok).toBe(false)
    if (!hue.ok) expect(hue.errors[0]!.field).toBe('appearance.accent')
  })

  it('accepts a custom markdown theme and normalizes its colors', () => {
    const result = validateAppearance({
      markdown: { paper: '#F4ECD8', ink: '#333', accent: '#8c5b2a', font: 'serif' },
    })
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.appearance.markdown).toEqual({
        paper: '#f4ecd8',
        ink: '#333333',
        accent: '#8c5b2a',
        font: 'serif',
      })
    }
  })

  it('refuses a custom theme whose paper and ink are too close', () => {
    const unreadable = validateAppearance({ markdown: { paper: '#ffffff', ink: '#fefef0' } })
    expect(unreadable.ok).toBe(false)
    if (!unreadable.ok) {
      expect(unreadable.errors[0]!.field).toBe('appearance.markdown')
      expect(unreadable.errors[0]!.message).toMatch(/too close \(\d+\.\d:1\).*4\.5:1/)
    }
    // and an accent the paper swallows
    const invisible = validateAppearance({
      markdown: { paper: '#ffffff', ink: '#111111', accent: '#f5f5f5' },
    })
    expect(invisible.ok).toBe(false)
    if (!invisible.ok) expect(invisible.errors[0]!.field).toBe('appearance.markdown.accent')
  })

  it('refuses a custom theme with bad colors, font, or extra fields', () => {
    expect(validateAppearance({ markdown: { paper: 'green', ink: '#111111' } }).ok).toBe(false)
    expect(validateAppearance({ markdown: { paper: '#ffffff' } }).ok).toBe(false)
    expect(
      validateAppearance({ markdown: { paper: '#ffffff', ink: '#111111', font: 'comic' } }).ok,
    ).toBe(false)
    expect(
      validateAppearance({ markdown: { paper: '#ffffff', ink: '#111111', bold: true } }).ok,
    ).toBe(false)
  })

  it('computes the WCAG ratio and the custom theme variables', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 0)
    expect(contrastRatio('#ffffff', '#000000')).toBeCloseTo(21, 0)
    const { scheme, vars } = markdownCustomThemeVars({
      paper: '#1e2022',
      ink: '#d8d9da',
      accent: '#6cb2ff',
    })
    expect(scheme).toBe('dark')
    expect(vars).toEqual({
      '--md-custom-paper': '#1e2022',
      '--md-custom-ink': '#d8d9da',
      '--md-custom-accent': '#6cb2ff',
    })
    expect(markdownCustomThemeVars({ paper: '#f4ecd8', ink: '#43341f' }).scheme).toBe('light')
  })

  it('accepts a markdown reading theme by name and refuses an invented one', () => {
    for (const name of ['nord', 'drake-dark', 'notion-light', 'blubook', 'newsprint', 'default']) {
      const theme = validateAppearance({ markdown: name })
      expect(theme.ok).toBe(true)
      if (theme.ok) expect(theme.appearance.markdown).toBe(name)
    }
    const bad = validateAppearance({ markdown: 'monokai' })
    expect(bad.ok).toBe(false)
    if (!bad.ok) expect(bad.errors[0]!.field).toBe('appearance.markdown')
  })

  it('refuses unknown fields and an empty appearance', () => {
    const extra = validateAppearance({ background: 'surface', fontSize: 22 })
    expect(extra.ok).toBe(false)
    if (!extra.ok) expect(extra.errors[0]!.field).toBe('appearance.fontSize')

    expect(validateAppearance({}).ok).toBe(false)
    expect(validateAppearance(null).ok).toBe(false)
    expect(validateAppearance('surface').ok).toBe(false)
  })

  it('picks the most recently changed feature that declares one', () => {
    const older = record('a', '2026-01-01T00:00:00.000Z', { accent: 'teal' })
    const newer = record('b', '2026-02-01T00:00:00.000Z', { accent: 'violet' })
    const plain = record('c', '2026-03-01T00:00:00.000Z')
    expect(activeAppearance([older, newer, plain])).toEqual({ accent: 'violet' })
    expect(activeAppearance([older])).toEqual({ accent: 'teal' })
    expect(activeAppearance([plain])).toBeNull()
    expect(activeAppearance([])).toBeNull()
  })
})

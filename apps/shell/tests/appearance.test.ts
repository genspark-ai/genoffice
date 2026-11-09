import { describe, expect, it } from 'vitest'
import {
  DEFAULT_APPEARANCE,
  contrastRatio,
  loadAppearance,
  parseHex,
  saveAppearance,
  textOnAccent,
} from '../src/renderer/src/appearance'

function memoryStorage(initial?: string) {
  const data = new Map<string, string>()
  if (initial !== undefined) data.set('genoffice.appearance', initial)
  return {
    data,
    getItem: (k: string) => data.get(k) ?? null,
    setItem: (k: string, v: string) => void data.set(k, v),
    removeItem: (k: string) => void data.delete(k),
  }
}

describe('parseHex', () => {
  it('normalizes 3 and 6 digit colors with or without #', () => {
    expect(parseHex('#339CFF')).toBe('#339cff')
    expect(parseHex('fff')).toBe('#ffffff')
    expect(parseHex('  #0B0B0B ')).toBe('#0b0b0b')
  })
  it('rejects anything else', () => {
    for (const bad of ['', '#12', '#12345', 'red', '#gggggg', '#1234567']) {
      expect(parseHex(bad)).toBeNull()
    }
  })
})

describe('contrast', () => {
  it('matches the WCAG extremes', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 0)
    expect(contrastRatio('#777777', '#777777')).toBeCloseTo(1, 5)
  })
  it('picks readable text on an accent fill', () => {
    expect(textOnAccent('#ffffff')).toBe('#000000')
    expect(textOnAccent('#0b0b0b')).toBe('#ffffff')
    expect(textOnAccent('#339cff')).toBe('#000000')
  })
})

describe('persistence', () => {
  it('round-trips a custom appearance', () => {
    const storage = memoryStorage()
    saveAppearance(storage, {
      ...DEFAULT_APPEARANCE,
      accent: '#339cff',
      font: 'mono',
      transparency: false,
    })
    expect(loadAppearance(storage)).toEqual({
      accent: '#339cff',
      font: 'mono',
      transparency: false,
      confirmDelete: false,
    })
  })
  it('stores nothing for the defaults', () => {
    const storage = memoryStorage()
    saveAppearance(storage, { ...DEFAULT_APPEARANCE, accent: '#339cff' })
    saveAppearance(storage, { ...DEFAULT_APPEARANCE })
    expect(storage.data.size).toBe(0)
  })
  it('survives corrupt or hostile stored data', () => {
    expect(loadAppearance(memoryStorage('{nope'))).toEqual(DEFAULT_APPEARANCE)
    const loaded = loadAppearance(
      memoryStorage(JSON.stringify({ accent: 'url(x)', font: '__proto__' })),
    )
    expect(loaded.accent).toBeNull()
    expect(loaded.font).toBe('system')
    expect(loadAppearance(undefined)).toEqual(DEFAULT_APPEARANCE)
  })
})

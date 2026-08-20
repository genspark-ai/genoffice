import { existsSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { advanceWidths, isFamilyInstalled } from '../src/index'

const darwin = process.platform === 'darwin'
const hasHelvetica = darwin && existsSync('/System/Library/Fonts/Helvetica.ttc')

describe('isFamilyInstalled', () => {
  it('resolves exact normalized family names only', () => {
    expect(isFamilyInstalled('No Such Font Family ZZZ')).toBe(false)
    expect(isFamilyInstalled('')).toBe(false)
    if (!hasHelvetica) return
    expect(isFamilyInstalled('Helvetica')).toBe(true)
    expect(isFamilyInstalled('helvetica')).toBe(true)
    expect(isFamilyInstalled('Helveticaish')).toBe(false)
  })
})

describe('advanceWidths', () => {
  it('returns null for uninstalled families', () => {
    expect(advanceWidths('No Such Font Family ZZZ', 'x', 12)).toBeNull()
  })

  it('returns per-codepoint twips that scale linearly with size', () => {
    if (!hasHelvetica) return
    const at12 = advanceWidths('Helvetica', 'Hi x', 12)!
    expect(at12).toHaveLength(4)
    for (const w of at12) expect(w).toBeGreaterThan(0)
    const at24 = advanceWidths('Helvetica', 'Hi x', 24)!
    at24.forEach((w, i) => expect(w).toBeCloseTo(2 * at12[i]!, 6))
    // 'i' is narrower than 'H' in a proportional face
    expect(at12[1]!).toBeLessThan(at12[0]!)
  })

  it('selects a heavier face for bold and reports wider stems', () => {
    if (!hasHelvetica) return
    const sum = (a: number[]) => a.reduce((x, y) => x + y, 0)
    const reg = advanceWidths('Helvetica', 'Hamburgefonstiv', 12)!
    const bold = advanceWidths('Helvetica', 'Hamburgefonstiv', 12, { bold: true })!
    expect(sum(bold)).toBeGreaterThan(sum(reg))
  })

  it('marks unmapped codepoints as NaN instead of guessing', () => {
    if (!hasHelvetica) return
    const w = advanceWidths('Helvetica', '\u{10FFF0}x', 12)!
    expect(Number.isNaN(w[0]!)).toBe(true)
    expect(w[1]!).toBeGreaterThan(0)
  })
})

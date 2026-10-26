import { describe, expect, it } from 'vitest'
import { escapeReLiteral } from '../src/renderer/components/ribbon-references-tab'

const seqRe = (lbl: string) => new RegExp(`SEQ\\s+${escapeReLiteral(lbl)}[\\s\\\\]`)

describe('escapeReLiteral (caption SEQ matcher)', () => {
  it('matches plain labels literally', () => {
    expect(seqRe('Figure').test('SEQ Figure 1')).toBe(true)
    expect(seqRe('Table').test('SEQ Table\t2')).toBe(true)
    expect(seqRe('Figure').test('SEQ Figures 1')).toBe(false)
  })

  it('does not throw for labels with regex metacharacters', () => {
    // a locale label like "Figure (2)" previously threw SyntaxError during the
    // dialog render, because the placeholder calls nextNumber on every render
    expect(() => seqRe('Figure (2)')).not.toThrow()
    expect(seqRe('Figure (2)').test('SEQ Figure (2) 1')).toBe(true)
    expect(seqRe('Figure (2)').test('SEQ Figure X (2) 1')).toBe(false)
  })

  it('treats . + $ as literals', () => {
    expect(seqRe('A.B+C$').test('SEQ A.B+C$ 1')).toBe(true)
    expect(seqRe('A.B+C$').test('SEQ AXBYCZ 1')).toBe(false)
  })
})

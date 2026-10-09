/**
 * Sakkal Majalla aliases (fonts.css): Word renders the M365 cloud face real;
 * the per-class stand-ins must land on its advances (hmtx of the cloud cut)
 * and pin its hhea box.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/fonts/fonts.css'), 'utf8')

function faces(family: string) {
  return [...css.matchAll(/@font-face \{([^}]*)\}/g)]
    .map((m) => m[1])
    .filter((body) => body.includes(`font-family: '${family}'`))
    .map((body) => ({
      range: /unicode-range:\s*([^;]+);/.exec(body)?.[1].replace(/\s+/g, ' ').trim(),
      adjust: Number(/size-adjust:\s*([\d.]+)%/.exec(body)?.[1] ?? 100) / 100,
      bold: /font-weight: bold/.test(body),
      src: /src:\s*([^;]+);/.exec(body)?.[1] ?? '',
      ascent: /ascent-override:\s*([\d.]+)%/.exec(body)?.[1],
      descent: /descent-override:\s*([\d.]+)%/.exec(body)?.[1],
    }))
}
const find = (rules: ReturnType<typeof faces>, range: string, bold = false) =>
  rules.find((r) => r.range === range && r.bold === bold)

describe('Sakkal Majalla GO', () => {
  const rules = faces('Sakkal Majalla GO')

  it('Arabic letters ride the Noto Naskh subset at 0.8055, weight-normal only', () => {
    const arabic = rules.filter((r) => r.range?.startsWith('U+0600-065F'))
    expect(arabic).toHaveLength(1)
    expect(arabic[0].src).toContain('NotoNaskhArabic-Regular-subset.woff2')
    expect(arabic[0].adjust).toBeCloseTo(0.8055, 4)
    expect(arabic[0].range).not.toContain('U+0660')
    // Noto Arabic-Indic digit 0.449 em -> Sakkal 0.4219
    expect(0.449 * (find(rules, 'U+0660-0669, U+06F0-06F9')?.adjust ?? 0)).toBeCloseTo(0.4219, 3)
  })

  it('ASCII digits and space land on the cloud hmtx (0.4053 / 0.1514 em of Liberation Sans)', () => {
    expect(0.5562 * (find(rules, 'U+0030-0039')?.adjust ?? 0)).toBeCloseTo(0.4053, 3)
    expect(0.2778 * (find(rules, 'U+0020, U+00A0')?.adjust ?? 0)).toBeCloseTo(0.1514, 3)
    expect(find(rules, 'U+0030-0039')?.src).toContain('LiberationSans-Regular.ttf')
  })

  it('dashes and symbols stay in the alias instead of falling to Al Bayan', () => {
    // Sakkal en dash 0.5068 em vs Liberation Sans 0.5562
    expect(0.5562 * (find(rules, 'U+2010-2015')?.adjust ?? 0)).toBeCloseTo(0.5068, 3)
    expect(rules.some((r) => r.range?.includes('U+00D7') && r.range.includes('U+2016-2027'))).toBe(
      true,
    )
  })

  it('every face is weight-normal (one weight group per family in Blink) and pins the Word hhea box', () => {
    expect(rules.length).toBeGreaterThan(0)
    for (const r of [...rules, ...faces('Sakkal Majalla Latin GO')]) {
      expect(r.bold).toBe(false)
      expect(r.ascent).toBe('88.38')
      expect(r.descent).toBe('51.27')
    }
  })

  it('Latin letters live in their own family, scaled 0.67', () => {
    const latin = faces('Sakkal Majalla Latin GO')
    expect(find(latin, 'U+0041-005A, U+0061-007A')?.adjust).toBeCloseTo(0.67, 3)
    expect(rules.some((r) => r.range?.includes('U+0041'))).toBe(false)
  })
})

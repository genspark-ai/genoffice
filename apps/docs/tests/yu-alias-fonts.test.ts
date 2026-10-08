/**
 * Yu Mincho / Yu Gothic Latin advance aliases (fonts.css): Word lays the
 * Office-private faces out with its own yumin.ttf / YuGoth*.ttc, whose Latin
 * is ~8% narrower than the Noto Serif CJK stand-in a missing declare fell to.
 * The Liberation tier must reproduce the Word advances the class scale is
 * exact for (digits, space), and the Apple tier must not scale a cut whose
 * advances already equal Word's.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const FONTS = join(__dirname, '../src/renderer/fonts')
const css = readFileSync(join(FONTS, 'fonts.css'), 'utf8')

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
    }))
}
const find = (rules: ReturnType<typeof faces>, range: string, bold = false) =>
  rules.find((r) => r.range === range && r.bold === bold)

describe('Yu Mincho Serif GO (Liberation Serif tier)', () => {
  const rules = faces('Yu Mincho Serif GO')
  // Liberation Serif carries the Times advances: space 0.25 em, digits 0.5 em

  it('space and digits land on the Word yumin.ttf advances (0.2598 / 0.542 em)', () => {
    const space = find(rules, 'U+0020')
    const digit = find(rules, 'U+0030-0039')
    expect(space?.src).toContain('LiberationSerif-Regular.ttf')
    expect(0.25 * (space?.adjust ?? 0)).toBeCloseTo(0.2598, 3)
    expect(0.5 * (digit?.adjust ?? 0)).toBeCloseTo(0.542, 3)
  })

  it('bold tier rides Liberation Serif Bold and every face pins the Word hhea box', () => {
    expect(find(rules, 'U+0020', true)?.src).toContain('LiberationSerif-Bold.ttf')
    for (const r of rules) expect(r.ascent).toBe('88')
  })
})

describe('Apple tiers', () => {
  it('Yu Mincho GO leads with the Apple Medium cut, space at 0.78x of its 1/3 em', () => {
    const rules = faces('Yu Mincho GO')
    expect(find(rules, 'U+0020')?.src).toContain("local('YuMin-Medium')")
    expect(find(rules, 'U+0020')?.adjust).toBeCloseTo(0.78, 2)
    expect(find(rules, 'U+0020-007E, U+00A0-00FF', true)?.src).toContain("local('YuMin-Demibold')")
  })

  it('Yu Gothic GO takes the Apple Medium/Bold cuts unscaled (advances equal Word yugoth)', () => {
    const rules = faces('Yu Gothic GO')
    expect(rules).toHaveLength(2)
    for (const r of rules) expect(r.adjust).toBe(1)
    expect(find(rules, 'U+0020-007E, U+00A0-00FF')?.src).toContain("local('YuGo-Medium')")
    expect(find(rules, 'U+0020-007E, U+00A0-00FF', true)?.src).toContain("local('YuGo-Bold')")
  })
})

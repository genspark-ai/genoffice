/** Word's default endnote numbering is lowercase roman (footnotes stay arabic) — the visual cue that separates the two note kinds. */
const ROMAN: Array<[number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
]

export function toRoman(n: number): string {
  if (!Number.isFinite(n) || n < 1) return String(n)
  let rest = Math.floor(n)
  let out = ''
  for (const [value, glyph] of ROMAN) {
    while (rest >= value) {
      out += glyph
      rest -= value
    }
  }
  return out
}

export function noteMarkText(kind: 'footnote' | 'endnote', no: number): string {
  return kind === 'endnote' ? toRoman(no) : String(no)
}

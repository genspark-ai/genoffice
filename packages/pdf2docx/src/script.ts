/**
 * Unicode script classification — the gate for every language-sensitive rule
 * (word spacing, span splitting, font-slot mapping, RTL detection). P1 covers
 * the scripts the pipeline must distinguish; everything else is 'common'.
 */

export type UnicodeScript =
  'latin' | 'cjk' | 'hangul' | 'kana' | 'thai' | 'arabic' | 'hebrew' | 'common'

interface Range {
  from: number
  to: number
  script: UnicodeScript
}

// Ordered sparse ranges; lookup is linear (short list, hot path caches per char anyway)
const RANGES: Range[] = [
  { from: 0x41, to: 0x5a, script: 'latin' },
  { from: 0x61, to: 0x7a, script: 'latin' },
  { from: 0xaa, to: 0xaa, script: 'latin' },
  { from: 0xba, to: 0xba, script: 'latin' },
  { from: 0xc0, to: 0xd6, script: 'latin' },
  { from: 0xd8, to: 0xf6, script: 'latin' },
  { from: 0xf8, to: 0x2af, script: 'latin' }, // Latin-1 letters + Extended-A/B + IPA
  { from: 0x1e00, to: 0x1eff, script: 'latin' }, // Latin Extended Additional
  { from: 0x2c60, to: 0x2c7f, script: 'latin' },
  { from: 0xa720, to: 0xa7ff, script: 'latin' },
  { from: 0x590, to: 0x5ff, script: 'hebrew' },
  { from: 0xfb1d, to: 0xfb4f, script: 'hebrew' }, // presentation forms
  { from: 0x600, to: 0x6ff, script: 'arabic' },
  { from: 0x750, to: 0x77f, script: 'arabic' },
  { from: 0x8a0, to: 0x8ff, script: 'arabic' },
  { from: 0xfb50, to: 0xfdff, script: 'arabic' }, // presentation forms A
  { from: 0xfe70, to: 0xfeff, script: 'arabic' }, // presentation forms B (FEFF itself is BOM but never appears as text)
  { from: 0xe00, to: 0xe7f, script: 'thai' },
  { from: 0x1100, to: 0x11ff, script: 'hangul' }, // Jamo
  { from: 0x3130, to: 0x318f, script: 'hangul' }, // compatibility Jamo
  { from: 0xa960, to: 0xa97f, script: 'hangul' },
  { from: 0xac00, to: 0xd7ff, script: 'hangul' }, // syllables + Jamo Extended-B
  { from: 0x3040, to: 0x30ff, script: 'kana' }, // hiragana + katakana
  { from: 0x31f0, to: 0x31ff, script: 'kana' }, // katakana phonetic extensions
  { from: 0xff66, to: 0xff9f, script: 'kana' }, // halfwidth katakana
  { from: 0x2e80, to: 0x2fdf, script: 'cjk' }, // radicals
  { from: 0x3000, to: 0x303f, script: 'cjk' }, // CJK symbols & punctuation (fullwidth 。、「」)
  { from: 0x3400, to: 0x4dbf, script: 'cjk' }, // ext A
  { from: 0x4e00, to: 0x9fff, script: 'cjk' },
  { from: 0xf900, to: 0xfaff, script: 'cjk' }, // compatibility ideographs
  { from: 0xfe30, to: 0xfe4f, script: 'cjk' }, // vertical forms
  { from: 0xff00, to: 0xff65, script: 'cjk' }, // fullwidth ASCII + punctuation
  { from: 0xffe0, to: 0xffee, script: 'cjk' }, // fullwidth signs
  { from: 0x20000, to: 0x3134f, script: 'cjk' }, // ext B..G
]

export function scriptOf(code: number): UnicodeScript {
  for (const r of RANGES) if (code >= r.from && code <= r.to) return r.script
  return 'common'
}

/**
 * Scripts whose characters NEVER get machine-inserted spaces between them
 * (hard rule; violating it garbles zh/ja/th text). Hangul is excluded: Korean
 * writes real inter-word spaces, so it follows the Latin gap rule.
 */
export function isNoSpaceScript(script: UnicodeScript): boolean {
  return script === 'cjk' || script === 'kana' || script === 'thai'
}

/** scripts written right-to-left (page-level detection gates the P1 bitmap fallback) */
export function isRtlScript(script: UnicodeScript): boolean {
  return script === 'arabic' || script === 'hebrew'
}

/** CJK-family scripts that map to the w:eastAsia font slot in OOXML */
export function isEastAsianScript(script: UnicodeScript): boolean {
  return script === 'cjk' || script === 'kana' || script === 'hangul'
}

// Combining marks / zero-width code points that must merge into the preceding
// base character's cluster (Thai vowels/tones, Arabic harakat, Latin diacritics…)
const COMBINING: Array<[number, number]> = [
  [0x300, 0x36f], // combining diacritical marks
  [0x483, 0x489],
  [0x591, 0x5c7], // Hebrew points
  [0x610, 0x61a],
  [0x64b, 0x65f], // Arabic harakat
  [0x670, 0x670],
  [0x6d6, 0x6dc],
  [0x6df, 0x6e4],
  [0x6e7, 0x6e8],
  [0x6ea, 0x6ed],
  [0xe31, 0xe31], // Thai mai han-akat
  [0xe34, 0xe3a], // Thai vowels below/above
  [0xe47, 0xe4e], // Thai tone marks
  [0xeb1, 0xeb1],
  [0xeb4, 0xebc],
  [0xec8, 0xecd],
  [0x1ab0, 0x1aff],
  [0x1dc0, 0x1dff],
  [0x200b, 0x200f], // zero-width space/joiners + directional marks
  [0x20d0, 0x20ff],
  [0xfe00, 0xfe0f], // variation selectors
  [0xfe20, 0xfe2f],
]

export function isCombiningMark(code: number): boolean {
  for (const [from, to] of COMBINING) if (code >= from && code <= to) return true
  return false
}

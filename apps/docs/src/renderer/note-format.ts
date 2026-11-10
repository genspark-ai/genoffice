import type { NoteProps } from '@genoffice/docx-engine'

export type NoteKind = 'footnote' | 'endnote'

/** Word's default numbering: footnotes arabic, endnotes lowercase roman — the visual cue that separates the two note kinds. */
const DEFAULT_FMT: Record<NoteKind, string> = { footnote: 'decimal', endnote: 'lowerRoman' }

/** current document's w:numFmt per note kind (set on load, before the body renders its reference marks) */
const numFmt: Record<NoteKind, string> = { ...DEFAULT_FMT }

export function setNoteNumFmts(props: { footnote?: NoteProps; endnote?: NoteProps }): void {
  numFmt.footnote = props.footnote?.numFmt ?? DEFAULT_FMT.footnote
  numFmt.endnote = props.endnote?.numFmt ?? DEFAULT_FMT.endnote
}

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

function toLetter(n: number): string {
  if (!Number.isFinite(n) || n < 1) return String(n)
  const floored = Math.floor(n)
  // A footnote count is file-controlled: cap the repeat expansion so a huge
  // number cannot allocate a multi-megabyte marker string.
  if (floored > 26 * MAX_NOTE_REPS) return String(floored)
  const idx = (floored - 1) % 26
  const reps = Math.floor((floored - 1) / 26) + 1
  return String.fromCharCode(97 + idx).repeat(reps)
}

/** Word's chicago sequence: * † ‡ §, doubled on each wrap */
const CHICAGO = ['*', '†', '‡', '§']
/** Repeat expansions beyond this fall back to the decimal number. */
const MAX_NOTE_REPS = 20
function toChicago(n: number): string {
  if (!Number.isFinite(n) || n < 1) return String(n)
  const floored = Math.floor(n)
  if (floored > CHICAGO.length * MAX_NOTE_REPS) return String(floored)
  const idx = (floored - 1) % CHICAGO.length
  const reps = Math.floor((floored - 1) / CHICAGO.length) + 1
  return CHICAGO[idx].repeat(reps)
}

/** display text of note number `n` in a w:numFmt (unknown formats fall back to decimal) */
export function formatNoteNumber(fmt: string | undefined, n: number): string {
  switch (fmt) {
    case 'lowerRoman':
      return toRoman(n)
    case 'upperRoman':
      return toRoman(n).toUpperCase()
    case 'lowerLetter':
      return toLetter(n)
    case 'upperLetter':
      return toLetter(n).toUpperCase()
    case 'chicago':
      return toChicago(n)
    default:
      return String(n)
  }
}

export function noteMarkText(kind: NoteKind, no: number): string {
  return formatNoteNumber(numFmt[kind], no)
}

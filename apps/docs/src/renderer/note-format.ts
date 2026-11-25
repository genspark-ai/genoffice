import type { NoteProps, SectionInfo, SectionSettings } from '@genoffice/docx-engine'

export type NoteKind = 'footnote' | 'endnote'
export type SectionNoteProps = Record<NoteKind, NoteProps>
/** what places a block in its section (editor/section-endnotes' topBlockSections) */
export type NoteSection = Pick<SectionInfo, 'lastBlockIndex' | 'sectPrXml'> & {
  pendingBreak?: boolean
}

/** Word's default numbering: footnotes arabic, endnotes lowercase roman — the visual cue that separates the two note kinds. */
const DEFAULT_FMT: Record<NoteKind, string> = { footnote: 'decimal', endnote: 'lowerRoman' }
export const defaultNoteNumFmt = (kind: NoteKind): string => DEFAULT_FMT[kind]

/** the note properties in effect in each section: the document's, overlaid by the section's own */
export function sectionNotePropsOf(
  base: Partial<SectionNoteProps> | undefined,
  sections: ReadonlyArray<Pick<SectionSettings, 'footnotePr' | 'endnotePr'>>,
): SectionNoteProps[] {
  return sections.map((s) => ({
    footnote: { ...base?.footnote, ...s.footnotePr },
    endnote: { ...base?.endnote, ...s.endnotePr },
  }))
}

/** current document's note properties per section (set on load, before the body renders its reference marks) */
let props: SectionNoteProps[] = []
let sections: readonly NoteSection[] = []

/** one set of properties stands for every section */
export function setNoteProps(
  next: Partial<SectionNoteProps> | SectionNoteProps[],
  sectionList?: readonly NoteSection[],
): void {
  props = Array.isArray(next)
    ? next
    : [{ footnote: next.footnote ?? {}, endnote: next.endnote ?? {} }]
  if (sectionList) sections = sectionList
}

const propsIn = (kind: NoteKind, section?: number): NoteProps =>
  props[Math.min(section ?? props.length - 1, props.length - 1)]?.[kind] ?? {}

/** the number format of a section's notes; the last section's when none is named */
export const noteNumFmt = (kind: NoteKind, section?: number): string =>
  propsIn(kind, section).numFmt ?? DEFAULT_FMT[kind]

export type NoteRestart = 'continuous' | 'eachSect' | 'eachPage'
export interface NoteCountRule {
  start: number
  restart: NoteRestart
  fmt: string
}

/** the page (0-based) the last layout put each footnote's first reference on; empty before a layout */
let footnotePages: ReadonlyMap<string, number> = new Map()
export function setFootnotePages(next: ReadonlyMap<string, number>): void {
  footnotePages = next
}

/** how the reference marks of a kind count: per section, what the count opens at and where it starts over */
export function noteNumbering(kind: NoteKind): {
  sections: readonly NoteSection[]
  of: (section: number) => NoteCountRule
  pageOf: (id: string) => number | undefined
} {
  return {
    sections,
    of: (section) => {
      const p = propsIn(kind, section)
      return {
        start: p.numStart ?? 1,
        // Word numbers endnotes continuously or per section only
        restart:
          p.numRestart === 'eachSect' || (p.numRestart === 'eachPage' && kind === 'footnote')
            ? p.numRestart
            : 'continuous',
        fmt: p.numFmt ?? DEFAULT_FMT[kind],
      }
    },
    pageOf: (id) => (kind === 'footnote' ? footnotePages.get(id) : undefined),
  }
}

/** whether any section restarts its footnote numbers on each page */
export const footnotesRestartPerPage = (list: readonly SectionNoteProps[] = props): boolean =>
  list.some((p) => p.footnote.numRestart === 'eachPage')

/** a numbered reference mark in body order: its note, its section, and the page it was laid out on when known */
export interface NoteRefSite {
  id: string
  section: number
  page?: number
}

/**
 * The number each note shows. A mark is its section's start value plus the
 * running count; the count starts over where an eachSect or eachPage section
 * begins and, in an eachPage section, at the first mark of every page. A mark whose page is
 * not known yet counts on the page of the mark before it. A note referenced
 * twice keeps its first number.
 */
export function assignNoteNumbers(
  refs: readonly NoteRefSite[],
  of: (section: number) => NoteCountRule,
): Map<string, { num: number; fmt: string }> {
  const shown = new Map<string, { num: number; fmt: string }>()
  let count = 0
  let section = -1
  let page: number | undefined
  for (const ref of refs) {
    const rule = of(ref.section)
    // a section without marks of its own still starts a per-section count over; a
    // per-page section opens at its start value even where it begins in the middle of a page
    if (section < ref.section && rule.restart === 'eachPage') count = 0
    for (; section < ref.section; section++) if (of(section + 1).restart === 'eachSect') count = 0
    if (shown.has(ref.id)) continue
    const here = ref.page ?? page
    if (rule.restart === 'eachPage' && here !== page) count = 0
    page = here
    shown.set(ref.id, { num: rule.start + count++, fmt: rule.fmt })
  }
  return shown
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

/** a mark in the format its section numbers in; the last section's when the mark carries none */
export function noteMarkText(kind: NoteKind, no: number, fmt?: string | null): string {
  return formatNoteNumber(fmt ?? noteNumFmt(kind), no)
}

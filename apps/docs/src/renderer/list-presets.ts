/**
 * List galleries and dialogs share these level presets: the bullet / numbering /
 * multilevel libraries, the number-style catalogue with samples, the recently
 * used presets (localStorage) and the "lists in this document" scan.
 */
import { formatNumber, type CustomNumberingLevel, type NumberingDef } from '@genoffice/docx-engine'
import { customLevelFromNumberingLevel } from '@genoffice/docx-engine'

export const TWIPS_PER_CM = 567

/** Bullet library: the chosen symbol lands on the caret's level (`extraLevel`); the other levels rotate through ○/■ */
export function bulletPresetLevels(
  glyph: string,
  extra: Partial<CustomNumberingLevel> = {},
  extraLevel = 0,
): CustomNumberingLevel[] {
  const rotation = [glyph, '○', '■']
  return Array.from({ length: 9 }, (_, i) => ({
    numFmt: 'bullet',
    lvlText: rotation[(((i - extraLevel) % 3) + 3) % 3],
    indentLeft: 720 * (i + 1),
    hanging: 360,
    ...(i === extraLevel ? extra : {}),
  }))
}

/** Numbering library: the same format continues per level (%1 in the pattern becomes each level's counter); the extras land on the caret's level */
export function numberPresetLevels(
  numFmt: string,
  pattern: string,
  extra: Partial<CustomNumberingLevel> = {},
  extraLevel = 0,
): CustomNumberingLevel[] {
  return Array.from({ length: 9 }, (_, i) => ({
    numFmt,
    lvlText: pattern.replace('%1', `%${i + 1}`),
    indentLeft: 720 * (i + 1),
    hanging: 360,
    ...(i === extraLevel ? extra : {}),
  }))
}

export const BULLET_LIBRARY = ['•', '○', '■', '◆', '➢', '✦']

export const NUMBER_LIBRARY: Array<{ numFmt: string; pattern: string }> = [
  { numFmt: 'decimal', pattern: '%1.' },
  { numFmt: 'decimal', pattern: '%1)' },
  { numFmt: 'upperRoman', pattern: '%1.' },
  { numFmt: 'upperLetter', pattern: '%1.' },
  { numFmt: 'lowerLetter', pattern: '%1)' },
  { numFmt: 'lowerRoman', pattern: '%1.' },
  { numFmt: 'chineseCountingThousand', pattern: '%1、' },
  { numFmt: 'decimalEnclosedCircle', pattern: '%1' },
]

const nine = (make: (i: number) => CustomNumberingLevel) =>
  Array.from({ length: 9 }, (_, i) => make(i))
const legal = (i: number) => `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}.`
const HEADING_STYLES = Array.from({ length: 9 }, (_, i) => `Heading${i + 1}`)

/** Word's multilevel list library (List Library gallery order) */
export const MULTILEVEL_LIBRARY: CustomNumberingLevel[][] = [
  // 1) / a) / i)
  nine((i) => {
    const fmts = ['decimal', 'lowerLetter', 'lowerRoman']
    return { numFmt: fmts[i % 3], lvlText: `%${i + 1})`, indentLeft: 720 * (i + 1), hanging: 360 }
  }),
  // 1. / 1.1. / 1.1.1.
  nine((i) => ({ numFmt: 'decimal', lvlText: legal(i), indentLeft: 720 * (i + 1), hanging: 432 })),
  // ● / ○ / ■
  bulletPresetLevels('•'),
  // Article I. / Section 1.01 / (a)
  nine((i) => {
    if (i === 0)
      return {
        numFmt: 'upperRoman',
        lvlText: 'Article %1.',
        indentLeft: 0,
        hanging: 0,
        suff: 'space',
      }
    if (i === 1)
      return {
        numFmt: 'decimalZero',
        lvlText: 'Section %1.%2',
        indentLeft: 0,
        hanging: 0,
        suff: 'space',
      }
    const fmts = [
      'lowerLetter',
      'lowerRoman',
      'decimal',
      'lowerLetter',
      'lowerRoman',
      'decimal',
      'lowerLetter',
    ]
    return { numFmt: fmts[i - 2], lvlText: `(%${i + 1})`, indentLeft: 720 * (i - 1), hanging: 360 }
  }),
  // 1 Heading 1 / 1.1 Heading 2 (levels linked to heading styles)
  nine((i) => ({
    numFmt: 'decimal',
    lvlText: legal(i).slice(0, -1),
    indentLeft: 432 * (i + 1),
    hanging: 432 * (i + 1),
    pStyle: HEADING_STYLES[i],
  })),
  // I. / A. / 1. / a) / (1) / (a) / (i) / (a) / (i)
  nine((i) => {
    const spec = [
      ['upperRoman', '%1.'],
      ['upperLetter', '%2.'],
      ['decimal', '%3.'],
      ['lowerLetter', '%4)'],
      ['decimal', '(%5)'],
      ['lowerLetter', '(%6)'],
      ['lowerRoman', '(%7)'],
      ['lowerLetter', '(%8)'],
      ['lowerRoman', '(%9)'],
    ][i]
    return { numFmt: spec[0], lvlText: spec[1], indentLeft: 720 * (i + 1), hanging: 360 }
  }),
  // Chapter 1 (heading-linked, deeper levels unnumbered)
  nine((i) =>
    i === 0
      ? {
          numFmt: 'decimal',
          lvlText: 'Chapter %1',
          indentLeft: 0,
          hanging: 0,
          suff: 'space',
          pStyle: 'Heading1',
        }
      : {
          numFmt: 'none',
          lvlText: '',
          indentLeft: 0,
          hanging: 0,
          suff: 'nothing',
          pStyle: HEADING_STYLES[i],
        },
  ),
  // CJK official-document hierarchy: numeral + comma / parenthesized numeral / 1.
  nine((i) => {
    if (i === 0)
      return { numFmt: 'chineseCountingThousand', lvlText: '%1、', indentLeft: 720, hanging: 425 }
    if (i === 1)
      return { numFmt: 'chineseCountingThousand', lvlText: '(%2)', indentLeft: 1440, hanging: 425 }
    return { numFmt: 'decimal', lvlText: `%${i + 1}.`, indentLeft: 720 * (i + 1), hanging: 360 }
  }),
]

/** Number styles offered by the define dialogs (Word's list, CJK formats included) */
export const LIST_NUM_FMTS = [
  'decimal',
  'decimalZero',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
  'ordinal',
  'cardinalText',
  'ordinalText',
  'decimalEnclosedCircle',
  'decimalEnclosedParen',
  'decimalFullWidth',
  'chineseCounting',
  'chineseCountingThousand',
  'chineseLegalSimplified',
  'ideographTraditional',
  'ideographZodiac',
  'ideographLegalTraditional',
  'ideographEnclosedCircle',
  'japaneseCounting',
  'aiueoFullWidth',
  'irohaFullWidth',
  'koreanDigital',
  'ganada',
  'chosung',
  'russianLower',
  'russianUpper',
  'lowerGreek',
  'upperGreek',
  'none',
] as const

export type ListNumFmt = (typeof LIST_NUM_FMTS)[number]

/** "1, 2, 3, ..." style sample of a number format */
export function numFmtSample(numFmt: string): string {
  if (numFmt === 'bullet') return '● ○ ■'
  if (numFmt === 'none') return '(none)'
  return `${[1, 2, 3].map((v) => formatNumber(v, numFmt)).join(', ')}, ...`
}

/** The level's number text when every level counter is 1 (gallery/dialog preview) */
export function previewLevelText(levels: CustomNumberingLevel[], ilvl: number): string {
  const l = levels[ilvl]
  if (!l) return ''
  if (l.numFmt === 'bullet') return l.lvlText
  return l.lvlText.replace(/%(\d)/g, (_, n: string) =>
    formatNumber(
      levels[Number(n) - 1]?.start ?? 1,
      l.isLgl ? 'decimal' : (levels[Number(n) - 1]?.numFmt ?? 'decimal'),
    ),
  )
}

export type ListPresetKind = 'bullets' | 'numbers' | 'multi'

const RECENT_KEY: Record<ListPresetKind, string> = {
  bullets: 'aidocs.list.recent.bullets',
  numbers: 'aidocs.list.recent.numbers',
  multi: 'aidocs.list.recent.multi',
}
const RECENT_MAX = 6

export function recentListPresets(kind: ListPresetKind): CustomNumberingLevel[][] {
  try {
    const raw = localStorage.getItem(RECENT_KEY[kind])
    const parsed: unknown = raw ? JSON.parse(raw) : []
    return Array.isArray(parsed)
      ? parsed.filter((p): p is CustomNumberingLevel[] => Array.isArray(p) && p.length > 0)
      : []
  } catch {
    return []
  }
}

export function presetKey(levels: CustomNumberingLevel[]): string {
  return levels
    .slice(0, 3)
    .map((l) => `${l.numFmt}|${l.lvlText}|${l.picBulletId ?? ''}|${l.font ?? ''}`)
    .join('/')
}

export function rememberListPreset(kind: ListPresetKind, levels: CustomNumberingLevel[]): void {
  const key = presetKey(levels)
  const next = [levels, ...recentListPresets(kind).filter((p) => presetKey(p) !== key)].slice(
    0,
    RECENT_MAX,
  )
  try {
    localStorage.setItem(RECENT_KEY[kind], JSON.stringify(next))
  } catch {
    /* storage full or disabled: the gallery just shows no recents */
  }
}

/** Distinct list formats the document already uses, for the "Document …" gallery sections */
export function documentListPresets(
  defs: Iterable<NumberingDef>,
): Record<ListPresetKind, CustomNumberingLevel[][]> {
  const out: Record<ListPresetKind, CustomNumberingLevel[][]> = {
    bullets: [],
    numbers: [],
    multi: [],
  }
  const seen: Record<ListPresetKind, Set<string>> = {
    bullets: new Set(),
    numbers: new Set(),
    multi: new Set(),
  }
  for (const def of defs) {
    const ilvls = Object.keys(def.levels)
      .map(Number)
      .sort((a, b) => a - b)
    if (ilvls.length === 0) continue
    const levels = ilvls.map((i) => {
      const l = customLevelFromNumberingLevel(def.levels[i])
      // a picture bullet's image is only display data; the preset keeps the id
      return l
    })
    const first = levels[0]
    const kind: ListPresetKind =
      first.numFmt === 'bullet'
        ? 'bullets'
        : levels.slice(1, 3).some((l) => /%[2-9]/.test(l.lvlText) && /%1/.test(l.lvlText))
          ? 'multi'
          : 'numbers'
    const key = presetKey(levels)
    if (seen[kind].has(key)) continue
    seen[kind].add(key)
    out[kind].push(levels)
  }
  return out
}

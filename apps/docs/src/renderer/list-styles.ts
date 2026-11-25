/**
 * Named list styles (styles.xml type "numbering") for the List Styles tab of
 * Bullets and Numbering: Word's built-in four plus the ones the document defines.
 */
import {
  computeListMarkers,
  customLevelFromNumberingLevel,
  numberingLevelFromCustom,
  type CustomNumberingLevel,
  type ListStyleInfo,
  type NumberingDef,
} from '@genoffice/docx-engine'
import type { PendingNumbering } from './doc-state'
import type { StringKey } from './i18n/locale'

export const NO_LIST_STYLE_ID = 'NoList'

export interface ListStyleEntry {
  styleId: string
  /** w:name written when the style has to be created */
  name: string
  /** Word lists its built-in list styles under these names, not their w:name */
  label?: string
  labelKey?: StringKey
  levels: CustomNumberingLevel[]
  /** the document's styles.xml has it */
  inDocument: boolean
  /** made by a user, in this document or this session: only these can be deleted */
  custom?: boolean
  /** the document names the style without giving it levels: there is no definition to modify */
  numberless?: boolean
}

const nine = (make: (i: number) => CustomNumberingLevel) =>
  Array.from({ length: 9 }, (_, i) => make(i))
const ROTATION = ['decimal', 'lowerLetter', 'lowerRoman']
const LEGAL_LEFT = [360, 792, 1224, 1728, 2232, 2736, 3240, 3744, 4320]
const LEGAL_HANGING = [360, 432, 504, 648, 792, 936, 1080, 1224, 1440]

function articleLevel(i: number): CustomNumberingLevel {
  if (i === 0) return { numFmt: 'upperRoman', lvlText: 'Article %1.', indentLeft: 0, hanging: 0 }
  if (i === 1)
    return {
      numFmt: 'decimalZero',
      lvlText: 'Section %1.%2',
      indentLeft: 0,
      hanging: 0,
      isLgl: true,
    }
  const numFmt = ['lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman'][(i - 2) % 5]
  const roman = numFmt === 'lowerRoman'
  return {
    numFmt,
    lvlText: i < 4 ? `(%${i + 1})` : i < 7 ? `%${i + 1})` : `%${i + 1}.`,
    indentLeft: 720 + 144 * (i - 2),
    hanging: roman ? (i === 6 ? 288 : 144) : 432,
    ...(roman ? { lvlJc: 'right' as const } : {}),
  }
}

/** Word's built-in list styles, in the order its dialog lists them; ids and names are Word's own */
export const BUILTIN_LIST_STYLES: Array<Omit<ListStyleEntry, 'inDocument'>> = [
  { styleId: NO_LIST_STYLE_ID, name: 'No List', labelKey: 'ribbonBnListStyleNoList', levels: [] },
  {
    styleId: '111111',
    name: 'Outline List 2',
    label: '1 / 1.1 / 1.1.1',
    levels: nine((i) => ({
      numFmt: 'decimal',
      lvlText: `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}.`,
      indentLeft: LEGAL_LEFT[i],
      hanging: LEGAL_HANGING[i],
    })),
  },
  {
    styleId: '1ai',
    name: 'Outline List 1',
    label: '1 / a / i',
    levels: nine((i) => ({
      numFmt: ROTATION[i % 3],
      lvlText: i < 3 ? `%${i + 1})` : i < 6 ? `(%${i + 1})` : `%${i + 1}.`,
      indentLeft: 360 * (i + 1),
      hanging: 360,
    })),
  },
  {
    styleId: 'ArticleSection',
    name: 'Outline List 3',
    labelKey: 'ribbonBnListStyleArticle',
    levels: nine(articleLevel),
  },
]

/** a list style made in the dialog and not saved yet */
export interface PendingListStyle {
  styleId: string
  name: string
  levels: CustomNumberingLevel[]
  /** one of Word's own list styles, given levels of its own before the document has it */
  builtIn?: true
}

/** what this session changed on list styles and has not saved yet */
export interface ListStyleEdits {
  /** levels rewritten on a definition the document holds, by style and level */
  levels: Map<string, Map<number, CustomNumberingLevel>>
  /** whole definitions of styles the document does not have yet */
  defined: Map<string, CustomNumberingLevel[]>
  names: Map<string, string>
}

export function listStyleEdits(
  parsed: { numbering: Map<string, NumberingDef>; listStyles?: ListStyleInfo[] },
  pending: PendingNumbering,
  names: Map<string, string> = new Map(),
): ListStyleEdits {
  const levels = new Map<string, Map<number, CustomNumberingLevel>>()
  for (const style of parsed.listStyles ?? []) {
    const abstractNumId =
      style.numId === undefined ? undefined : parsed.numbering.get(style.numId)?.abstractNumId
    const edits = pending.levelEdits.filter((e) => e.abstractNumId === abstractNumId)
    if (edits.length > 0) levels.set(style.styleId, new Map(edits.map((e) => [e.ilvl, e.level])))
  }
  const defined = new Map<string, CustomNumberingLevel[]>()
  for (const link of pending.styleNums)
    if (link.define) defined.set(link.styleId, link.define.levels)
  return { levels, defined, names }
}

/** The tab's list: the built-ins first (the document's own definition wins), then the document's other list styles by name */
export function listStyleEntries(
  documentStyles: ListStyleInfo[] = [],
  added: PendingListStyle[] = [],
  removed: string[] = [],
  edits?: ListStyleEdits,
): ListStyleEntry[] {
  const byId = new Map(
    documentStyles.filter((s) => !removed.includes(s.styleId)).map((s) => [s.styleId, s]),
  )
  const levelsOf = (s: ListStyleInfo) => {
    const edited = edits?.levels.get(s.styleId)
    return Object.keys(s.levels)
      .map(Number)
      .sort((a, b) => a - b)
      .map((i) => edited?.get(i) ?? customLevelFromNumberingLevel(s.levels[i]))
  }
  const nameOf = (s: ListStyleInfo) => edits?.names.get(s.styleId) ?? s.name
  const redefined = new Map(added.filter((s) => s.builtIn).map((s) => [s.styleId, s.levels]))
  const builtins = BUILTIN_LIST_STYLES.map((b) => {
    const own = byId.get(b.styleId)
    byId.delete(b.styleId)
    if (!own) {
      const levels = edits?.defined.get(b.styleId) ?? redefined.get(b.styleId) ?? b.levels
      return { ...b, levels, inDocument: false }
    }
    return {
      ...b,
      name: own.name,
      inDocument: true,
      levels: own.numId ? levelsOf(own) : b.levels,
      ...(own.numId || b.levels.length === 0 ? {} : { numberless: true }),
    }
  })
  const others = [...byId.values()]
    // a second numberless style would only be another "No List"
    .filter((s) => s.numId !== undefined)
    .map((s) => ({
      styleId: s.styleId,
      name: nameOf(s),
      label: nameOf(s),
      levels: levelsOf(s),
      inDocument: true,
      custom: true,
    }))
  const pending = added
    .filter((s) => !s.builtIn)
    .map((s) => ({
      ...s,
      levels: edits?.defined.get(s.styleId) ?? s.levels,
      label: s.name,
      inDocument: false,
      custom: true,
    }))
  return [...builtins, ...[...others, ...pending].sort((a, b) => a.name.localeCompare(b.name))]
}

/** Word's style id for a name: the name without what an id cannot hold, made unique */
export function newListStyleId(name: string, taken: (styleId: string) => boolean): string {
  const base = name.replace(/[^A-Za-z0-9]/g, '') || 'ListStyle'
  let id = base
  for (let n = 1; taken(id); n++) id = `${base}${n}`
  return id
}

/** The marker of each level as the page draws it, one item per level nested under the one before */
export function listStylePreview(levels: CustomNumberingLevel[]): string[] {
  const def = {
    numId: 'preview',
    abstractNumId: 'preview',
    levels: Object.fromEntries(levels.map((l, i) => [i, numberingLevelFromCustom(l)])),
    startOverrides: {},
  }
  return computeListMarkers(
    levels.map((_, ilvl) => ({ numId: def.numId, ilvl })),
    new Map([[def.numId, def]]),
  ).map((marker) => marker ?? '')
}

/** the nine levels a list definition has, the missing ones as Word's plain decimal levels */
export function nineListLevels(levels: CustomNumberingLevel[]): CustomNumberingLevel[] {
  const out = [...levels]
  while (out.length < 9)
    out.push({
      numFmt: 'decimal',
      lvlText: `%${out.length + 1}.`,
      indentLeft: 720 * (out.length + 1),
      hanging: 360,
    })
  return out.slice(0, 9)
}

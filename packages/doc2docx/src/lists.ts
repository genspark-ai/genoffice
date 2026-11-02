/**
 * Stage 3c-4: Word 97-2003 list/numbering tables.
 *
 * PlfLst (FIB pair 73) holds cLst 28-byte LSTF records; the LVL array follows
 * them (it is not counted in lcbPlfLst) and each LSTF owns one level
 * (fSimpleList) or nine. A LVL is LVLF (28 bytes) + grpprlPapx + grpprlChpx +
 * Xst; the Xst characters 0x0000-0x0008 are level-number placeholders and the
 * bullet levels carry a symbol-font glyph. PlfLfo (pair 74) maps 1-based ilfo
 * values to lsids through its LFO records and its LFOData carries the per-level
 * overrides (restart value and/or a replacement LVL). Malformed tables degrade: a broken
 * level grpprl keeps the level without its indents, a truncated level drops
 * the rest of the list definitions, a bad table range drops that table; every
 * case is reported in `warnings`.
 */
import type { FibInfo } from './fib'
import { iterateSprmsLenient, readSprmI16 } from './sprm'

const LSTF_BYTES = 28
const LVLF_BYTES = 28
const MAX_LISTS = 4096
const MAX_LEVEL_TEXT = 256

/* paragraph sprms that carry the level indents */
const SPRM_P_DXA_LEFT = 0x845e
const SPRM_P_DXA_LEFT_OLD = 0x840f
const SPRM_P_DXA_LEFT1 = 0x8460
const SPRM_P_DXA_LEFT1_OLD = 0x8411
const SPRM_P_DXA_RIGHT = 0x845d
const SPRM_P_DXA_RIGHT_OLD = 0x840e

export interface ListLevel {
  /** sprmPIlvl start value (iStartAt) */
  startAt: number
  /** raw nfc numbering format (0 decimal, 23 bullet, ...) */
  numberFormat: number
  alignment: 'left' | 'center' | 'right'
  legal: boolean
  noRestart: boolean
  /** level text with number placeholders as %1..%9 (bullet glyphs kept raw) */
  levelText: string
  /** how the number is separated from the text (w:suff) */
  follow: 'tab' | 'space' | 'nothing'
  indentLeftTwips?: number
  indentRightTwips?: number
  /** positive = first line indent, negative = hanging indent */
  indentFirstLineTwips?: number
  /** raw grpprlPapx / grpprlChpx of the level (kept for the writer) */
  paragraphSprms: Uint8Array
  characterSprms: Uint8Array
}

export interface ListDefinition {
  lsid: number
  simple: boolean
  levels: ListLevel[]
}

/** One LFOLVL: a level of an override that restarts and/or reformats the base level. */
export interface ListLevelOverride {
  ilvl: number
  startAt?: number
  level?: ListLevel
}

export interface ListTables {
  /** lsid -> definition */
  lists: Map<number, ListDefinition>
  /** 1-based ilfo -> lsid */
  overrides: Map<number, number>
  /** 1-based ilfo -> its LFOLVL overrides (only ilfos that carry any) */
  levelOverrides: Map<number, ListLevelOverride[]>
  warnings: string[]
}

/** The level a paragraph at (ilfo, ilvl) is numbered by: the override's level, else the base one. */
export function resolveListLevel(
  tables: ListTables,
  ilfo: number,
  ilvl: number,
): ListLevel | undefined {
  const overridden = tables.levelOverrides.get(ilfo)?.find((entry) => entry.ilvl === ilvl)?.level
  if (overridden) return overridden
  const lsid = tables.overrides.get(ilfo)
  return lsid === undefined ? undefined : tables.lists.get(lsid)?.levels[ilvl]
}

function u16At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(offset, true)
}

function u32At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, true)
}

const ALIGNMENTS: Array<'left' | 'center' | 'right'> = ['left', 'center', 'right', 'left']

function parseLevel(
  table: Uint8Array,
  offset: number,
  listIndex: number,
  levelIndex: number,
  warnings: string[],
): { level: ListLevel; next: number } | string {
  const where =
    listIndex < 0 ? 'override level ' + levelIndex : 'level ' + levelIndex + ' of list ' + listIndex
  if (offset + LVLF_BYTES > table.length) return 'the ' + where + ' is truncated (no LVLF)'
  const startAt = u32At(table, offset)
  const numberFormat = table[offset + 4]!
  const info = table[offset + 5]!
  const ixchFollow = table[offset + 0x0f]!
  // LVLF: 0x18 cbGrpprlChpx, 0x19 cbGrpprlPapx; the data after it is grpprlPapx then grpprlChpx
  const cbChpx = table[offset + 0x18]!
  const cbPapx = table[offset + 0x19]!
  const papxStart = offset + LVLF_BYTES
  const chpxStart = papxStart + cbPapx
  const xstStart = chpxStart + cbChpx
  if (chpxStart > table.length || xstStart + 2 > table.length)
    return 'the ' + where + ' overruns the table stream'
  const cch = u16At(table, xstStart)
  if (cch > MAX_LEVEL_TEXT) return 'the ' + where + ' has an oversized level text'
  const xstEnd = xstStart + 2 + cch * 2
  if (xstEnd > table.length) return 'the level text of the ' + where + ' overruns the table stream'
  let levelText = ''
  for (let index = 0; index < cch; index++) {
    const code = u16At(table, xstStart + 2 + index * 2)
    if (code <= 8) levelText += '%' + (code + 1)
    else levelText += String.fromCharCode(code)
  }
  let indentLeftTwips: number | undefined
  let indentRightTwips: number | undefined
  let indentFirstLineTwips: number | undefined
  const paragraphSprms = table.slice(papxStart, papxStart + cbPapx)
  const read = iterateSprmsLenient(paragraphSprms)
  if (read.truncated !== undefined)
    warnings.push(
      'the paragraph properties of the ' +
        where +
        ' are invalid (' +
        read.truncated +
        '); the properties after that point were ignored',
    )
  for (const sprm of read.sprms) {
    const value = readSprmI16(sprm.operand)
    if (value === undefined) continue
    if (sprm.opcode === SPRM_P_DXA_LEFT || sprm.opcode === SPRM_P_DXA_LEFT_OLD)
      indentLeftTwips = value
    else if (sprm.opcode === SPRM_P_DXA_LEFT1 || sprm.opcode === SPRM_P_DXA_LEFT1_OLD)
      indentFirstLineTwips = value
    else if (sprm.opcode === SPRM_P_DXA_RIGHT || sprm.opcode === SPRM_P_DXA_RIGHT_OLD)
      indentRightTwips = value
  }
  const follow: ListLevel['follow'] =
    ixchFollow === 1 ? 'space' : ixchFollow === 2 ? 'nothing' : 'tab'
  return {
    level: {
      startAt: startAt === 0 ? 1 : startAt,
      numberFormat,
      alignment: ALIGNMENTS[info & 0x03] ?? 'left',
      legal: (info & 0x04) !== 0,
      noRestart: (info & 0x08) !== 0,
      levelText,
      follow,
      ...(indentLeftTwips !== undefined ? { indentLeftTwips } : {}),
      ...(indentRightTwips !== undefined ? { indentRightTwips } : {}),
      ...(indentFirstLineTwips !== undefined ? { indentFirstLineTwips } : {}),
      paragraphSprms,
      characterSprms: table.slice(chpxStart, chpxStart + cbChpx),
    },
    next: xstEnd,
  }
}

/** Parse the PlfLst / PlfLfo tables of a .doc. Missing tables mean no lists. */
export function parseListTables(fib: FibInfo, table: Uint8Array): ListTables {
  const warnings: string[] = []
  const lists = new Map<number, ListDefinition>()
  const overrides = new Map<number, number>()
  const dropLists = (detail: string): void => {
    warnings.push('the list definitions are invalid (' + detail + '); list formatting was omitted')
  }
  const parseLists = (): void => {
    if (fib.fcPlfLst > table.length || fib.lcbPlfLst > table.length - fib.fcPlfLst)
      return dropLists('the list table range lies outside the table stream')
    if (fib.lcbPlfLst < 2) return dropLists('the list table is too small for a cLst count')
    const cLst = u16At(table, fib.fcPlfLst)
    if (cLst > MAX_LISTS) return dropLists('the list table declares ' + cLst + ' lists')
    const headerEnd = fib.fcPlfLst + 2 + cLst * LSTF_BYTES
    if (headerEnd > table.length)
      return dropLists('the ' + cLst + ' list descriptors overrun the table stream')
    if (headerEnd > fib.fcPlfLst + fib.lcbPlfLst)
      warnings.push(
        'the list table length does not cover its descriptors; the trailing level data is read anyway',
      )
    let offset = headerEnd
    for (let listIndex = 0; listIndex < cLst; listIndex++) {
      const base = fib.fcPlfLst + 2 + listIndex * LSTF_BYTES
      const lsid = u32At(table, base)
      const simple = (table[base + 0x1a]! & 0x01) !== 0
      const levelCount = simple ? 1 : 9
      const levels: ListLevel[] = []
      for (let levelIndex = 0; levelIndex < levelCount; levelIndex++) {
        const parsed = parseLevel(table, offset, listIndex, levelIndex, warnings)
        if (typeof parsed === 'string') {
          // the level boundaries are lost from here on: keep what was read so far
          warnings.push(parsed + '; this list and the list definitions after it were dropped')
          return
        }
        levels.push(parsed.level)
        offset = parsed.next
      }
      lists.set(lsid, { lsid, simple, levels })
    }
  }
  const levelOverrides = new Map<number, ListLevelOverride[]>()
  const parseOverrides = (): void => {
    const drop = (detail: string): void => {
      warnings.push(
        'the list override table is invalid (' + detail + '); list formatting was omitted',
      )
    }
    if (fib.fcPlfLfo > table.length || fib.lcbPlfLfo > table.length - fib.fcPlfLfo)
      return drop('the range lies outside the table stream')
    if (fib.lcbPlfLfo < 4) return drop('it is too small for an lfoMac count')
    const lfoMac = u32At(table, fib.fcPlfLfo)
    if (lfoMac > MAX_LISTS) return drop('it declares ' + lfoMac + ' entries')
    if (4 + lfoMac * 16 > fib.lcbPlfLfo)
      return drop('the ' + lfoMac + ' list overrides overrun the table')
    const levelCounts: number[] = []
    for (let index = 0; index < lfoMac; index++) {
      const base = fib.fcPlfLfo + 4 + index * 16
      overrides.set(index + 1, u32At(table, base))
      levelCounts.push(table[base + 0x0c]!)
    }
    // LFOData follows the LFO array: a cp, then clfolvl LFOLVLs, each optionally followed by a LVL
    const end = Math.min(table.length, fib.fcPlfLfo + fib.lcbPlfLfo)
    let offset = fib.fcPlfLfo + 4 + lfoMac * 16
    for (let index = 0; index < lfoMac; index++) {
      const ilfo = index + 1
      offset += 4
      const entries: ListLevelOverride[] = []
      for (let count = 0; count < levelCounts[index]!; count++) {
        if (offset + 8 > end) {
          warnings.push(
            'the level overrides of list override ' +
              ilfo +
              ' are truncated; the remaining overrides use their base lists',
          )
          if (entries.length > 0) levelOverrides.set(ilfo, entries)
          return
        }
        const startAt = u32At(table, offset)
        const flags = table[offset + 4]!
        offset += 8
        const entry: ListLevelOverride = { ilvl: flags & 0x0f }
        if (flags & 0x10) entry.startAt = startAt
        if (flags & 0x20) {
          const parsed = parseLevel(table, offset, -1, entry.ilvl, warnings)
          if (typeof parsed === 'string') {
            warnings.push(
              'the level override of list override ' +
                ilfo +
                ' is invalid (' +
                parsed +
                '); the remaining overrides use their base lists',
            )
            if (entries.length > 0) levelOverrides.set(ilfo, entries)
            return
          }
          entry.level = parsed.level
          offset = parsed.next
        }
        entries.push(entry)
      }
      if (entries.length > 0) levelOverrides.set(ilfo, entries)
    }
  }
  if (fib.lcbPlfLst > 0) parseLists()
  if (fib.lcbPlfLfo > 0 && fib.fcPlfLfo > 0) parseOverrides()
  return { lists, overrides, levelOverrides, warnings: [...new Set(warnings)] }
}

/** OOXML numFmt value for a Word nfc (fallback decimal). */
export const NUMBER_FORMATS: Record<number, string> = {
  0: 'decimal',
  1: 'upperRoman',
  2: 'lowerRoman',
  3: 'upperLetter',
  4: 'lowerLetter',
  5: 'ordinal',
  6: 'cardinalText',
  7: 'ordinalText',
  8: 'hex',
  9: 'chicago',
  10: 'ideographDigital',
  11: 'japaneseCounting',
  12: 'aiueo',
  13: 'iroha',
  14: 'decimalFullWidth',
  15: 'decimalHalfWidth',
  16: 'japaneseLegal',
  17: 'japaneseDigitalTenThousand',
  18: 'decimalEnclosedCircle',
  19: 'decimalFullWidth2',
  20: 'japaneseLeadingZero',
  23: 'bullet',
  24: 'picture',
}

/** the symbol-font bullet glyphs Word writes into the Xst, mapped to text bullets */
const BULLET_GLYPHS: Record<number, string> = {
  0xf0b7: '•',
  0xf0a7: '▪',
  0xf0d8: '➢',
  0xf0fc: '✓',
  0xf0e0: '→',
  0xf0a8: '◆',
  0xf06e: '■',
  0xf06c: '●',
  0xf0b8: '○',
}

/** true when the level is a bullet level in OOXML terms */
export function isBulletLevel(level: ListLevel): boolean {
  return level.numberFormat === 23 || level.numberFormat === 24
}

/** the level text the writer should put into w:lvlText */
export function levelTextForWriter(level: ListLevel): string {
  if (!isBulletLevel(level)) return level.levelText
  const first = level.levelText[0]
  if (first === undefined) return '•'
  const mapped = BULLET_GLYPHS[first.charCodeAt(0)]
  return mapped ?? (first.charCodeAt(0) < 0x20 ? '•' : first)
}

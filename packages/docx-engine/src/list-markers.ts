import { decodeSymbolText, isSymbolFont, toSymbolPua } from './symbol-fonts'
import type { NumberingDef, NumberingLevel } from './types'

/** Word represents bullets with Symbol/Wingdings private-use characters; map them to common glyphs */
const BULLET_GLYPHS: Record<string, string> = {
  '': '•',
  '': '▪',
  '': '➢',
  '': '❖',
  '': '✓',
  o: '◦',
}

const DEFAULT_BULLETS = ['•', '◦', '▪', '•', '◦', '▪', '•', '◦', '▪']

function toLetters(value: number): string {
  // Word: 1..26 -> A..Z, 27 -> AA (repeated same letter, not positional notation)
  const n = ((value - 1) % 26) + 1
  const repeat = Math.floor((value - 1) / 26) + 1
  return String.fromCharCode(64 + n).repeat(repeat)
}

function toGreek(value: number, base: number): string {
  if (value < 1) return String(value)
  // 24-letter alphabet (no final sigma); 25 -> αα, repeated like toLetters
  const n = ((value - 1) % 24) + 1
  const repeat = Math.floor((value - 1) / 24) + 1
  // skip the final-sigma slot (ς / unassigned) between the 17th and 18th letters
  return String.fromCharCode(base + n - 1 + (n >= 18 ? 1 : 0)).repeat(repeat)
}

const ROMAN: Array<[number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
]

function toRoman(value: number): string {
  let n = Math.max(1, value)
  let out = ''
  for (const [v, s] of ROMAN) {
    while (n >= v) {
      out += s
      n -= v
    }
  }
  return out
}

const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
const CN_UNITS = ['', '\u5341', '\u767e', '\u5343']
const CN_LEGAL_SIMPLIFIED = [
  '\u96f6',
  '\u58f9',
  '\u8d30',
  '\u53c1',
  '\u8086',
  '\u4f0d',
  '\u9646',
  '\u67d2',
  '\u634c',
  '\u7396',
]
const CN_LEGAL_TRADITIONAL = [
  '\u96f6',
  '\u58f9',
  '\u8cb3',
  '\u53c3',
  '\u8086',
  '\u4f0d',
  '\u9678',
  '\u67d2',
  '\u634c',
  '\u7396',
]
const CN_LEGAL_UNITS = ['', '\u62fe', '\u4f70', '\u4edf']

function toCjkCounting(
  value: number,
  digits: string[],
  units: string[],
  dropLeadingOne: boolean,
): string {
  if (value <= 0 || value > 9999) return String(value)
  if (value < 10) return digits[value]
  const zero = digits[0]
  const parts = String(value).split('').map(Number)
  let out = ''
  for (let i = 0; i < parts.length; i++) {
    const d = parts[i]
    const unit = units[parts.length - 1 - i]
    if (d === 0) {
      if (!out.endsWith(zero) && i < parts.length - 1) out += zero
    } else {
      out += digits[d] + unit
    }
  }
  out = out.replace(new RegExp(`${zero}+$`), '')
  // 10-19 drop the leading "one" digit: read as "shi x" rather than "yi shi x"
  return dropLeadingOne ? out.replace(new RegExp(`^${digits[1]}${units[1]}`), units[1]) : out
}

function toChinese(value: number): string {
  return toCjkCounting(value, CN_DIGITS, CN_UNITS, true)
}

/** digit-by-digit transliteration (Word "digital" formats) */
function digitWise(value: number, digits: string): string {
  if (value < 0) return String(value)
  const table = Array.from(digits)
  return String(value)
    .split('')
    .map((d) => table[Number(d)])
    .join('')
}

/** 1..n -> the nth code point of an alphabet, else the decimal fallback */
function fromAlphabet(value: number, alphabet: string, fallback = String(value)): string {
  const letters = Array.from(alphabet)
  return value >= 1 && value <= letters.length ? letters[value - 1] : fallback
}

/** enclosed digit runs (U+2460 circled, U+2474 parenthesized, U+2488 with full stop): 1..20 */
function enclosed(value: number, base: number, count = 20): string {
  return value >= 1 && value <= count ? String.fromCodePoint(base + value - 1) : String(value)
}

const FW_DIGITS = '０１２３４５６７８９'
const IDEOGRAPH_DIGITS = '\u3007\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d'
const AIUEO_FW =
  'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン'
const AIUEO_HW = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ'
const IROHA_FW =
  'イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス'
const GANADA = '가나다라마바사아자차카타파하'
const CHOSUNG = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ'
const RUSSIAN_LOWER = 'абвгдежзиклмнопрстуфхцчшщэюя'
const HEAVENLY_STEMS = '\u7532\u4e59\u4e19\u4e01\u620a\u5df1\u5e9a\u8f9b\u58ec\u7678'
const EARTHLY_BRANCHES = '\u5b50\u4e11\u5bc5\u536f\u8fb0\u5df3\u5348\u672a\u7533\u9149\u620c\u4ea5'

const ONES = [
  '',
  'One',
  'Two',
  'Three',
  'Four',
  'Five',
  'Six',
  'Seven',
  'Eight',
  'Nine',
  'Ten',
  'Eleven',
  'Twelve',
  'Thirteen',
  'Fourteen',
  'Fifteen',
  'Sixteen',
  'Seventeen',
  'Eighteen',
  'Nineteen',
]
const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
const ORDINAL_IRREGULAR: Record<string, string> = {
  One: 'First',
  Two: 'Second',
  Three: 'Third',
  Five: 'Fifth',
  Eight: 'Eighth',
  Nine: 'Ninth',
  Twelve: 'Twelfth',
}

function cardinalWords(value: number): string {
  if (value < 1 || value > 999999) return String(value)
  if (value < 20) return ONES[value]
  if (value < 100) return TENS[Math.floor(value / 10)] + (value % 10 ? `-${ONES[value % 10]}` : '')
  if (value < 1000)
    return `${ONES[Math.floor(value / 100)]} Hundred${value % 100 ? ` ${cardinalWords(value % 100)}` : ''}`
  return `${cardinalWords(Math.floor(value / 1000))} Thousand${value % 1000 ? ` ${cardinalWords(value % 1000)}` : ''}`
}

function ordinalWords(value: number): string {
  const words = cardinalWords(value)
  if (!/[A-Za-z]$/.test(words)) return words
  return words.replace(/[A-Za-z]+$/, (last) => {
    if (ORDINAL_IRREGULAR[last]) return ORDINAL_IRREGULAR[last]
    if (last.endsWith('y')) return `${last.slice(0, -1)}ieth`
    return `${last}th`
  })
}

function ordinalSuffix(value: number): string {
  const mod100 = value % 100
  if (mod100 >= 11 && mod100 <= 13) return `${value}th`
  const suffix = ['th', 'st', 'nd', 'rd'][value % 10] ?? 'th'
  return `${value}${suffix}`
}

/** w14 custom numFmt enumeration ("α, β, γ, ..."): comma-separated items, a trailing "..." marks continuation */
export function customEnumItems(format: string): string[] | null {
  const items = format.split(',').map((s) => s.trim())
  while (items.length > 0 && /^(\.{3}|…)?$/.test(items[items.length - 1])) items.pop()
  return items.length >= 2 && items.every(Boolean) ? items : null
}

export function formatNumber(rawValue: number, numFmt: string, customFormat?: string): string {
  // Numbering values arrive from the file (w:start/w:val): a hostile 1e9
  // would loop toRoman ~1M times and build a 38MB toLetters repeat string,
  // and Infinity never terminates. Bound once at the choke point.
  const value = Number.isFinite(rawValue) ? Math.min(Math.floor(rawValue), 999_999) : 0
  if (numFmt === 'custom') {
    const items = customFormat ? customEnumItems(customFormat) : null
    // enumeration exhausted: cycle (best-effort; Word's continuation rules are undocumented)
    return items && value >= 1 ? items[(value - 1) % items.length] : String(value)
  }
  switch (numFmt) {
    case 'decimalZero':
      return value < 10 && value >= 0 ? `0${value}` : String(value)
    case 'lowerLetter':
      return toLetters(value).toLowerCase()
    case 'upperLetter':
      return toLetters(value)
    case 'lowerRoman':
      return toRoman(value).toLowerCase()
    case 'upperRoman':
      return toRoman(value)
    case 'lowerGreek':
      return toGreek(value, 0x3b1)
    case 'upperGreek':
      return toGreek(value, 0x391)
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
    case 'taiwaneseCounting':
    case 'taiwaneseCountingThousand':
      return toChinese(value)
    case 'chineseLegalSimplified':
      return toCjkCounting(value, CN_LEGAL_SIMPLIFIED, CN_LEGAL_UNITS, false)
    case 'ideographLegalTraditional':
      return toCjkCounting(value, CN_LEGAL_TRADITIONAL, CN_LEGAL_UNITS, false)
    case 'ideographDigital':
    case 'japaneseDigitalTenThousand':
    case 'taiwaneseDigital':
      return digitWise(value, IDEOGRAPH_DIGITS)
    case 'koreanDigital':
      return digitWise(value, '영일이삼사오육칠팔구')
    case 'koreanDigital2':
      return digitWise(value, IDEOGRAPH_DIGITS)
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return digitWise(value, FW_DIGITS)
    case 'thaiNumbers':
      return digitWise(value, '๐๑๒๓๔๕๖๗๘๙')
    case 'hindiNumbers':
      return digitWise(value, '०१२३४५६७८९')
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return enclosed(value, 0x2460)
    case 'decimalEnclosedParen':
      return enclosed(value, 0x2474)
    case 'decimalEnclosedFullstop':
      return enclosed(value, 0x2488)
    case 'ideographEnclosedCircle':
      return enclosed(value, 0x3280, 10)
    case 'ideographTraditional':
      return fromAlphabet(value, HEAVENLY_STEMS)
    case 'ideographZodiac':
      return fromAlphabet(value, EARTHLY_BRANCHES)
    case 'aiueo':
      return fromAlphabet(value, AIUEO_HW)
    case 'aiueoFullWidth':
      return fromAlphabet(value, AIUEO_FW)
    case 'irohaFullWidth':
      return fromAlphabet(value, IROHA_FW)
    case 'ganada':
      return fromAlphabet(value, GANADA)
    case 'chosung':
      return fromAlphabet(value, CHOSUNG)
    case 'russianLower':
      return fromAlphabet(value, RUSSIAN_LOWER)
    case 'russianUpper':
      return fromAlphabet(value, RUSSIAN_LOWER.toUpperCase())
    case 'ordinal':
      return ordinalSuffix(value)
    case 'cardinalText':
      return cardinalWords(value)
    case 'ordinalText':
      return ordinalWords(value)
    case 'numberInDash':
      return `- ${value} -`
    case 'none':
      return ''
    default:
      return String(value)
  }
}

/** w:isLgl leaves Arabic-digit formats alone and turns every other one into decimal */
const ARABIC_FORMATS = new Set(['decimal', 'decimalZero'])

export interface ListItemRef {
  numId: string | null
  ilvl: number
  /** paragraph mark is a tracked deletion: Word merges it into the next paragraph, so it
   *  shows that paragraph's number without consuming one */
  deleted?: boolean
}

export interface ListMarkerInfo {
  /** display text (symbol glyphs decoded to Unicode) */
  text: string
  /** bullets declared in a symbol-encoded font: original glyph in U+F0xx form, so the renderer can pass it through when the font is installed */
  symbolChar?: string
  symbolFont?: string
  /** picture bullet (w:lvlPicBulletId): data URL drawn in place of the text */
  picBulletSrc?: string
  /** literal bullet text declared in an ordinary text font (Word's "o" in Courier New) */
  font?: string
  /** the item's own counter value (numbered levels only) */
  value?: number
}

/** text-font substitutes draw solid round bullets smaller than the Word symbol glyph
 *  (glyph-box ratios: Symbol bullet 0.29em vs Arial 0.25em; Wingdings circle 0.58em vs Arial 0.43em) */
const SUBSTITUTE_BULLET_SCALE: Record<string, number> = { '•': 1.25, '●': 1.35 }

export function bulletMarkerScale(glyph: string): number {
  return SUBSTITUTE_BULLET_SCALE[glyph] ?? 1
}

function bulletInfo(text: string, level: NumberingLevel): ListMarkerInfo {
  const font = level.font?.trim()
  if (!font || !isSymbolFont(font)) return { text }
  return { text, symbolChar: toSymbolPua(level.lvlText), symbolFont: font }
}

/**
 * Word numbering semantics: counters accumulate document-wide per abstractNum
 * (plain paragraphs in between don't break the sequence); when a level appears,
 * its deeper levels reset; startOverride applies the first time that numId uses
 * the level. Returns a marker array the same length as items; items without a
 * definition return null (CSS fallback handles them).
 */
export function computeListMarkerInfos(
  items: ListItemRef[],
  defs: Map<string, NumberingDef>,
): (ListMarkerInfo | null)[] {
  const counters = new Map<string, number[]>()
  const overrideApplied = new Set<string>()
  return items.map((item) => {
    const def = item.numId !== null ? defs.get(item.numId) : undefined
    if (!def) return null
    const lvl = Math.max(0, item.ilvl)
    const level = def.levels[lvl]
    if (!level) return null

    if (level.numFmt === 'bullet') {
      // the lvlText of a picture bullet is a placeholder glyph, never shown
      if (level.picBulletId !== undefined) {
        return level.picBulletSrc
          ? { text: '', picBulletSrc: level.picBulletSrc }
          : { text: DEFAULT_BULLETS[0] }
      }
      // symbol-encoded marker fonts (Wingdings "l" = ●): decode by font first
      const decoded = level.font ? decodeSymbolText(level.font, level.lvlText) : null
      const textFont = level.font?.trim()
      if (textFont && !isSymbolFont(textFont) && /^[^\s\uf000-\uf0ff]+$/.test(level.lvlText))
        return { text: level.lvlText, font: textFont }
      const glyph = decoded ?? BULLET_GLYPHS[level.lvlText] ?? level.lvlText
      // undecodable symbol-font glyphs, private-use or empty ones fall back to the per-level default
      if (!glyph || /[-]/.test(glyph) || (decoded === null && isSymbolFont(level.font)))
        return bulletInfo(DEFAULT_BULLETS[lvl % 9], level)
      return bulletInfo(glyph, level)
    }

    const live = counters.get(def.abstractNumId) ?? []
    const c = item.deleted ? [...live] : live
    if (!item.deleted) counters.set(def.abstractNumId, c)
    const applied = item.deleted ? new Set(overrideApplied) : overrideApplied
    // Word: an item at level L instantiates untouched shallower levels at
    // their start value, so a later explicit item there increments past it
    // ("1.1 Objetivos" before any level-0 item makes the first level-0 "2.")
    for (let a = 0; a < lvl; a++) {
      if (c[a] !== undefined) continue
      const aKey = `${def.numId}:${a}`
      if (def.startOverrides[a] !== undefined && !applied.has(aKey)) {
        applied.add(aKey)
        c[a] = def.startOverrides[a]
      } else c[a] = def.levels[a]?.start ?? 1
    }
    const overrideKey = `${def.numId}:${lvl}`
    if (def.startOverrides[lvl] !== undefined && !applied.has(overrideKey)) {
      applied.add(overrideKey)
      c[lvl] = def.startOverrides[lvl]
    } else {
      c[lvl] = (c[lvl] ?? level.start - 1) + 1
    }
    c.length = lvl + 1

    const marker = level.lvlText.replace(/%(\d)/g, (_, d: string) => {
      const refLvl = Number(d) - 1
      const refDef = def.levels[refLvl]
      const value = c[refLvl] ?? refDef?.start ?? 1
      const fmt = refDef?.numFmt ?? 'decimal'
      if (level.isLgl && !ARABIC_FORMATS.has(fmt)) return String(value)
      return formatNumber(value, fmt, refDef?.customFormat)
    })
    // numFmt "none": an explicit empty marker (Word shows nothing) so renderer
    // counter fallbacks don't kick in on the null
    if (!marker && level.numFmt !== 'none') return null
    return { text: marker, value: c[lvl] }
  })
}

/** Per item: the counter value its own level shows (1 for bullets and unresolved items) */
export function computeListValues(items: ListItemRef[], defs: Map<string, NumberingDef>): number[] {
  return computeListMarkerInfos(items, defs).map((m) => m?.value ?? 1)
}

export function computeListMarkers(
  items: ListItemRef[],
  defs: Map<string, NumberingDef>,
): (string | null)[] {
  return computeListMarkerInfos(items, defs).map((m) => m?.text ?? null)
}

/**
 * Word's default tab after a numbering marker: when the marker fits the hanging
 * area the text starts at the text indent; otherwise it jumps to the next tab
 * stop past the marker end — the paragraph's custom stops first, the default
 * grid only beyond the last custom stop (a custom stop clears the default stops
 * left of it). All positions are twips relative to the text column edge.
 * Returns the marker-box width (marker start -> text start), or null when the
 * hanging area already fits.
 */
export function markerTabAdvance(
  markerStart: number,
  markerWidth: number,
  textIndent: number,
  defaultTab = 720,
  customStops: number[] = [],
): number | null {
  const end = markerStart + markerWidth
  if (markerStart < textIndent && end <= textIndent) return null
  const custom = customStops.filter((s) => s > end).sort((a, b) => a - b)[0]
  if (custom !== undefined) return custom - markerStart
  // no default grid (w:defaultTabStop 0): the text follows the marker directly
  if (!(defaultTab > 0)) return markerWidth
  const floor = Math.max(end, ...customStops)
  const stop = (Math.floor(floor / defaultTab) + 1) * defaultTab
  return stop - markerStart
}

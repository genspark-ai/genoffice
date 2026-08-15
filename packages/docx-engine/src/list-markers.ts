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

function toChinese(value: number): string {
  if (value <= 0 || value > 9999) return String(value)
  if (value < 10) return CN_DIGITS[value]
  const units = ['', '十', '百', '千']
  const digits = String(value).split('').map(Number)
  let out = ''
  for (let i = 0; i < digits.length; i++) {
    const d = digits[i]
    const unit = units[digits.length - 1 - i]
    if (d === 0) {
      if (!out.endsWith('零') && i < digits.length - 1) out += '零'
    } else {
      out += CN_DIGITS[d] + unit
    }
  }
  out = out.replace(/零+$/, '')
  // 10-19 drop the leading "one" digit: read as "shi x" rather than "yi shi x"
  return out.replace(/^一十/, '十')
}

export function formatNumber(value: number, numFmt: string): string {
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
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
      return toChinese(value)
    case 'decimalEnclosedCircle':
      return value >= 1 && value <= 20 ? String.fromCodePoint(0x2460 + value - 1) : String(value)
    case 'none':
      return ''
    default:
      return String(value)
  }
}

export interface ListItemRef {
  numId: string | null
  ilvl: number
}

export interface ListMarkerInfo {
  /** display text (symbol glyphs decoded to Unicode) */
  text: string
  /** bullets declared in a symbol-encoded font: original glyph in U+F0xx form, so the renderer can pass it through when the font is installed */
  symbolChar?: string
  symbolFont?: string
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
      // symbol-encoded marker fonts (Wingdings "l" = ●): decode by font first
      const decoded = level.font ? decodeSymbolText(level.font, level.lvlText) : null
      const glyph = decoded ?? BULLET_GLYPHS[level.lvlText] ?? level.lvlText
      // undecodable symbol-font glyphs, private-use or empty ones fall back to the per-level default
      if (!glyph || /[-]/.test(glyph) || (decoded === null && isSymbolFont(level.font)))
        return bulletInfo(DEFAULT_BULLETS[lvl % 9], level)
      return bulletInfo(glyph, level)
    }

    const c = counters.get(def.abstractNumId) ?? []
    counters.set(def.abstractNumId, c)
    const overrideKey = `${def.numId}:${lvl}`
    if (def.startOverrides[lvl] !== undefined && !overrideApplied.has(overrideKey)) {
      overrideApplied.add(overrideKey)
      c[lvl] = def.startOverrides[lvl]
    } else {
      c[lvl] = (c[lvl] ?? level.start - 1) + 1
    }
    c.length = lvl + 1

    const marker = level.lvlText.replace(/%(\d)/g, (_, d: string) => {
      const refLvl = Number(d) - 1
      const refDef = def.levels[refLvl]
      const value = c[refLvl] ?? refDef?.start ?? 1
      return formatNumber(value, refDef?.numFmt ?? 'decimal')
    })
    return marker ? { text: marker } : null
  })
}

export function computeListMarkers(
  items: ListItemRef[],
  defs: Map<string, NumberingDef>,
): (string | null)[] {
  return computeListMarkerInfos(items, defs).map((m) => m?.text ?? null)
}

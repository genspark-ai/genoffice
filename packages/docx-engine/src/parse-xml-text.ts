// Text and unit primitives shared by every parse-* module: entity decoding,
// plain-text extraction, on/off and measurement attribute readers.
import { attrsOf, findChild, type XNode } from './xml-utils'
import { resolveThemeColor } from './theme'
import type { ThemeColors } from './types'

export interface RelInfo {
  target: string
  type: string
  targetMode?: string
}

/** hex color value tolerance: Word/LO accept a leading '#' (tdf#57589) */
export function stripHash(v: string): string {
  return v.startsWith('#') ? v.slice(1) : v
}

/** Word drops schema-invalid w:line values (e.g. floats from Google Docs) wholesale */
export function lineTwipsOf(v: string | undefined): number {
  return v !== undefined && /^-?\d+$/.test(v) ? parseInt(v, 10) : NaN
}

/** w:color -> display hex; w:themeColor resolves against the live palette (beats stale w:val) */
export function colorFrom(
  container: XNode | undefined,
  theme?: ThemeColors | null,
): string | undefined {
  if (!container) return undefined
  const a = attrsOf(findChild(container, 'w:color') ?? {})
  if (a['w:themeColor'] && theme) {
    const resolved = resolveThemeColor(
      a['w:themeColor'],
      theme,
      a['w:themeTint'],
      a['w:themeShade'],
    )
    if (resolved) return resolved
  }
  const val = a['w:val']
  return val && val !== 'auto' ? stripHash(val) : undefined
}

/** OOXML on/off toggle, three-state: absent → undefined, explicit off → false */
export function onOffOf(parent: XNode, name: string): boolean | undefined {
  const child = findChild(parent, name)
  if (!child) return undefined
  const val = attrsOf(child)['w:val']
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

export function plainText(xml: string): string {
  const texts: string[] = []
  // a space at cell boundaries keeps table text from gluing together
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<\/w:tc>/g
  let m: RegExpExecArray | null
  let pendingGap = false
  while ((m = re.exec(xml)) !== null) {
    if (m[0] === '</w:tc>') {
      pendingGap = texts.length > 0
      continue
    }
    if (pendingGap) {
      texts.push(' ')
      pendingGap = false
    }
    texts.push(m[1])
  }
  return decodeEntities(texts.join(''))
}

/** Visible OMML leaf tokens; editing these preserves the surrounding formula tree. */
export function mathTokens(xml: string): string[] {
  const tokens: string[] = []
  const re = /<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) tokens.push(decodeEntities(m[1]))
  return tokens
}

/**
 * Numeric character references only (`&#65;` / `&#xF0B7;`), which fast-xml-parser
 * leaves alone. Text and attribute values read off the parse tree have already had
 * the five named entities resolved, so they must NOT be run through the named-entity
 * replacements again: a document whose visible text is literally `&lt;` is stored as
 * `&amp;lt;`, and a second decode would turn it into `<`.
 */
export function decodeNumericCharRefs(text: string): string {
  return text.replace(
    /&#(?:x([0-9a-f]+)|([0-9]+));/gi,
    (entity, hex: string | undefined, decimal: string | undefined) => {
      const codePoint = parseInt(hex ?? decimal ?? '', hex ? 16 : 10)
      return Number.isFinite(codePoint) &&
        codePoint >= 0 &&
        codePoint <= 0x10ffff &&
        !(codePoint >= 0xd800 && codePoint <= 0xdfff)
        ? String.fromCodePoint(codePoint)
        : entity
    },
  )
}

/** Full decode, for raw XML slices that never passed through the parser. */
export function decodeEntities(text: string): string {
  return decodeNumericCharRefs(text)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

export const EMU_PER_PX = 9525

export const EMU_PER_PT = 12700

import type { ThemeColors, ThemeFonts } from './types'
import { escapeXmlAttr } from './xml-utils'

/**
 * Theme part support (word/theme/theme1.xml). Theme fonts rewrite the
 * major/minor font pair; theme colors rewrite the a:clrScheme entries. Documents
 * whose styles reference theme fonts/colors re-render in Word accordingly.
 */

export const THEME_PART_PATH = 'word/theme/theme1.xml'

export const THEME_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

export const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'

/** read the latin/ea/cs major+minor typefaces from theme1.xml */
export function readThemeFonts(themeXml: string): ThemeFonts | null {
  const major = fontOf(themeXml, 'a:majorFont')
  const minor = fontOf(themeXml, 'a:minorFont')
  if (!major && !minor) return null
  const slot = (tag: string, kind: 'a:ea' | 'a:cs') =>
    sectionOf(themeXml, tag)?.match(new RegExp(`<${kind} typeface="([^"]*)"`))?.[1] || undefined
  const eastAsia = slot('a:minorFont', 'a:ea')
  const majorEastAsia = slot('a:majorFont', 'a:ea')
  const minorCs = slot('a:minorFont', 'a:cs')
  const majorCs = slot('a:majorFont', 'a:cs')
  return {
    major: major ?? '',
    minor: minor ?? '',
    ...(eastAsia ? { eastAsia } : {}),
    ...(majorEastAsia ? { majorEastAsia } : {}),
    ...(minorCs ? { minorCs } : {}),
    ...(majorCs ? { majorCs } : {}),
  }
}

function sectionOf(xml: string, tag: string): string | null {
  return new RegExp(`<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`).exec(xml)?.[0] ?? null
}

function fontOf(xml: string, tag: string): string | null {
  const section = sectionOf(xml, tag)
  return section ? (/<a:latin typeface="([^"]*)"/.exec(section)?.[1] ?? null) : null
}

/** rewrite the latin (and optional east-asian) typefaces of one font group */
function applyFontGroup(xml: string, tag: string, typeface: string, eastAsia?: string): string {
  const section = sectionOf(xml, tag)
  if (!section) return xml
  let next = section.replace(/(<a:latin typeface=")[^"]*(")/, `$1${escapeXmlAttr(typeface)}$2`)
  if (eastAsia !== undefined) {
    next = next.replace(/(<a:ea typeface=")[^"]*(")/, `$1${escapeXmlAttr(eastAsia)}$2`)
  }
  return xml.replace(section, next)
}

export function applyThemeFonts(themeXml: string, fonts: ThemeFonts): string {
  let xml = applyFontGroup(themeXml, 'a:majorFont', fonts.major, fonts.eastAsia)
  xml = applyFontGroup(xml, 'a:minorFont', fonts.minor, fonts.eastAsia)
  return xml
}

const COLOR_TAGS = [
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const

/** all clrScheme slots, including read-only ones (dk1/lt1 may be sysClr) */
const READ_TAGS = ['dk1', 'lt1', ...COLOR_TAGS, 'hlink', 'folHlink'] as const

export function readThemeColors(themeXml: string): ThemeColors | null {
  const scheme = sectionOf(themeXml, 'a:clrScheme')
  if (!scheme) return null
  const out: ThemeColors = {}
  const name = /<a:clrScheme name="([^"]*)"/.exec(themeXml)?.[1]
  if (name) out.name = name
  for (const tag of READ_TAGS) {
    const m = new RegExp(
      `<a:${tag}>\\s*<a:(?:srgbClr val|sysClr[^>]*? lastClr)="([0-9A-Fa-f]{6})"`,
    ).exec(scheme)
    if (m) out[tag] = m[1].toUpperCase()
  }
  return Object.keys(out).length > 0 ? out : null
}

/** w:themeColor attribute value -> clrScheme slot (default clrSchemeMapping) */
const THEME_COLOR_SLOTS: Record<string, keyof ThemeColors> = {
  dark1: 'dk1',
  text1: 'dk1',
  light1: 'lt1',
  background1: 'lt1',
  dark2: 'dk2',
  text2: 'dk2',
  light2: 'lt2',
  background2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
}

const SLOT_FALLBACK: Partial<Record<keyof ThemeColors, string>> = { dk1: '000000', lt1: 'FFFFFF' }

/**
 * Resolve a w:themeColor reference (+ optional w:themeTint / w:themeShade,
 * hex 00-FF) against the palette. sRGB per-channel approximation of Word's
 * tint/shade math — close enough for display; w:val stays authoritative on save.
 */
export function resolveThemeColor(
  themeColor: string,
  colors: ThemeColors,
  tint?: string,
  shade?: string,
): string | null {
  const slot = THEME_COLOR_SLOTS[themeColor]
  if (!slot) return null
  const base = (colors[slot] as string | undefined) ?? SLOT_FALLBACK[slot]
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16))
  const factorOf = (hex?: string) => {
    const v = hex ? parseInt(hex, 16) : NaN
    return Number.isFinite(v) ? Math.max(0, Math.min(255, v)) / 255 : null
  }
  const s = factorOf(shade)
  if (s !== null) rgb = rgb.map((c) => c * s)
  const t = factorOf(tint)
  if (t !== null) rgb = rgb.map((c) => c * t + 255 * (1 - t))
  return rgb.map((c) => Math.round(c).toString(16).padStart(2, '0').toUpperCase()).join('')
}

export function applyThemeColors(themeXml: string, colors: ThemeColors): string {
  let scheme = sectionOf(themeXml, 'a:clrScheme')
  if (!scheme) return themeXml
  const original = scheme
  for (const tag of COLOR_TAGS) {
    const value = colors[tag]
    if (!value) continue
    scheme = scheme.replace(
      new RegExp(`(<a:${tag}>\\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")`),
      `$1${value}$2`,
    )
    // sysClr entries (windowText etc.) are left alone; dk2/lt2 are usually srgbClr
  }
  let xml = themeXml.replace(original, scheme)
  if (colors.name) {
    xml = xml.replace(/(<a:clrScheme name=")[^"]*(")/, `$1${escapeXmlAttr(colors.name)}$2`)
  }
  return xml
}

/**
 * Minimal but complete theme part for documents that have none (e.g. the
 * blank template). Word requires fmtScheme; this ships the standard Office
 * format scheme skeleton.
 */
export function buildThemeXml(fonts: ThemeFonts, colors: ThemeColors): string {
  const c = (tag: (typeof COLOR_TAGS)[number], fallback: string) =>
    `<a:${tag}><a:srgbClr val="${colors[tag] ?? fallback}"/></a:${tag}>`
  const font = (tag: string, typeface: string) =>
    `<${tag}><a:latin typeface="${escapeXmlAttr(typeface)}"/>` +
    `<a:ea typeface="${escapeXmlAttr(fonts.eastAsia ?? '')}"/><a:cs typeface=""/></${tag}>`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
    '<a:themeElements>' +
    `<a:clrScheme name="${escapeXmlAttr(colors.name ?? 'Office')}">` +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    c('dk2', '44546A') +
    c('lt2', 'E7E6E6') +
    c('accent1', '4472C4') +
    c('accent2', 'ED7D31') +
    c('accent3', 'A5A5A5') +
    c('accent4', 'FFC000') +
    c('accent5', '5B9BD5') +
    c('accent6', '70AD47') +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme>' +
    `<a:fontScheme name="Office">${font('a:majorFont', fonts.major)}${font('a:minorFont', fonts.minor)}</a:fontScheme>` +
    '<a:fmtScheme name="Office">' +
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
    '<a:lnStyleLst>' +
    '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '</a:lnStyleLst>' +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements></a:theme>'
  )
}

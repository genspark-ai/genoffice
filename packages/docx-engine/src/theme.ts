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
  const majorScripts = scriptFontsOf(themeXml, 'a:majorFont')
  const minorScripts = scriptFontsOf(themeXml, 'a:minorFont')
  return {
    major: major ?? '',
    minor: minor ?? '',
    ...(eastAsia ? { eastAsia } : {}),
    ...(majorEastAsia ? { majorEastAsia } : {}),
    ...(minorCs ? { minorCs } : {}),
    ...(majorCs ? { majorCs } : {}),
    ...(majorScripts ? { majorScripts } : {}),
    ...(minorScripts ? { minorScripts } : {}),
  }
}

/** per-script faces of a font group (<a:font script="Hang" typeface="…"/>) */
function scriptFontsOf(themeXml: string, tag: string): Record<string, string> | undefined {
  const section = sectionOf(themeXml, tag)
  if (!section) return undefined
  const out: Record<string, string> = {}
  for (const m of section.matchAll(/<a:font script="([^"]+)" typeface="([^"]+)"/g)) {
    out[m[1]] = m[2]
  }
  return Object.keys(out).length > 0 ? out : undefined
}

function sectionOf(xml: string, tag: string): string | null {
  return new RegExp(`<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`).exec(xml)?.[0] ?? null
}

function fontOf(xml: string, tag: string): string | null {
  const section = sectionOf(xml, tag)
  return section ? (/<a:latin typeface="([^"]*)"/.exec(section)?.[1] ?? null) : null
}

interface FontGroup {
  latin: string
  ea?: string
  cs?: string
  scripts?: Record<string, string>
}

/** rewrite the typefaces of one font group; a face left undefined keeps what the part has */
function applyFontGroup(xml: string, tag: string, group: FontGroup): string {
  const section = sectionOf(xml, tag)
  if (!section) return xml
  const face = (kind: string, typeface: string) => (all: string, was: string) =>
    was === typeface ? all : `<a:${kind} typeface="${escapeXmlAttr(typeface)}"/>`
  // a panose number describes the face it was written for
  let next = section.replace(/<a:latin typeface="([^"]*)"[^>]*>/, face('latin', group.latin))
  if (group.ea !== undefined)
    next = next.replace(/<a:ea typeface="([^"]*)"[^>]*>/, face('ea', group.ea))
  if (group.cs !== undefined)
    next = next.replace(/<a:cs typeface="([^"]*)"[^>]*>/, face('cs', group.cs))
  if (group.scripts) {
    const fonts = Object.entries(group.scripts)
      .map(
        ([script, typeface]) =>
          `<a:font script="${script}" typeface="${escapeXmlAttr(typeface)}"/>`,
      )
      .join('')
    next = next
      .replace(/\s*<a:font script="[^"]*"[^>]*>/g, '')
      .replace(`</${tag}>`, () => `${fonts}</${tag}>`)
  }
  return xml.replace(section, () => next)
}

export function applyThemeFonts(themeXml: string, fonts: ThemeFonts): string {
  let xml = applyFontGroup(themeXml, 'a:majorFont', {
    latin: fonts.major,
    ea: fonts.majorEastAsia ?? fonts.eastAsia,
    cs: fonts.majorCs,
    scripts: fonts.majorScripts,
  })
  xml = applyFontGroup(xml, 'a:minorFont', {
    latin: fonts.minor,
    ea: fonts.eastAsia,
    cs: fonts.minorCs,
    scripts: fonts.minorScripts,
  })
  if (fonts.name)
    xml = xml.replace(
      /(<a:fontScheme name=")[^"]*(")/,
      (_m, open: string, close: string) => open + escapeXmlAttr(fonts.name!) + close,
    )
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

/** built-in Office palette: what Word resolves theme colors against when a
 *  document ships no word/theme/theme1.xml part */
export const DEFAULT_THEME_COLORS: Readonly<ThemeColors> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
}

/**
 * Resolve a w:themeColor reference (+ optional w:themeTint / w:themeShade,
 * hex 00-FF) against the palette. Word scales HSL lightness (shade: L*s;
 * tint: L*t + (1-t)), which reproduces its cached w:val/w:fill within 1/255;
 * the cached value stays authoritative on save.
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
  const baseTrimmed = (base ?? '').trim()
  if (!baseTrimmed || !/^[0-9A-Fa-f]{6}$/.test(baseTrimmed)) return null
  const factorOf = (hex?: string) => {
    const v = hex ? parseInt(hex, 16) : NaN
    return Number.isFinite(v) ? Math.max(0, Math.min(255, v)) / 255 : null
  }
  const s = factorOf(shade)
  const t = factorOf(tint)
  if (s === null && t === null) return baseTrimmed.toUpperCase()
  const [h, sat, l0] = rgbToHsl(
    [0, 2, 4].map((i) => parseInt(baseTrimmed.slice(i, i + 2), 16) / 255),
  )
  let l = l0
  if (s !== null) l *= s
  if (t !== null) l = l * t + (1 - t)
  return hslToRgb(h, sat, l)
    .map((c) =>
      Math.round(c * 255)
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

function rgbToHsl([r, g, b]: number[]): [number, number, number] {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (d === 0) return [0, 0, l]
  const s = d / (1 - Math.abs(2 * l - 1))
  let h: number
  if (max === r) h = ((g - b) / d + 6) % 6
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToRgb(h: number, s: number, l: number): number[] {
  if (s === 0) return [l, l, l]
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const channel = (t0: number) => {
    const t = ((t0 % 1) + 1) % 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  return [channel(h + 1 / 3), channel(h), channel(h - 1 / 3)]
}

const SYSTEM_SLOTS: Partial<Record<keyof ThemeColors, [name: string, color: string]>> = {
  dk1: ['windowText', '000000'],
  lt1: ['window', 'FFFFFF'],
}

/** a clrScheme slot's colour element; black text and a white page stay the system's pair */
function slotColorXml(tag: (typeof READ_TAGS)[number], hex: string): string {
  const system = SYSTEM_SLOTS[tag]
  return system && system[1] === hex.toUpperCase()
    ? `<a:sysClr val="${system[0]}" lastClr="${system[1]}"/>`
    : `<a:srgbClr val="${hex}"/>`
}

export function applyThemeColors(themeXml: string, colors: ThemeColors): string {
  let scheme = sectionOf(themeXml, 'a:clrScheme')
  if (!scheme) return themeXml
  const original = scheme
  for (const tag of READ_TAGS) {
    const value = colors[tag]
    if (!value || !/^[0-9A-Fa-f]{6}$/.test(value)) continue
    scheme = scheme.replace(
      new RegExp(
        `(<a:${tag}>\\s*)<a:(?:srgbClr val|sysClr[^>]*? lastClr)="([0-9A-Fa-f]{6})"[^>]*?(?:/>|>\\s*</a:s(?:rgb|ys)Clr>)`,
      ),
      (all, open: string, was: string) =>
        was.toUpperCase() === value.toUpperCase() ? all : open + slotColorXml(tag, value),
    )
  }
  let xml = themeXml.replace(original, () => scheme)
  if (colors.name) {
    xml = xml.replace(
      /(<a:clrScheme name=")[^"]*(")/,
      (_m, open: string, close: string) => open + escapeXmlAttr(colors.name!) + close,
    )
  }
  return xml
}

/** the theme's name and its format scheme (fill, line and effect styles) as an a:fmtScheme element */
export interface ThemeFormat {
  name: string
  xml: string
}

export function readThemeFormat(themeXml: string): ThemeFormat | null {
  const xml = sectionOf(themeXml, 'a:fmtScheme')
  const name = /<a:theme\b[^>]*? name="([^"]*)"/.exec(themeXml)?.[1]
  return xml ? { name: name ?? '', xml } : null
}

export function applyThemeFormat(themeXml: string, format: ThemeFormat): string {
  const scheme = sectionOf(themeXml, 'a:fmtScheme')
  const xml = scheme ? themeXml.replace(scheme, () => format.xml) : themeXml
  return xml.replace(
    /(<a:theme\b[^>]*? name=")[^"]*(")/,
    (_m, open: string, close: string) => open + escapeXmlAttr(format.name) + close,
  )
}

/** Word refuses a theme part without a:fmtScheme; this is the Office one it writes itself */
const OFFICE_FMT_SCHEME =
  '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/>' +
  '</a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">' +
  '<a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr>' +
  '</a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/>' +
  '<a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000">' +
  '<a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr>' +
  '</a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">' +
  '<a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr>' +
  '</a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/>' +
  '<a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst>' +
  '<a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst>' +
  '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr">' +
  '<a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill>' +
  '<a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' +
  '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>' +
  '</a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst>' +
  '<a:effectStyle><a:effectLst>' +
  '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">' +
  '<a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst>' +
  '</a:effectStyle><a:effectStyle><a:effectLst>' +
  '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' +
  '<a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>' +
  '</a:effectStyle><a:effectStyle><a:effectLst>' +
  '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' +
  '<a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>' +
  '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/>' +
  '</a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/>' +
  '</a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>' +
  '</a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill>' +
  '<a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst>' +
  '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/>' +
  '</a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/>' +
  '<a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000">' +
  '<a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr>' +
  '</a:gs></a:gsLst><a:path path="circle">' +
  '<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">' +
  '<a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000">' +
  '<a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr>' +
  '</a:gs></a:gsLst><a:path path="circle">' +
  '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>' +
  '</a:bgFillStyleLst></a:fmtScheme>'

const DEFAULT_FONTS: ThemeFonts = { major: 'Calibri Light', minor: 'Calibri', eastAsia: '' }

function clrSchemeXml(colors: ThemeColors): string {
  const c = (tag: (typeof READ_TAGS)[number]) =>
    `<a:${tag}>${slotColorXml(tag, colors[tag] ?? DEFAULT_THEME_COLORS[tag]!)}</a:${tag}>`
  return (
    `<a:clrScheme name="${escapeXmlAttr(colors.name ?? 'Office')}">` +
    READ_TAGS.map(c).join('') +
    '</a:clrScheme>'
  )
}

function fontSchemeXml(fonts: ThemeFonts): string {
  const font = (tag: string, typeface: string) =>
    `<${tag}><a:latin typeface="${escapeXmlAttr(typeface)}"/>` +
    `<a:ea typeface="${escapeXmlAttr(fonts.eastAsia ?? '')}"/><a:cs typeface=""/></${tag}>`
  return `<a:fontScheme name="Office">${font('a:majorFont', fonts.major)}${font('a:minorFont', fonts.minor)}</a:fontScheme>`
}

/** Minimal but complete theme part for documents that have none (e.g. the blank template). */
export function buildThemeXml(
  fonts: ThemeFonts,
  colors: ThemeColors,
  format?: ThemeFormat,
): string {
  const skeleton =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
    '<a:themeElements>' +
    clrSchemeXml(colors) +
    fontSchemeXml(fonts) +
    OFFICE_FMT_SCHEME +
    '</a:themeElements></a:theme>'
  // the faces beyond the Latin pair and the format scheme go in the way they do for an existing part
  const themed = applyThemeFonts(skeleton, fonts)
  return format ? applyThemeFormat(themed, format) : themed
}

const hasElement = (xml: string, tag: string, child?: string) => {
  const section = sectionOf(xml, tag)
  return !!section && (!child || section.includes(`<${child}`))
}

/**
 * The schema requires clrScheme, fontScheme and fmtScheme, in that order; Word
 * refuses a document whose theme has no fmtScheme. Fills in what an a:-prefixed
 * theme lacks; a complete theme comes back unchanged.
 */
export function completeThemeElements(themeXml: string): string {
  if (!/<a:theme[\s>]/.test(themeXml)) return themeXml
  let xml = themeXml.replace(/<a:themeElements\s*\/>/, '<a:themeElements></a:themeElements>')
  if (!xml.includes('<a:themeElements')) {
    xml = xml.replace(/(<a:theme(?:\s[^>]*)?>)/, '$1<a:themeElements></a:themeElements>')
  }
  // a scheme without its required children goes whole, whichever way it is written
  const dropEmpty = (tag: string) => {
    xml = xml.replace(sectionOf(xml, tag) ?? new RegExp(`<${tag}(?:\\s[^>]*)?/>`), '')
  }
  if (!hasElement(xml, 'a:clrScheme', 'a:dk1')) {
    dropEmpty('a:clrScheme')
    xml = xml.replace(/(<a:themeElements(?:\s[^>]*)?>)/, `$1${clrSchemeXml({})}`)
  }
  if (!hasElement(xml, 'a:fontScheme', 'a:majorFont')) {
    dropEmpty('a:fontScheme')
    xml = xml.replace('</a:clrScheme>', `</a:clrScheme>${fontSchemeXml(DEFAULT_FONTS)}`)
  }
  if (!hasElement(xml, 'a:fmtScheme', 'a:fillStyleLst')) {
    dropEmpty('a:fmtScheme')
    xml = xml.replace('</a:fontScheme>', `</a:fontScheme>${OFFICE_FMT_SCHEME}`)
  }
  return xml
}

/** a Design > Themes / Colors / Fonts pick the save has not written yet */
export interface ThemeEdit {
  fonts?: ThemeFonts | null
  colors?: ThemeColors | null
  format?: ThemeFormat | null
}

/** the theme part a save of `edit` leaves: the document's own with the pick written in, or a new one */
export function editedThemeXml(themeXml: string | undefined, edit: ThemeEdit): string {
  const { fonts, colors, format } = edit
  if (themeXml === undefined) {
    return buildThemeXml(
      fonts ?? { major: 'Calibri Light', minor: 'Calibri', eastAsia: '' },
      colors ?? {},
      format ?? undefined,
    )
  }
  let xml = completeThemeElements(themeXml)
  if (fonts) xml = applyThemeFonts(xml, fonts)
  if (colors) xml = applyThemeColors(xml, colors)
  if (format) xml = applyThemeFormat(xml, format)
  return xml
}

/** a theme edit applied to an existing part, which leaves it a part Word opens */
export function rewriteThemeXml(
  themeXml: string,
  fonts: ThemeFonts | undefined,
  colors: ThemeColors | undefined,
): string {
  return editedThemeXml(themeXml, { fonts, colors })
}

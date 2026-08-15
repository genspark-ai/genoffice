/**
 * Theme reading (read-only, never written back).
 * Provides color scheme and font scheme lookups so inheritance resolution can turn
 * schemeClr / theme fonts into final values.
 */
import { XMLParser } from 'fast-xml-parser'
import { asXmlNode, type XmlNode } from './xml-utils'

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

export interface Theme {
  /** clrScheme: dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink → #RRGGBB */
  colors: Record<string, string>
  /**
   * Master p:clrMap (or slide/layout a:overrideClrMapping): schemeClr name → clrScheme key.
   * Dark masters invert it (bg1="dk1" tx1="lt1" …); absent → standard mapping.
   */
  clrMap?: Record<string, string>
  /** fontScheme: major / minor Latin fonts */
  majorFont?: string
  minorFont?: string
  /** fontScheme: major / minor East Asian fonts */
  majorEaFont?: string
  minorEaFont?: string
  /** fontScheme: major / minor Complex Script fonts (Arabic/Hebrew/Thai/Hindi etc.) */
  majorCsFont?: string
  minorCsFont?: string
  /**
   * fmtScheme templates (referenced by style refs: fillRef/lnRef/effectRef; phClr is
   * substituted by the referencing side). Each entry is a parsed node (e.g.
   * { 'a:gradFill': {...} } / { 'a:ln': {...} } / { 'a:effectStyle': {...} }), mapping
   * to idx 1..3 in original XML order; bgFills map to idx 1001..1003.
   */
  fillStyles?: XmlNode[]
  lnStyles?: XmlNode[]
  effectStyles?: XmlNode[]
  bgFillStyles?: XmlNode[]
}

/** Scan the top-level children (depth 1) of an xml fragment and return each child's full fragment (order preserved). */
function topLevelFragments(xml: string): string[] {
  const out: string[] = []
  const re = /<(\/?)([a-zA-Z][\w:]*)((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/g
  let depth = 0
  let curStart = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const closing = m[1] === '/'
    const selfClose = m[4] === '/'
    if (!closing && !selfClose) {
      if (depth === 0) curStart = m.index
      depth++
    } else if (closing) {
      depth--
      if (depth === 0 && curStart >= 0) out.push(xml.slice(curStart, re.lastIndex))
    } else if (selfClose && depth === 0) {
      out.push(xml.slice(m.index, re.lastIndex))
    }
  }
  return out
}

/** Get the list of inner fragments of <tag>...</tag> and parse each into a regular node (preserving idx order). */
function parseStyleList(themeXml: string, tag: string): XmlNode[] | undefined {
  const m = new RegExp(`<a:${tag}\\b[^>]*>([\\s\\S]*?)</a:${tag}>`).exec(themeXml)
  if (!m) return undefined
  const items = topLevelFragments(m[1]!).map((frag) => asXmlNode(parser.parse(frag)))
  return items.length ? items : undefined
}

function readColorNode(node: unknown): string | undefined {
  if (!node) return undefined
  const n = asXmlNode(node)
  const srgb = asXmlNode(n['a:srgbClr'])
  if (n['a:srgbClr']) return '#' + String(srgb['@_val']).toUpperCase()
  const sys = asXmlNode(n['a:sysClr'])
  if (n['a:sysClr']) return '#' + String(sys['@_lastClr'] ?? '000000').toUpperCase()
  return undefined
}

/** Font typeface attribute of e.g. fontScheme['a:majorFont']['a:latin'] (undefined when absent). */
function typeface(scheme: XmlNode, font: string, script: string): string | undefined {
  const v = asXmlNode(asXmlNode(scheme[font])[script])['@_typeface']
  return typeof v === 'string' && v ? v : undefined
}

export function parseTheme(themeXml: string): Theme {
  const doc = asXmlNode(parser.parse(themeXml))
  const themeEl = asXmlNode(doc['a:theme'] ?? doc.theme)
  const elements = asXmlNode(themeEl['a:themeElements'])
  const clrScheme = asXmlNode(elements['a:clrScheme'])
  const colors: Record<string, string> = {}
  for (const key of [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]) {
    const c = readColorNode(clrScheme['a:' + key])
    if (c) colors[key] = c
  }
  const fontScheme = asXmlNode(elements['a:fontScheme'])
  const majorFont = typeface(fontScheme, 'a:majorFont', 'a:latin')
  const minorFont = typeface(fontScheme, 'a:minorFont', 'a:latin')
  const majorEaFont = typeface(fontScheme, 'a:majorFont', 'a:ea')
  const minorEaFont = typeface(fontScheme, 'a:minorFont', 'a:ea')
  const majorCsFont = typeface(fontScheme, 'a:majorFont', 'a:cs')
  const minorCsFont = typeface(fontScheme, 'a:minorFont', 'a:cs')
  // fmtScheme templates are parsed in order from raw-text slices (fillStyleLst children mix element types; regular parsing would lose the idx order)
  const fmtM = /<a:fmtScheme\b[^>]*>[\s\S]*?<\/a:fmtScheme>/.exec(themeXml)
  const fmt = fmtM?.[0] ?? ''
  const fillStyles = parseStyleList(fmt, 'fillStyleLst')
  const lnStyles = parseStyleList(fmt, 'lnStyleLst')
  const effectStyles = parseStyleList(fmt, 'effectStyleLst')
  const bgFillStyles = parseStyleList(fmt, 'bgFillStyleLst')
  return {
    colors,
    majorFont,
    minorFont,
    majorEaFont,
    minorEaFont,
    majorCsFont,
    minorCsFont,
    ...(fillStyles ? { fillStyles } : {}),
    ...(lnStyles ? { lnStyles } : {}),
    ...(effectStyles ? { effectStyles } : {}),
    ...(bgFillStyles ? { bgFillStyles } : {}),
  }
}

/**
 * Theme font reference ("+mj-lt" / "+mn-ea" etc.) → final font name.
 * Values not starting with "+" are returned as-is; returns undefined when the theme has no match.
 */
export function resolveFontRef(
  typeface: string | undefined,
  theme: Theme | undefined,
): string | undefined {
  if (!typeface) return undefined
  if (!typeface.startsWith('+')) return typeface
  switch (typeface) {
    case '+mj-lt':
      return theme?.majorFont
    case '+mn-lt':
      return theme?.minorFont
    case '+mj-ea':
      return theme?.majorEaFont ?? theme?.majorFont
    case '+mn-ea':
      return theme?.minorEaFont ?? theme?.minorFont
    case '+mj-cs':
      return theme?.majorCsFont ?? theme?.majorFont
    case '+mn-cs':
      return theme?.minorCsFont ?? theme?.minorFont
    default:
      return theme?.minorFont
  }
}

/** schemeClr name (e.g. 'tx1','bg1','accent1','phClr') → final #RRGGBB. */
export function resolveSchemeColor(
  name: string,
  theme: Theme | undefined,
  phClr?: string,
): string | undefined {
  if (name === 'phClr') return phClr
  // Standard mapping: tx1→dk1, bg1→lt1, tx2→dk2, bg2→lt2
  const standard: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }
  const key = theme?.clrMap?.[name] ?? standard[name] ?? name
  return theme?.colors[key]
}

const CLR_MAP_NAMES = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]

function clrMapFromTag(tag: string): Record<string, string> | undefined {
  const out: Record<string, string> = {}
  for (const name of CLR_MAP_NAMES) {
    const m = new RegExp(`\\b${name}="([^"]+)"`).exec(tag)
    if (m) out[name] = m[1]
  }
  return Object.keys(out).length ? out : undefined
}

/**
 * Effective color mapping for a slide: slide/layout <a:overrideClrMapping> wins over the
 * master's <p:clrMap> (<a:masterClrMapping/> just inherits). Undefined → standard mapping.
 */
export function parseClrMap(
  masterXml?: string,
  layoutXml?: string,
  slideXml?: string,
): Record<string, string> | undefined {
  for (const xml of [slideXml, layoutXml]) {
    const tag = xml ? /<a:overrideClrMapping\b[^>]*\/?>/.exec(xml)?.[0] : undefined
    const map = tag ? clrMapFromTag(tag) : undefined
    if (map) return map
  }
  const tag = masterXml ? /<p:clrMap\b[^>]*\/?>/.exec(masterXml)?.[0] : undefined
  return tag ? clrMapFromTag(tag) : undefined
}

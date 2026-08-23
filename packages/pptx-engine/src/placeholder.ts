/**
 * Placeholder geometry inheritance (Phase 2, filling a gap left by Phase 1).
 *
 * Placeholder shapes on a slide (<p:ph>) often omit their own <a:xfrm>; geometry
 * must be backfilled from the slideLayout per OOXML inheritance rules, and from
 * the slideMaster when the layout lacks it.
 *
 * Matching rules (aligned with PowerPoint / pptxtojson):
 *   1. exact (type, idx) match first;
 *   2. then by idx (when type is missing/inconsistent);
 *   3. then by type;
 *   4. body-like types (body/subTitle, etc.) fall back to each other, as do
 *      title/ctrTitle.
 *
 * Read-only: layout/master geometry is used only for render inheritance, never
 * written back to the original pptx.
 */
import { XMLParser } from 'fast-xml-parser'
import type { Transform, TextAlign } from './types'
import { type Theme, resolveFontRef } from './theme'
import { resolveColorNode } from './color'
import { asXmlNode, xmlArray, type XmlNode } from './xml-utils'

const phParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => ['p:sp'].includes(name),
})

/** Default run/paragraph style for one indent level (from lstStyle's lvlNpPr/defRPr). */
export interface LevelTextStyle {
  /** Font size (pt) */
  fontSize?: number
  bold?: boolean
  italic?: boolean
  /** Character casing <a:defRPr cap>: 'all' | 'small' | 'none' */
  cap?: string
  color?: string
  /** <a:defRPr><a:effectLst><a:outerShdw> (rare; e.g. watermark text boxes) */
  shadow?: import('./types').ShadowEffect
  latinFont?: string
  eaFont?: string
  csFont?: string
  align?: TextAlign
  /** Bullet default (master bodyStyle levels commonly use buChar '•') */
  bullet?: {
    type: 'none' | 'char' | 'number'
    char?: string
    /** <a:buFont> typeface (symbol fonts like Wingdings) */
    font?: string
    color?: string
    /** <a:buSzPct> (%, 100 = same size as text) */
    sizePct?: number
    /** <a:buAutoNum type> (arabicPeriod/romanLcParen…) */
    numType?: string
    /** <a:buAutoNum startAt> */
    startAt?: number
  }
  /** Paragraph left indent (EMU) */
  marL?: number
  /** First-line indent (EMU, negative = hanging) */
  indent?: number
  /** Line/paragraph spacing defaults (lnSpc/spcBef/spcAft from master txStyles; source of inherited line spacing) */
  lineHeight?: number
  lineExact?: number
  spaceBefore?: number
  spaceAfter?: number
  spaceBeforePct?: number
  spaceAfterPct?: number
}

/** Default styles for the 9 levels (index = level, 0-based). */
export interface TextStyleLevels {
  levels: Array<LevelTextStyle | undefined>
}

/** master <p:txStyles>: the title/body/other families. */
export interface MasterTextStyles {
  title?: TextStyleLevels
  body?: TextStyleLevels
  other?: TextStyleLevels
}

export interface PlaceholderGeom {
  type: string // ph type, normalized to 'body' when missing
  idx: string // ph idx, '' when missing
  transform: Transform | null
  /** Default text style from this placeholder txBody's <a:lstStyle> */
  textStyle?: TextStyleLevels
  /** Vertical anchor from this placeholder's <a:bodyPr anchor=""> */
  anchor?: 'top' | 'middle' | 'bottom'
  /** Raw spPr node when it carries an explicit fill (parse.ts resolves it with the part's rels) */
  fillSpPr?: unknown
}

/** Placeholder geometry table for one layer (layout or master). */
export interface PlaceholderMap {
  entries: PlaceholderGeom[]
}

const TITLE_TYPES = new Set(['title', 'ctrTitle'])
const BODY_TYPES = new Set(['body', 'subTitle', 'obj', ''])

const ANCHOR_MAP: Record<string, PlaceholderGeom['anchor']> = {
  t: 'top',
  ctr: 'middle',
  b: 'bottom',
}

const FILL_TAGS = ['a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:noFill']

function parseXfrmNode(xfrmRaw: unknown): Transform | null {
  if (!xfrmRaw) return null
  const xfrm = asXmlNode(xfrmRaw)
  const offRaw = xfrm['a:off']
  const extRaw = xfrm['a:ext']
  if (!offRaw && !extRaw) return null
  const off = asXmlNode(offRaw)
  const ext = asXmlNode(extRaw)
  return {
    offset: {
      x: offRaw ? parseInt(String(off['@_x']), 10) || 0 : 0,
      y: offRaw ? parseInt(String(off['@_y']), 10) || 0 : 0,
      cx: extRaw ? parseInt(String(ext['@_cx']), 10) || 0 : 0,
      cy: extRaw ? parseInt(String(ext['@_cy']), 10) || 0 : 0,
    },
    rot: xfrm['@_rot'] ? parseInt(String(xfrm['@_rot']), 10) || 0 : 0,
    flipH: xfrm['@_flipH'] === '1' || xfrm['@_flipH'] === 'true',
    flipV: xfrm['@_flipV'] === '1' || xfrm['@_flipV'] === 'true',
  }
}

/**
 * Extract the placeholder geometry table + lstStyle text style defaults from a
 * layout/master's full XML. Collects shapes that carry <p:ph> and at least one of
 * <a:xfrm> or <a:lstStyle>.
 */
export function parsePlaceholderMap(layoutOrMasterXml: string, theme?: Theme): PlaceholderMap {
  const entries: PlaceholderGeom[] = []
  let doc: XmlNode
  try {
    doc = asXmlNode(phParser.parse(layoutOrMasterXml))
  } catch {
    return { entries }
  }
  // Path: p:sldLayout / p:sldMaster → p:cSld → p:spTree → p:sp[]
  const root = asXmlNode(doc['p:sldLayout'] ?? doc['p:sldMaster'])
  const spTreeRaw = asXmlNode(root['p:cSld'])['p:spTree']
  if (!spTreeRaw) return { entries }
  const spTree = asXmlNode(spTreeRaw)
  for (const sp of xmlArray(spTree['p:sp'])) {
    const phRaw = asXmlNode(asXmlNode(sp['p:nvSpPr'])['p:nvPr'])['p:ph']
    if (!phRaw) continue
    const ph = asXmlNode(phRaw)
    const type = String(ph['@_type'] ?? 'body')
    const idx = ph['@_idx'] != null ? String(ph['@_idx']) : ''
    const spPr = asXmlNode(sp['p:spPr'])
    const transform = parseXfrmNode(spPr['a:xfrm'])
    const textStyle = parseLstStyleLevels(asXmlNode(sp['p:txBody'])['a:lstStyle'], theme)
    const anchorAttr = asXmlNode(asXmlNode(sp['p:txBody'])['a:bodyPr'])['@_anchor']
    const anchor = ANCHOR_MAP[String(anchorAttr ?? '')]
    const hasFill = FILL_TAGS.some((tag) => tag in spPr)
    if (!transform && !textStyle && !anchor && !hasFill) continue
    entries.push({
      type,
      idx,
      transform,
      ...(textStyle ? { textStyle } : {}),
      ...(anchor ? { anchor } : {}),
      ...(hasFill ? { fillSpPr: spPr } : {}),
    })
  }
  return { entries }
}

// ── lstStyle / txStyles text style parsing ───────────────────────────

const ALIGN_MAP: Record<string, TextAlign> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
}

/** fast-xml-parser does not decode numeric character references in attributes (&#x2022; etc.); done here. */
function decodeAttrCharRefs(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
}

/** <a:spcPct val="150000"/> → 150 (%). */
function spcPctVal(node: unknown): number | undefined {
  const v = asXmlNode(asXmlNode(node)['a:spcPct'])['@_val']
  return v != null ? parseInt(String(v), 10) / 1000 : undefined
}

/** <a:spcPts val="2400"/> → 24 (pt). */
function spcPtsVal(node: unknown): number | undefined {
  const v = asXmlNode(asXmlNode(node)['a:spcPts'])['@_val']
  return v != null ? parseInt(String(v), 10) / 100 : undefined
}

/** Typeface attribute of e.g. <a:latin typeface="…"/> as a string (undefined when absent). */
function typefaceAttr(node: unknown): string | undefined {
  const v = asXmlNode(node)['@_typeface']
  return v != null ? String(v) : undefined
}

/** <a:lvlNpPr> (including defRPr) → LevelTextStyle. */
function parseLvlPPr(pPrRaw: unknown, theme?: Theme): LevelTextStyle | undefined {
  if (!pPrRaw || typeof pPrRaw !== 'object') return undefined
  const pPr = asXmlNode(pPrRaw)
  const out: LevelTextStyle = {}
  const algn = String(pPr['@_algn'] ?? '')
  if (algn && ALIGN_MAP[algn]) out.align = ALIGN_MAP[algn]
  const lineHeight = spcPctVal(pPr['a:lnSpc'])
  const lineExact = spcPtsVal(pPr['a:lnSpc'])
  if (lineHeight != null) out.lineHeight = lineHeight
  if (lineExact != null) out.lineExact = lineExact
  const spcBef = pPr['a:spcBef']
  const spcAft = pPr['a:spcAft']
  if (spcPtsVal(spcBef) != null) out.spaceBefore = spcPtsVal(spcBef)
  if (spcPctVal(spcBef) != null) out.spaceBeforePct = spcPctVal(spcBef)
  if (spcPtsVal(spcAft) != null) out.spaceAfter = spcPtsVal(spcAft)
  if (spcPctVal(spcAft) != null) out.spaceAfterPct = spcPctVal(spcAft)
  // Bullet + indent defaults (a body placeholder's bullets mostly come from here)
  const buChar = asXmlNode(pPr['a:buChar'])['@_char']
  if (pPr['a:buNone'] !== undefined) out.bullet = { type: 'none' }
  else if (buChar != null) {
    out.bullet = { type: 'char', char: decodeAttrCharRefs(String(buChar)) }
  } else if (pPr['a:buAutoNum']) {
    out.bullet = { type: 'number' }
    const an = asXmlNode(pPr['a:buAutoNum'])
    if (an['@_type'] != null) out.bullet.numType = String(an['@_type'])
    const startAt = parseInt(String(an['@_startAt']), 10)
    if (Number.isFinite(startAt) && startAt > 1) out.bullet.startAt = startAt
  }
  if (out.bullet && out.bullet.type !== 'none') {
    const buFont = typefaceAttr(pPr['a:buFont'])
    if (buFont) out.bullet.font = buFont
    const buColor = resolveColorNode(pPr['a:buClr'], theme)
    if (buColor) out.bullet.color = buColor
    const buSz = asXmlNode(pPr['a:buSzPct'])['@_val']
    if (buSz != null) out.bullet.sizePct = (parseInt(String(buSz), 10) || 0) / 1000
  }
  if (pPr['@_marL'] != null) {
    const v = parseInt(String(pPr['@_marL']), 10)
    if (!Number.isNaN(v)) out.marL = v
  }
  if (pPr['@_indent'] != null) {
    const v = parseInt(String(pPr['@_indent']), 10)
    if (!Number.isNaN(v)) out.indent = v
  }
  const defRPrRaw = pPr['a:defRPr']
  if (defRPrRaw && typeof defRPrRaw === 'object') {
    const defRPr = asXmlNode(defRPrRaw)
    if (defRPr['@_sz']) out.fontSize = parseInt(String(defRPr['@_sz']), 10) / 100
    if (defRPr['@_b'] != null) out.bold = defRPr['@_b'] === '1' || defRPr['@_b'] === 'true'
    if (defRPr['@_i'] != null) out.italic = defRPr['@_i'] === '1' || defRPr['@_i'] === 'true'
    if (defRPr['@_cap'] != null) out.cap = String(defRPr['@_cap'])
    const color = resolveColorNode(defRPr['a:solidFill'], theme)
    if (color) out.color = color
    const shdw = asXmlNode(asXmlNode(defRPr['a:effectLst'])['a:outerShdw'])
    const shdwColor = resolveColorNode(shdw, theme)
    if (shdwColor) {
      const num = (k: string) => {
        const v = parseInt(String(shdw[k] ?? ''), 10)
        return Number.isFinite(v) ? v : 0
      }
      out.shadow = {
        color: shdwColor,
        blurRad: num('@_blurRad'),
        dist: num('@_dist'),
        dirDeg: num('@_dir') / 60000,
      }
    }
    const latin = resolveFontRef(typefaceAttr(defRPr['a:latin']), theme)
    if (latin) out.latinFont = latin
    const ea = resolveFontRef(typefaceAttr(defRPr['a:ea']), theme)
    if (ea) out.eaFont = ea
    const cs = resolveFontRef(typefaceAttr(defRPr['a:cs']), theme)
    if (cs) out.csFont = cs
  }
  return Object.keys(out).length ? out : undefined
}

/** <a:lstStyle> (or one txStyles family) → 9-level style table. */
export function parseLstStyleLevels(lst: unknown, theme?: Theme): TextStyleLevels | undefined {
  if (!lst || typeof lst !== 'object') return undefined
  const l = asXmlNode(lst)
  const levels: Array<LevelTextStyle | undefined> = []
  for (let i = 1; i <= 9; i++) levels[i - 1] = parseLvlPPr(l[`a:lvl${i}pPr`], theme)
  return levels.some(Boolean) ? { levels } : undefined
}

/** presentation.xml <p:defaultTextStyle> → base defaults for non-placeholder text boxes. */
export function parseDefaultTextStyle(
  presentationXml: string,
  theme?: Theme,
): TextStyleLevels | undefined {
  let doc: XmlNode
  try {
    doc = asXmlNode(phParser.parse(presentationXml))
  } catch {
    return undefined
  }
  return parseLstStyleLevels(asXmlNode(doc['p:presentation'])['p:defaultTextStyle'], theme)
}

/** master <p:txStyles> → default styles for the title/body/other families. */
export function parseMasterTextStyles(masterXml: string, theme?: Theme): MasterTextStyles {
  let doc: XmlNode
  try {
    doc = asXmlNode(phParser.parse(masterXml))
  } catch {
    return {}
  }
  const txRaw = asXmlNode(doc['p:sldMaster'])['p:txStyles']
  if (!txRaw) return {}
  const tx = asXmlNode(txRaw)
  return {
    ...(parseLstStyleLevels(tx['p:titleStyle'], theme)
      ? { title: parseLstStyleLevels(tx['p:titleStyle'], theme)! }
      : {}),
    ...(parseLstStyleLevels(tx['p:bodyStyle'], theme)
      ? { body: parseLstStyleLevels(tx['p:bodyStyle'], theme)! }
      : {}),
    ...(parseLstStyleLevels(tx['p:otherStyle'], theme)
      ? { other: parseLstStyleLevels(tx['p:otherStyle'], theme)! }
      : {}),
  }
}

/** Find the lstStyle in a layer's map by placeholder (type, idx). Same match order as geometry inheritance. */
function findStyleInMap(
  map: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): TextStyleLevels | undefined {
  if (!map || map.entries.length === 0) return undefined
  const t = type ?? 'body'
  const i = idx ?? ''
  const styled = map.entries.filter((e) => e.textStyle)
  let hit = styled.find((e) => e.type === t && e.idx === i)
  if (!hit && i !== '') hit = styled.find((e) => e.idx === i)
  if (!hit) hit = styled.find((e) => e.type === t)
  if (!hit && TITLE_TYPES.has(t)) hit = styled.find((e) => TITLE_TYPES.has(e.type))
  if (!hit && BODY_TYPES.has(t)) hit = styled.find((e) => BODY_TYPES.has(e.type))
  return hit?.textStyle
}

/** Find the bodyPr anchor in a layer's map by placeholder (type, idx). Unlike
 *  findStyleInMap there is no idx-only step: master placeholders reuse idx values
 *  across types (a dt idx="2" must not anchor a body idx="2" to the bottom). */
function findAnchorInMap(
  map: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): PlaceholderGeom['anchor'] {
  if (!map || map.entries.length === 0) return undefined
  const t = type ?? 'body'
  const i = idx ?? ''
  const anchored = map.entries.filter((e) => e.anchor)
  let hit = anchored.find((e) => e.type === t && e.idx === i)
  if (!hit) hit = anchored.find((e) => e.type === t)
  if (!hit && TITLE_TYPES.has(t)) hit = anchored.find((e) => TITLE_TYPES.has(e.type))
  if (!hit && BODY_TYPES.has(t)) hit = anchored.find((e) => BODY_TYPES.has(e.type))
  return hit?.anchor
}

/** Resolve a placeholder's inherited vertical anchor: layout first, master as fallback. */
export function resolvePlaceholderAnchor(
  layout: PlaceholderMap | undefined,
  master: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): PlaceholderGeom['anchor'] {
  return findAnchorInMap(layout, type, idx) ?? findAnchorInMap(master, type, idx)
}

/**
 * Placeholder text style inheritance chain (highest priority first):
 * layout placeholder lstStyle → master placeholder lstStyle → master txStyles
 * (by ph family). The caller may prepend the slide shape's own lstStyle.
 */
export function placeholderStyleChain(
  layout: PlaceholderMap | undefined,
  master: PlaceholderMap | undefined,
  masterTx: MasterTextStyles | undefined,
  type: string | undefined,
  idx: string | undefined,
): TextStyleLevels[] {
  const chain: TextStyleLevels[] = []
  const fromLayout = findStyleInMap(layout, type, idx)
  if (fromLayout) chain.push(fromLayout)
  const fromMaster = findStyleInMap(master, type, idx)
  if (fromMaster) chain.push(fromMaster)
  const t = type ?? 'body'
  const family = TITLE_TYPES.has(t)
    ? masterTx?.title
    : BODY_TYPES.has(t)
      ? masterTx?.body
      : masterTx?.other
  if (family) chain.push(family)
  return chain
}

/**
 * Merge one level's default style field by field along the inheritance chain
 * (earlier layers win). When a layer lacks the level, fall back to that layer's
 * lvl1 (lenient fallback).
 */
export function mergeTextStyleChain(
  chain: Array<TextStyleLevels | undefined>,
  level: number,
): LevelTextStyle | undefined {
  const out: LevelTextStyle = {}
  let any = false
  for (const layer of chain) {
    if (!layer) continue
    const lvl = layer.levels[level] ?? layer.levels[0]
    if (!lvl) continue
    any = true
    if (out.fontSize == null && lvl.fontSize != null) out.fontSize = lvl.fontSize
    if (out.bold == null && lvl.bold != null) out.bold = lvl.bold
    if (out.italic == null && lvl.italic != null) out.italic = lvl.italic
    if (out.cap == null && lvl.cap != null) out.cap = lvl.cap
    if (out.color == null && lvl.color != null) out.color = lvl.color
    if (out.shadow == null && lvl.shadow != null) out.shadow = lvl.shadow
    if (out.latinFont == null && lvl.latinFont != null) out.latinFont = lvl.latinFont
    if (out.eaFont == null && lvl.eaFont != null) out.eaFont = lvl.eaFont
    if (out.csFont == null && lvl.csFont != null) out.csFont = lvl.csFont
    if (out.align == null && lvl.align != null) out.align = lvl.align
    if (out.bullet == null && lvl.bullet != null) out.bullet = lvl.bullet
    if (out.marL == null && lvl.marL != null) out.marL = lvl.marL
    if (out.indent == null && lvl.indent != null) out.indent = lvl.indent
    // Line/paragraph spacing inherit as attribute pairs (pct and pts are two value forms of the same lnSpc/spcBef/spcAft node)
    if (
      out.lineHeight == null &&
      out.lineExact == null &&
      (lvl.lineHeight != null || lvl.lineExact != null)
    ) {
      if (lvl.lineHeight != null) out.lineHeight = lvl.lineHeight
      if (lvl.lineExact != null) out.lineExact = lvl.lineExact
    }
    if (
      out.spaceBefore == null &&
      out.spaceBeforePct == null &&
      (lvl.spaceBefore != null || lvl.spaceBeforePct != null)
    ) {
      if (lvl.spaceBefore != null) out.spaceBefore = lvl.spaceBefore
      if (lvl.spaceBeforePct != null) out.spaceBeforePct = lvl.spaceBeforePct
    }
    if (
      out.spaceAfter == null &&
      out.spaceAfterPct == null &&
      (lvl.spaceAfter != null || lvl.spaceAfterPct != null)
    ) {
      if (lvl.spaceAfter != null) out.spaceAfter = lvl.spaceAfter
      if (lvl.spaceAfterPct != null) out.spaceAfterPct = lvl.spaceAfterPct
    }
  }
  return any ? out : undefined
}

function findInMap(
  map: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): Transform | undefined {
  if (!map || map.entries.length === 0) return undefined
  const t = type ?? 'body'
  const i = idx ?? ''
  // Search only entries with geometry (pure lstStyle entries do not participate in geometry inheritance)
  const geo = map.entries.filter((e) => e.transform)

  // 1. Exact (type, idx)
  let hit = geo.find((e) => e.type === t && e.idx === i)
  if (hit) return hit.transform!
  // 2. By idx (when idx is non-empty)
  if (i !== '') {
    hit = geo.find((e) => e.idx === i)
    if (hit) return hit.transform!
  }
  // 3. By type
  hit = geo.find((e) => e.type === t)
  if (hit) return hit.transform!
  // 4. Family fallback
  if (TITLE_TYPES.has(t)) {
    hit = geo.find((e) => TITLE_TYPES.has(e.type))
    if (hit) return hit.transform!
  }
  if (BODY_TYPES.has(t)) {
    hit = geo.find((e) => BODY_TYPES.has(e.type))
    if (hit) return hit.transform!
  }
  return undefined
}

/**
 * Placeholder fill inheritance: the nearest layer (layout, then master) whose matching
 * placeholder spPr carries an explicit fill — including <a:noFill>, which stops the fallback.
 * Returns the raw spPr plus its layer so the caller resolves blip rIds in that part's rels.
 */
export function resolvePlaceholderFillSpPr(
  layout: PlaceholderMap | undefined,
  master: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): { spPr: unknown; layer: 'layout' | 'master' } | undefined {
  const t = type ?? 'body'
  const i = idx ?? ''
  for (const [map, layer] of [
    [layout, 'layout'],
    [master, 'master'],
  ] as const) {
    if (!map) continue
    // Match the placeholder first (same order as geometry inheritance), THEN take its
    // fill. Filtering to fill-carrying entries up front made an unrelated same-type
    // sibling donate its fill when the true match had none (fed deck: a body ph
    // picked up a decorative content placeholder's dark fill).
    const entries = map.entries
    const hit =
      entries.find((e) => e.type === t && e.idx === i) ??
      (i !== '' ? entries.find((e) => e.idx === i) : undefined) ??
      entries.find((e) => e.type === t) ??
      (TITLE_TYPES.has(t) ? entries.find((e) => TITLE_TYPES.has(e.type)) : undefined) ??
      (BODY_TYPES.has(t) ? entries.find((e) => BODY_TYPES.has(e.type)) : undefined)
    if (hit?.fillSpPr != null) return { spPr: hit.fillSpPr, layer }
  }
  return undefined
}

/**
 * Resolve placeholder geometry: layout first, master as fallback.
 * undefined means neither layer has inheritable geometry (the render layer may
 * further fall back to a default or (0,0)).
 */
export function resolvePlaceholderTransform(
  layout: PlaceholderMap | undefined,
  master: PlaceholderMap | undefined,
  type: string | undefined,
  idx: string | undefined,
): Transform | undefined {
  return findInMap(layout, type, idx) ?? findInMap(master, type, idx)
}

// DrawingML geometry and color helpers: line presets, scheme/preset colors,
// anchor metadata and page-position resolution, group transforms.
import { attrsOf, childrenOf, findChild, nameOf, type XNode } from './xml-utils'
import { EMU_PER_PX } from './parse-xml-text'
import { placedOutsideGroup, type GroupXfrm } from './group-drawings'
import { DEFAULT_THEME_COLORS } from './theme'
import type {
  SectionSettings,
  TextGlow,
  TextOutline,
  TextFill,
  TextFillStop,
  TextReflection,
  TextShadow,
  TextboxDisplay,
  ThemeColors,
} from './types'

/**
 * Display-only extraction of anchored textboxes: DrawingML (wps:wsp, converter
 * output for code boxes / callout cards) and legacy VML (v:shape etc. inside
 * w:pict, common in broker research templates). Expects
 * fallback-stripped XML, otherwise the mc:Fallback VML twin would duplicate
 * every box.
 */
export const LINE_PRSTS_RE =
  /<a:prstGeom[^>]*prst="(?:line|straightConnector1|bentConnector[234]|curvedConnector[234])"/

/** stroke-only line/connector prsts shown as display boxes despite no text body */
export const LINE_PRSTS = new Set([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
])

/** a:xfrm flipH / flipV onto a display box */
export function applyXfrmFlips(
  box: { flipH?: boolean; flipV?: boolean },
  xfrmAttrs: Record<string, string>,
): void {
  if (xfrmAttrs['flipH'] === '1' || xfrmAttrs['flipH'] === 'true') box.flipH = true
  if (xfrmAttrs['flipV'] === '1' || xfrmAttrs['flipV'] === 'true') box.flipV = true
}

/** the box a level or vertical line takes across its stroke */
export const LINE_BAND_PX = 12

/** wps line shape → display-only line box (synthetic prst carries the arrow ends) */
export function lineBoxOf(
  shape: XNode,
  theme?: ThemeColors | null,
  ctm?: GroupCtm | null,
): TextboxDisplay | null {
  const spPr = findChild(shape, 'wps:spPr')
  if (!spPr) return null
  const prst = attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst']
  if (!prst || !LINE_PRSTS.has(prst)) return null
  const box: TextboxDisplay = { paras: [], readOnly: true }
  const ln = findChild(spPr, 'a:ln')
  const border =
    (ln
      ? attrsOf(findChild(findChild(ln, 'a:solidFill') ?? {}, 'a:srgbClr') ?? {})['val']
      : undefined) ??
    // theme-styled connectors (Word gallery): stroke from wps:style a:lnRef
    colorNodeHex(findChild(findChild(shape, 'wps:style') ?? {}, 'a:lnRef'), theme)
  box.borderColor = border ?? '000000'
  const arrowEnd = (name: string): boolean => {
    const type = attrsOf(findChild(ln ?? {}, name) ?? {})['type']
    return !!type && type !== 'none'
  }
  const head = arrowEnd('a:headEnd')
  const tail = arrowEnd('a:tailEnd')
  box.prst = prst.startsWith('bentConnector')
    ? 'lineBent'
    : prst.startsWith('curvedConnector')
      ? 'lineCurved'
      : head && tail
        ? 'lineArrowDouble'
        : head || tail
          ? 'lineArrow'
          : 'line'
  const xfrm = findChild(spPr, 'a:xfrm')
  const xfrmAttrs = attrsOf(xfrm ?? {})
  const ext = findChild(xfrm ?? {}, 'a:ext')
  const cx = ext ? parseInt(attrsOf(ext)['cx'] ?? '', 10) : NaN
  const cy = ext ? parseInt(attrsOf(ext)['cy'] ?? '', 10) : NaN
  if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round((cx * (ctm?.sx ?? 1)) / EMU_PER_PX)
  const straight = box.prst === 'line' || box.prst === 'lineArrow' || box.prst === 'lineArrowDouble'
  const vertical = straight && cx === 0 && Number.isFinite(cy) && cy > 0
  // flips apply at any height: Word's horizontal lines carry cy="0", and a
  // flipH there still reverses which tip holds the arrow
  applyXfrmFlips(box, xfrmAttrs)
  // a:xfrm rot (60000ths of a degree) and a:ln w (EMU) round-trip on lines the
  // same way they do on text-bearing shapes, so a rotated/resized connector
  // keeps its transform across save and reload
  const rot = parseInt(xfrmAttrs['rot'] ?? '', 10)
  if (Number.isFinite(rot) && rot !== 0) box.rotDeg = Math.round(rot / 60000)
  const lnW = parseInt(ln ? (attrsOf(ln)['w'] ?? '') : '', 10)
  if (Number.isFinite(lnW) && lnW > 0)
    box.borderWidthPx = Math.round((lnW / EMU_PER_PX) * 100) / 100
  if (Number.isFinite(cy) && cy > 0) {
    box.heightPx = Math.round((cy * (ctm?.sy ?? 1)) / EMU_PER_PX)
    box.minHeightPx = box.heightPx
    // a real vertical extent means the connector runs corner to corner
    // (≤12 px stays level: our own inserted lines keep a 12 px grab band)
    if (vertical) {
      box.lineVert = true
      box.widthPx = LINE_BAND_PX
    } else if (straight && (box.heightPx > 12 || box.flipH || box.flipV)) box.lineDiag = true
  } else {
    // zero-height extent = Word's horizontal line; keep a 12 px grab band
    box.heightPx = 12
  }
  // a:headEnd decorates the start point: a head-only arrow renders as the
  // reversed segment so the renderer's single arrowhead lands on the right tip
  if (head && !tail && box.prst === 'lineArrow') {
    const fh = !box.flipH
    const fv = !box.flipV
    delete box.flipH
    delete box.flipV
    if (fh) box.flipH = true
    if (fv) box.flipV = true
  }
  box.insetTopPx = 0
  box.insetRightPx = 0
  box.insetBottomPx = 0
  box.insetLeftPx = 0
  return box
}

/** a:schemeClr val -> ThemeColors slot (DrawingML names; text/bg aliases mapped) */
const SCHEME_CLR_SLOTS: Record<string, keyof ThemeColors> = {
  tx1: 'dk1',
  bg1: 'lt1',
  tx2: 'dk2',
  bg2: 'lt2',
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
}

const PRST_CLR_HEX: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  gray: '808080',
}

/** one a:gs stop -> sRGB triple (srgbClr/sysClr/prstClr or theme-resolved schemeClr, lumMod/lumOff applied) */
function gradStopRgb(gs: XNode, theme?: ThemeColors | null): number[] | null {
  const srgb = attrsOf(findChild(gs, 'a:srgbClr') ?? {})['val']
  let base: string | undefined = srgb
  if (!base) {
    const sys = findChild(gs, 'a:sysClr')
    if (sys) {
      // Word writes the resolved system color into lastClr
      const a = attrsOf(sys)
      base = a['lastClr'] ?? (a['val'] === 'windowText' ? '000000' : 'FFFFFF')
    }
  }
  if (!base) {
    const prst = findChild(gs, 'a:prstClr')
    if (prst) base = PRST_CLR_HEX[attrsOf(prst)['val'] ?? '']
  }
  const scheme = base ? undefined : findChild(gs, 'a:schemeClr')
  if (!base && scheme) {
    const slot = SCHEME_CLR_SLOTS[attrsOf(scheme)['val'] ?? '']
    if (!slot) return null
    base = (theme?.[slot] as string | undefined) ?? DEFAULT_THEME_COLORS[slot]
  }
  const baseTrimmed = (base ?? '').trim()
  if (!baseTrimmed || !/^[0-9A-Fa-f]{6}$/.test(baseTrimmed)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(baseTrimmed.slice(i, i + 2), 16))
  if (scheme) {
    const pct = (name: string): number | null => {
      const v = parseInt(attrsOf(findChild(scheme, name) ?? {})['val'] ?? '', 10)
      return Number.isFinite(v) ? Math.min(100000, Math.max(0, v)) / 100000 : null
    }
    const lumMod = pct('a:lumMod')
    if (lumMod !== null) rgb = rgb.map((c) => c * lumMod)
    const lumOff = pct('a:lumOff')
    if (lumOff !== null) rgb = rgb.map((c) => c + 255 * lumOff)
    const shade = pct('a:shade')
    if (shade !== null) rgb = rgb.map((c) => c * shade)
    const tint = pct('a:tint')
    if (tint !== null) rgb = rgb.map((c) => c * tint + 255 * (1 - tint))
  }
  return rgb
}

/** color-bearing node (a:solidFill / a:fillRef / a:lnRef …) → hex without '#' */
export function colorNodeHex(
  node: XNode | undefined,
  theme?: ThemeColors | null,
): string | undefined {
  if (!node) return undefined
  const rgb = gradStopRgb(node, theme)
  if (!rgb) return undefined
  return rgb
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

const COLOR_NODE_NAMES = ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr']

/** a:alpha of a solid fill's color as 0..1; undefined when opaque or absent */
export function solidFillAlpha(solidFill: XNode | undefined): number | undefined {
  if (!solidFill) return undefined
  for (const name of COLOR_NODE_NAMES) {
    const color = findChild(solidFill, name)
    if (!color) continue
    const raw = parseInt(attrsOf(findChild(color, 'a:alpha') ?? {})['val'] ?? '', 10)
    return Number.isFinite(raw) && raw >= 0 && raw < 100000 ? raw / 100000 : undefined
  }
  return undefined
}

/** anchored-drawing placement of one top-level w:drawing fragment */
export interface DrawingAnchorMeta {
  offsetXEmu?: number
  offsetYEmu?: number
  /** wrapNone (front) / behindDoc anchors leave the text flow entirely */
  noWrap?: boolean
  /** behindDoc="1": paints under the body text (z-order only) */
  behind?: boolean
  /** raw page coordinates for a first-page page-anchored cover drawing: the
   *  renderer resolves the boxes against the page box, not the paragraph */
  pageXEmu?: number
  pageYEmu?: number
  /** wp:wrapTopAndBottom: body text is excluded from the drawing's vertical band */
  topBottom?: boolean
  /** layoutInCell="0": a cell-anchored drawing positions against the page/column, not its cell */
  outsideCell?: boolean
  anchored?: boolean
  /** wp:positionH/V relativeFrom */
  relH?: string
  relV?: string
  /** wp:align inside wp:positionH/V */
  alignH?: string
  alignV?: string
  /** wp14:pctPosH/VOffset in 1/1000 of a percent of the reference frame */
  pctH?: number
  pctV?: number
  /** wp:extent (drawing size) */
  extentXEmu?: number
  extentYEmu?: number
  /** wp:anchor distL/distR/distT/distB: text clearance around a side-wrapped drawing */
  distLEmu?: number
  distREmu?: number
  distTEmu?: number
  distBEmu?: number
  /** relativeHeight − Word's 251658240 base: paint order among overlapping anchors */
  z?: number
}

/**
 * Top-level `<w:drawing>` fragments of a paragraph, balanced (a textbox's
 * txbxContent may nest further drawings, which stay inside their parent
 * fragment). Used to attach each drawing's own anchor placement to the shapes
 * it carries — a paragraph can anchor several drawings at distinct offsets.
 */
/** next `<w:drawing>` open tag at or after `from`, attributes tolerated
 *  (`<w:drawing mc:MustUnderstand="wps">`, TestFiles mcdoc) */
function drawingOpenAt(xml: string, from: number): number {
  const re = /<w:drawing[\s>]/g
  re.lastIndex = from
  return re.exec(xml)?.index ?? -1
}

export function topLevelDrawings(xml: string): string[] {
  const out: string[] = []
  let i = 0
  for (;;) {
    const start = drawingOpenAt(xml, i)
    if (start === -1) break
    let depth = 0
    let j = start
    for (;;) {
      const open = drawingOpenAt(xml, j + 1)
      const close = xml.indexOf('</w:drawing>', j + 1)
      if (close === -1) return out // malformed; bail with what we have
      if (open !== -1 && open < close) {
        depth++
        j = open
      } else if (depth > 0) {
        depth--
        j = close
      } else {
        out.push(xml.slice(start, close + '</w:drawing>'.length))
        i = close + '</w:drawing>'.length
        break
      }
    }
  }
  return out
}

export function drawingAnchorMeta(frag: string): DrawingAnchorMeta {
  const anchorTag = /<wp:anchor[^>]*>/.exec(frag)?.[0]
  if (!anchorTag) return {}
  const meta: DrawingAnchorMeta = { anchored: true }
  const posOf = (dir: 'H' | 'V'): number | undefined => {
    const m = new RegExp(`<wp:position${dir}[^>]*>\\s*<wp:posOffset>(-?\\d+)</wp:posOffset>`).exec(
      frag,
    )
    const v = m ? parseInt(m[1], 10) : NaN
    return Number.isFinite(v) ? v : undefined
  }
  meta.offsetXEmu = posOf('H')
  meta.offsetYEmu = posOf('V')
  for (const side of ['L', 'R', 'T', 'B'] as const) {
    const v = parseInt(
      new RegExp(`\\bdist${side}\\s*=\\s*["'](\\d+)["']`).exec(anchorTag)?.[1] ?? '',
      10,
    )
    if (Number.isFinite(v)) meta[`dist${side}Emu`] = v
  }
  for (const dir of ['H', 'V'] as const) {
    const m = new RegExp(`<wp:position${dir}\\b([^>]*)>([\\s\\S]*?)</wp:position${dir}>`).exec(frag)
    if (!m) continue
    const rel = /relativeFrom\s*=\s*["'](\w+)["']/.exec(m[1])?.[1]
    const align = /<wp:align>(\w+)<\/wp:align>/.exec(m[2])?.[1]
    const pct = parseInt(
      new RegExp(`<wp14:pctPos${dir}Offset[^>]*>(-?\\d+)<`).exec(m[2])?.[1] ?? '',
      10,
    )
    if (dir === 'H') {
      meta.relH = rel
      meta.alignH = align
      if (Number.isFinite(pct)) meta.pctH = pct
    } else {
      meta.relV = rel
      meta.alignV = align
      if (Number.isFinite(pct)) meta.pctV = pct
    }
  }
  const extentTag = /<wp:extent[^>]*\/?>/.exec(frag)?.[0] ?? ''
  const extentX = parseInt(/\bcx\s*=\s*["'](\d+)["']/.exec(extentTag)?.[1] ?? '', 10)
  const extentY = parseInt(/\bcy\s*=\s*["'](\d+)["']/.exec(extentTag)?.[1] ?? '', 10)
  if (Number.isFinite(extentX)) meta.extentXEmu = extentX
  if (Number.isFinite(extentY)) meta.extentYEmu = extentY
  // the anchor's own wrap element sits before a:graphic; a nested drawing's
  // wrap must not leak up. behindDoc is z-order only: a behind-text picture
  // with wrapTight/wrapSquare/wrapTopAndBottom still excludes the text
  const graphicAt = frag.indexOf('<a:graphic')
  const ownXml = graphicAt === -1 ? frag : frag.slice(0, graphicAt)
  const behind = /behindDoc\s*=\s*["'](?:1|true)["']/.test(anchorTag)
  const ownWrapped = /<wp:wrap(?:Square|Tight|Through|TopAndBottom)\b/.test(ownXml)
  if (frag.includes('<wp:wrapNone') || (behind && !ownWrapped)) meta.noWrap = true
  if (behind) meta.behind = true
  const relHeight = Number(/relativeHeight\s*=\s*["'](\d+)["']/.exec(anchorTag)?.[1] ?? NaN)
  if (Number.isFinite(relHeight) && relHeight - 251658240 !== 0) meta.z = relHeight - 251658240
  if (ownXml.includes('<wp:wrapTopAndBottom')) meta.topBottom = true
  if (/layoutInCell\s*=\s*["'](?:0|false)["']/.test(anchorTag)) meta.outsideCell = true
  return meta
}

export const EMU_PER_TWIP = 635

/** narrowest column gap beside a float that Word still fills with text (36px) */
export const MIN_WRAP_SLIVER_EMU = 36 * 9525
/** Word's default distL/distR clearance (0.125 in) */
export const DEFAULT_WRAP_DIST_EMU = 114300

/** anchor position resolved to EMU offsets from the paragraph flow origin
 *  (column left / body top), page coordinates recoverable via the margins */
export interface ResolvedAnchorPos {
  xEmu: number
  yEmu?: number
  /** the drawing lies horizontally outside the body column (sidebar layout) */
  outsideColumn: boolean
}

/**
 * Page/margin-anchored placement (wp14:pctPosH/VOffset or wp:align) resolved
 * against the section's page geometry. Plain posOffset anchors keep the
 * legacy paragraph-relative path untouched — only the features the parser
 * previously ignored resolve here, so in-column wrapSquare boxes are not
 * repositioned. The vertical origin approximates the anchor paragraph sitting
 * at the top of the body (where Word puts these full-height sidebar groups).
 */
export function resolveAnchorPagePos(
  meta: DrawingAnchorMeta,
  sect: SectionSettings | undefined,
): ResolvedAnchorPos | null {
  if (!sect) return null
  // column-relative align in a single-column section: the column IS the margin
  // box, so it resolves identically (multi-column needs the hosting column and
  // stays on the legacy flow placement)
  const relH =
    meta.relH === 'column' && sect.columns <= 1 && meta.pctH === undefined ? 'margin' : meta.relH
  if (relH !== 'page' && relH !== 'margin') return null
  if (meta.pctH === undefined && meta.alignH === undefined) return null
  const pageW = sect.pageWidth * EMU_PER_TWIP
  const marL = sect.marginLeft * EMU_PER_TWIP
  const marR = sect.marginRight * EMU_PER_TWIP
  const w = meta.extentXEmu ?? 0
  const refW = relH === 'page' ? pageW : pageW - marL - marR
  const relX =
    meta.pctH !== undefined
      ? Math.round((refW * meta.pctH) / 100000)
      : meta.alignH === 'center'
        ? Math.round((refW - w) / 2)
        : meta.alignH === 'right' || meta.alignH === 'outside'
          ? refW - w
          : 0
  const pageX = relH === 'page' ? relX : marL + relX
  const pos: ResolvedAnchorPos = {
    xEmu: pageX - marL,
    outsideColumn: pageX + w <= marL || pageX >= pageW - marR,
  }
  const yEmu = resolveAnchorPageYEmu(meta, sect)
  if (yEmu !== undefined) pos.yEmu = yEmu
  return pos
}

/**
 * Page/margin-relative V placement (wp14:pctPosVOffset or wp:align) as an EMU
 * offset from the body top, independent of how X is anchored. Undefined for
 * plain posOffset or paragraph/line-relative anchors.
 */
export function resolveAnchorPageYEmu(
  meta: DrawingAnchorMeta,
  sect: SectionSettings | undefined,
): number | undefined {
  if (!sect) return undefined
  if (meta.relV !== 'page' && meta.relV !== 'margin') return undefined
  if (meta.pctV === undefined && meta.alignV === undefined) return undefined
  const pageH = sect.pageHeight * EMU_PER_TWIP
  const marT = sect.marginTop * EMU_PER_TWIP
  const marB = sect.marginBottom * EMU_PER_TWIP
  const h = meta.extentYEmu ?? 0
  const refH = meta.relV === 'page' ? pageH : pageH - marT - marB
  const relY =
    meta.pctV !== undefined
      ? Math.round((refH * meta.pctV) / 100000)
      : meta.alignV === 'center'
        ? Math.round((refH - h) / 2)
        : meta.alignV === 'bottom' || meta.alignV === 'outside'
          ? refH - h
          : 0
  return (meta.relV === 'page' ? relY : marT + relY) - marT
}

/**
 * a:gradFill approximated as a solid color (display only): equal-weight sRGB
 * average of all stops — the first stop alone is often white and loses the
 * visible tint entirely.
 */
export function gradFillApproxHex(spPr: XNode, theme?: ThemeColors | null): string | undefined {
  const gsLst = findChild(findChild(spPr, 'a:gradFill') ?? {}, 'a:gsLst')
  if (!gsLst) return undefined
  const stops = childrenOf(gsLst)
    .filter((n) => nameOf(n) === 'a:gs')
    .map((gs) => gradStopRgb(gs, theme))
    .filter((rgb): rgb is number[] => rgb !== null)
  if (stops.length === 0) return undefined
  return [0, 1, 2]
    .map((i) => stops.reduce((sum, rgb) => sum + rgb[i], 0) / stops.length)
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

/** sRGB triple → hex without '#' */
function rgbHex(rgb: number[]): string {
  return rgb
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

/** w14:srgbClr / w14:schemeClr child → sRGB triple with tint/shade/lum/satMod applied */
function w14ColorRgb(node: XNode, theme?: ThemeColors | null): number[] | null {
  const colorNode = findChild(node, 'w14:srgbClr') ?? findChild(node, 'w14:schemeClr')
  if (!colorNode) return null
  const isScheme = nameOf(colorNode) === 'w14:schemeClr'
  let base: string | undefined = attrsOf(colorNode)['w14:val']
  if (isScheme) {
    const slot = SCHEME_CLR_SLOTS[base ?? '']
    if (!slot) return null
    base = (theme?.[slot] as string | undefined) ?? DEFAULT_THEME_COLORS[slot]
  }
  const baseTrimmed = (base ?? '').trim()
  if (!baseTrimmed || !/^[0-9A-Fa-f]{6}$/.test(baseTrimmed)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(baseTrimmed.slice(i, i + 2), 16))
  const pct = (name: string): number | null => {
    const v = parseInt(attrsOf(findChild(colorNode, name) ?? {})['w14:val'] ?? '', 10)
    return Number.isFinite(v) && v >= 0 ? v / 100000 : null
  }
  const lumMod = pct('w14:lumMod')
  if (lumMod !== null) rgb = rgb.map((c) => c * lumMod)
  const lumOff = pct('w14:lumOff')
  if (lumOff !== null) rgb = rgb.map((c) => c + 255 * lumOff)
  const shade = pct('w14:shade')
  if (shade !== null) rgb = rgb.map((c) => c * shade)
  const tint = pct('w14:tint')
  if (tint !== null) rgb = rgb.map((c) => c * tint + 255 * (1 - tint))
  const satMod = pct('w14:satMod')
  if (satMod !== null && satMod !== 1) rgb = saturationModulate(rgb, satMod)
  return rgb
}

/** HSL saturation modulation (a:satMod / w14:satMod semantics, clamped) */
function saturationModulate(rgb: number[], mod: number): number[] {
  const [r, g, b] = rgb.map((c) => Math.min(255, Math.max(0, c)) / 255) as [number, number, number]
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (d === 0) return rgb
  let s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h =
    max === r
      ? ((g - b) / d + (g < b ? 6 : 0)) / 6
      : max === g
        ? ((b - r) / d + 2) / 6
        : ((r - g) / d + 4) / 6
  s = Math.min(1, s * mod)
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const hue = (t: number): number => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)].map((c) => c * 255)
}

/**
 * WordArt-styled run text (w14:textFill): solid fill directly, gradient fill
 * as the equal-weight average of its stops — display approximation only.
 */
export function w14TextFillHex(rPr: XNode, theme?: ThemeColors | null): string | undefined {
  const tf = findChild(rPr, 'w14:textFill')
  if (!tf) return undefined
  const solid = findChild(tf, 'w14:solidFill')
  if (solid) {
    const rgb = w14ColorRgb(solid, theme)
    return rgb ? rgbHex(rgb) : undefined
  }
  const gsLst = findChild(findChild(tf, 'w14:gradFill') ?? {}, 'w14:gsLst')
  if (!gsLst) return undefined
  const stops = childrenOf(gsLst)
    .filter((n) => nameOf(n) === 'w14:gs')
    .map((gs) => w14ColorRgb(gs, theme))
    .filter((rgb): rgb is number[] => rgb !== null)
  if (stops.length === 0) return undefined
  return rgbHex([0, 1, 2].map((i) => stops.reduce((sum, rgb) => sum + rgb[i], 0) / stops.length))
}

/** w14:alpha is how transparent the colour is: the opacity it leaves, when below 1 */
function w14OpacityOf(node: XNode): number | undefined {
  const colorNode = findChild(node, 'w14:srgbClr') ?? findChild(node, 'w14:schemeClr')
  const raw = parseInt(attrsOf(findChild(colorNode ?? {}, 'w14:alpha') ?? {})['w14:val'] ?? '', 10)
  return raw > 0 && raw <= 100000 ? Math.round((1 - raw / 100000) * 1000) / 1000 : undefined
}

/** w14:textFill as it is: no fill, a colour with its transparency, or a gradient with its stops */
export function w14TextFillOf(rPr: XNode, theme?: ThemeColors | null): TextFill | undefined {
  const tf = findChild(rPr, 'w14:textFill')
  if (!tf) return undefined
  if (findChild(tf, 'w14:noFill')) return { kind: 'none' }
  const solid = findChild(tf, 'w14:solidFill')
  if (solid) {
    const rgb = w14ColorRgb(solid, theme)
    if (!rgb) return undefined
    const alpha = w14OpacityOf(solid)
    return { kind: 'solid', color: rgbHex(rgb), ...(alpha !== undefined ? { alpha } : {}) }
  }
  const grad = findChild(tf, 'w14:gradFill')
  const gsLst = findChild(grad ?? {}, 'w14:gsLst')
  if (!grad || !gsLst) return undefined
  const stops: TextFillStop[] = []
  for (const gs of childrenOf(gsLst)) {
    if (nameOf(gs) !== 'w14:gs') continue
    const rgb = w14ColorRgb(gs, theme)
    const pos = parseInt(attrsOf(gs)['w14:pos'] ?? '', 10)
    if (!rgb || !(pos >= 0)) continue
    const alpha = w14OpacityOf(gs)
    stops.push({
      pos: Math.min(1, pos / 100000),
      color: rgbHex(rgb),
      ...(alpha !== undefined ? { alpha } : {}),
    })
  }
  if (stops.length === 0) return undefined
  stops.sort((a, b) => a.pos - b.pos)
  const ang = parseInt(attrsOf(findChild(grad, 'w14:lin') ?? {})['w14:ang'] ?? '', 10)
  const path = attrsOf(findChild(grad, 'w14:path') ?? {})['w14:path']
  return {
    kind: 'gradient',
    stops,
    angleDeg: Number.isFinite(ang) ? Math.round(ang / 600) / 100 : 0,
    ...(path === 'circle' || path === 'rect' || path === 'shape' ? { path } : {}),
  }
}

/** the fill where it shows something a text colour of `color` would not */
export function textFillBeyondColor(
  fill: TextFill | undefined,
  color: string | undefined,
): TextFill | undefined {
  if (!fill) return undefined
  if (fill.kind === 'solid' && fill.alpha === undefined) {
    return color !== undefined && fill.color.toUpperCase() !== color.toUpperCase()
      ? fill
      : undefined
  }
  return fill
}

/** w14:textOutline stroke (solid fills only; gradient/noFill outlines are not drawn) */
export function w14TextOutlineOf(rPr: XNode, theme?: ThemeColors | null): TextOutline | undefined {
  const outline = findChild(rPr, 'w14:textOutline')
  if (!outline) return undefined
  const solid = findChild(outline, 'w14:solidFill')
  if (!solid) return undefined
  const rgb = w14ColorRgb(solid, theme)
  if (!rgb) return undefined
  const widthEmu = parseInt(attrsOf(outline)['w14:w'] ?? '', 10)
  if (!(widthEmu > 0)) return undefined
  const colorNode = findChild(solid, 'w14:srgbClr') ?? findChild(solid, 'w14:schemeClr')
  const alphaRaw = parseInt(
    attrsOf(findChild(colorNode ?? {}, 'w14:alpha') ?? {})['w14:val'] ?? '',
    10,
  )
  const alpha = alphaRaw >= 0 && alphaRaw < 100000 ? alphaRaw / 100000 : undefined
  return {
    color: rgbHex(rgb),
    widthPt: Math.round((widthEmu / 12700) * 100) / 100,
    ...(alpha !== undefined ? { alpha } : {}),
  }
}

/** w14:glow halo (radius in EMU, color with optional alpha) */
export function w14GlowOf(rPr: XNode, theme?: ThemeColors | null): TextGlow | undefined {
  const glow = findChild(rPr, 'w14:glow')
  if (!glow) return undefined
  const rgb = w14ColorRgb(glow, theme)
  const radEmu = parseInt(attrsOf(glow)['w14:rad'] ?? '', 10)
  if (!rgb || !(radEmu > 0)) return undefined
  const colorNode = findChild(glow, 'w14:srgbClr') ?? findChild(glow, 'w14:schemeClr')
  const alphaRaw = parseInt(
    attrsOf(findChild(colorNode ?? {}, 'w14:alpha') ?? {})['w14:val'] ?? '',
    10,
  )
  const alpha = alphaRaw >= 0 && alphaRaw < 100000 ? alphaRaw / 100000 : undefined
  return {
    color: rgbHex(rgb),
    radiusPt: Math.round((radEmu / 12700) * 100) / 100,
    ...(alpha !== undefined ? { alpha } : {}),
  }
}

/** w14:shadow outer shadow (blur and distance in EMU, direction in 60000ths of a degree) */
export function w14ShadowOf(rPr: XNode, theme?: ThemeColors | null): TextShadow | undefined {
  const shadow = findChild(rPr, 'w14:shadow')
  if (!shadow) return undefined
  const rgb = w14ColorRgb(shadow, theme)
  if (!rgb) return undefined
  const attrs = attrsOf(shadow)
  const emu = (name: string) => parseInt(attrs[name] ?? '', 10) || 0
  const colorNode = findChild(shadow, 'w14:srgbClr') ?? findChild(shadow, 'w14:schemeClr')
  const alphaRaw = parseInt(
    attrsOf(findChild(colorNode ?? {}, 'w14:alpha') ?? {})['w14:val'] ?? '',
    10,
  )
  const alpha = alphaRaw >= 0 && alphaRaw < 100000 ? alphaRaw / 100000 : undefined
  return {
    color: rgbHex(rgb),
    ...(alpha !== undefined ? { alpha } : {}),
    blurPt: Math.round((emu('w14:blurRad') / 12700) * 100) / 100,
    distPt: Math.round((emu('w14:dist') / 12700) * 100) / 100,
    dirDeg: Math.round(emu('w14:dir') / 60000),
  }
}

/** w14:reflection (alphas and positions in 1000ths of a percent, blur and distance in EMU) */
export function w14ReflectionOf(rPr: XNode): TextReflection | undefined {
  const reflection = findChild(rPr, 'w14:reflection')
  if (!reflection) return undefined
  const attrs = attrsOf(reflection)
  const num = (name: string) => parseInt(attrs[name] ?? '', 10) || 0
  const startAlpha = num('w14:stA') / 100000
  const endPos = num('w14:endPos') / 100000
  if (!(startAlpha > 0) || !(endPos > 0)) return undefined
  return {
    startAlpha: Math.min(1, startAlpha),
    endPos: Math.min(1, endPos),
    blurPt: Math.round((num('w14:blurRad') / 12700) * 100) / 100,
    distPt: Math.round((num('w14:dist') / 12700) * 100) / 100,
  }
}

export interface ExtractTextboxOpts {
  /** include textless preset shapes (stars, block arrows…) as display boxes */
  shapes?: boolean
  /** include picture-only drawings (pic:pic without a shape) as picture boxes */
  pictures?: boolean
  /** page geometry of the governing section (page/margin-anchored placement) */
  section?: SectionSettings
  /** document.xml byte offset of the host paragraph: textbox paragraphs resolve
   *  their character-unit indents under that section's document grid */
  docOffset?: number
  /** the paragraph sits before the first explicit page break: page-V-anchored
   *  cover art may pin to the page origin instead of the paragraph origin */
  firstPage?: boolean
}

/** child-EMU → anchor-EMU affine transform of a wpg group (X = tx + x·sx) */
export interface GroupCtm {
  sx: number
  sy: number
  tx: number
  ty: number
  /** the groups from the outermost in, once one of them is turned or mirrored (the affine part cannot say that) */
  turned?: GroupLevel[]
}

interface GroupLevel {
  group: GroupXfrm
  chOff: { x: number; y: number }
  chExt: { cx: number; cy: number }
}

export const IDENTITY_CTM: GroupCtm = { sx: 1, sy: 1, tx: 0, ty: 0 }

function xfrmNumbers(xfrm: XNode): GroupXfrm {
  const num = (node: XNode | undefined, key: string): number => {
    const v = parseInt(attrsOf(node ?? {})[key] ?? '', 10)
    return Number.isFinite(v) ? v : 0
  }
  const on = (key: string): boolean => /^(?:1|true)$/.test(attrsOf(xfrm)[key] ?? '')
  return {
    x: num(findChild(xfrm, 'a:off'), 'x'),
    y: num(findChild(xfrm, 'a:off'), 'y'),
    cx: num(findChild(xfrm, 'a:ext'), 'cx'),
    cy: num(findChild(xfrm, 'a:ext'), 'cy'),
    rot: num(xfrm, 'rot'),
    flipH: on('flipH'),
    flipV: on('flipV'),
  }
}

/** wpg:wgp / wpg:grpSp xfrm (off/ext vs chOff/chExt) composed onto `outer` */
export function composeGroupCtm(group: XNode, outer: GroupCtm): GroupCtm | null {
  const xfrm = findChild(findChild(group, 'wpg:grpSpPr') ?? {}, 'a:xfrm')
  if (!xfrm) return null
  const num = (node: XNode | undefined, key: string, dflt: number): number => {
    const v = parseInt(attrsOf(node ?? {})[key] ?? '', 10)
    return Number.isFinite(v) ? v : dflt
  }
  const own = xfrmNumbers(xfrm)
  const { x: ox, y: oy, cx: ex, cy: ey } = own
  const chOffX = num(findChild(xfrm, 'a:chOff'), 'x', 0)
  const chOffY = num(findChild(xfrm, 'a:chOff'), 'y', 0)
  const chExtX = num(findChild(xfrm, 'a:chExt'), 'cx', 0)
  const chExtY = num(findChild(xfrm, 'a:chExt'), 'cy', 0)
  const sx = ex > 0 && chExtX > 0 ? ex / chExtX : 1
  const sy = ey > 0 && chExtY > 0 ? ey / chExtY : 1
  const level: GroupLevel = {
    group: own,
    chOff: { x: chOffX, y: chOffY },
    chExt: { cx: chExtX > 0 ? chExtX : ex, cy: chExtY > 0 ? chExtY : ey },
  }
  const turns = own.rot !== 0 || own.flipH || own.flipV
  const turned = outer.turned
    ? [...outer.turned, level]
    : turns
      ? [...plainLevels(outer), level]
      : undefined
  return {
    sx: outer.sx * sx,
    sy: outer.sy * sy,
    tx: outer.tx + outer.sx * (ox - chOffX * sx),
    ty: outer.ty + outer.sy * (oy - chOffY * sy),
    ...(turned ? { turned } : {}),
  }
}

/** the unturned groups `ctm` stands for, as one level */
function plainLevels(ctm: GroupCtm): GroupLevel[] {
  if (ctm.sx === 1 && ctm.sy === 1 && ctm.tx === 0 && ctm.ty === 0) return []
  return [
    {
      group: { x: ctm.tx, y: ctm.ty, cx: ctm.sx, cy: ctm.sy, rot: 0, flipH: false, flipV: false },
      chOff: { x: 0, y: 0 },
      chExt: { cx: 1, cy: 1 },
    },
  ]
}

/** Where a child of turned or mirrored groups stands, placed as Ungroup writes it; null when none is turned. */
export function placedInGroups(xfrm: XNode | undefined, ctm: GroupCtm | null): GroupXfrm | null {
  if (!xfrm || !ctm?.turned) return null
  let placed = xfrmNumbers(xfrm)
  for (let i = ctm.turned.length - 1; i >= 0; i--) {
    const level = ctm.turned[i]
    placed = placedOutsideGroup(placed, level.group, level.chOff, level.chExt)
  }
  return placed
}

export function placeGroupedBox(
  box: {
    offsetXEmu?: number
    offsetYEmu?: number
    rotDeg?: number
    flipH?: boolean
    flipV?: boolean
  },
  placed: GroupXfrm,
): void {
  box.offsetXEmu = Math.round(placed.x)
  box.offsetYEmu = Math.round(placed.y)
  const deg = Math.round(placed.rot / 60000) % 360
  if (deg) box.rotDeg = deg
  else delete box.rotDeg
  if (placed.flipH) box.flipH = true
  else delete box.flipH
  if (placed.flipV) box.flipV = true
  else delete box.flipV
}

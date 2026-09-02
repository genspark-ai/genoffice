// DrawingML geometry and color helpers: line presets, scheme/preset colors,
// anchor metadata and page-position resolution, group transforms.
import { attrsOf, childrenOf, findChild, nameOf, type XNode } from './xml-utils'
import { EMU_PER_PX } from './parse-xml-text'
import type { SectionSettings, TextboxDisplay, ThemeColors } from './types'

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

/** wps line shape → display-only line box (synthetic prst carries the arrow ends) */
export function lineBoxOf(shape: XNode, theme?: ThemeColors | null): TextboxDisplay | null {
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
  if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round(cx / EMU_PER_PX)
  const straight = box.prst === 'line' || box.prst === 'lineArrow' || box.prst === 'lineArrowDouble'
  // flips apply at any height: Word's horizontal lines carry cy="0", and a
  // flipH there still reverses which tip holds the arrow
  if (straight) {
    if (xfrmAttrs['flipH'] === '1' || xfrmAttrs['flipH'] === 'true') box.flipH = true
    if (xfrmAttrs['flipV'] === '1' || xfrmAttrs['flipV'] === 'true') box.flipV = true
  }
  if (Number.isFinite(cy) && cy > 0) {
    box.heightPx = Math.round(cy / EMU_PER_PX)
    box.minHeightPx = box.heightPx
    // a real vertical extent means the connector runs corner to corner
    // (≤12 px stays level: our own inserted lines keep a 12 px grab band)
    if (straight && (box.heightPx > 12 || box.flipH || box.flipV)) box.lineDiag = true
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
    base =
      (theme?.[slot] as string | undefined) ??
      (slot === 'dk1' ? '000000' : slot === 'lt1' ? 'FFFFFF' : undefined)
  }
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(base!.slice(i, i + 2), 16))
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
  for (const dir of ['H', 'V'] as const) {
    const m = new RegExp(`<wp:position${dir}\\b([^>]*)>([\\s\\S]*?)</wp:position${dir}>`).exec(frag)
    if (!m) continue
    const rel = /relativeFrom="(\w+)"/.exec(m[1])?.[1]
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
  const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(frag)
  if (extent) {
    meta.extentXEmu = parseInt(extent[1], 10)
    meta.extentYEmu = parseInt(extent[2], 10)
  }
  if (frag.includes('<wp:wrapNone') || /behindDoc="(?:1|true)"/.test(anchorTag)) meta.noWrap = true
  if (/behindDoc="(?:1|true)"/.test(anchorTag)) meta.behind = true
  const relHeight = Number(/relativeHeight="(\d+)"/.exec(anchorTag)?.[1] ?? NaN)
  if (Number.isFinite(relHeight) && relHeight - 251658240 !== 0) meta.z = relHeight - 251658240
  // the anchor's own wrap element sits before a:graphic; a nested drawing's
  // wrap must not leak up. behindDoc="1" + wrapTopAndBottom coexist in
  // generated docs — Word still excludes the band (behindDoc is z-order only)
  const graphicAt = frag.indexOf('<a:graphic')
  const ownXml = graphicAt === -1 ? frag : frag.slice(0, graphicAt)
  if (ownXml.includes('<wp:wrapTopAndBottom')) meta.topBottom = true
  return meta
}

export const EMU_PER_TWIP = 635

/** narrowest column gap beside a float that Word still fills with text (36px) */
export const MIN_WRAP_SLIVER_EMU = 36 * 9525

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
  const pageH = sect.pageHeight * EMU_PER_TWIP
  const marL = sect.marginLeft * EMU_PER_TWIP
  const marR = sect.marginRight * EMU_PER_TWIP
  const marT = sect.marginTop * EMU_PER_TWIP
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
  if (
    (meta.relV === 'page' || meta.relV === 'margin') &&
    (meta.pctV !== undefined || meta.alignV !== undefined)
  ) {
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
    pos.yEmu = (meta.relV === 'page' ? relY : marT + relY) - marT
  }
  return pos
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
    base =
      (theme?.[slot] as string | undefined) ??
      (slot === 'dk1' ? '000000' : slot === 'lt1' ? 'FFFFFF' : undefined)
  }
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(base!.slice(i, i + 2), 16))
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
}

export const IDENTITY_CTM: GroupCtm = { sx: 1, sy: 1, tx: 0, ty: 0 }

/** wpg:wgp / wpg:grpSp xfrm (off/ext vs chOff/chExt) composed onto `outer` */
export function composeGroupCtm(group: XNode, outer: GroupCtm): GroupCtm | null {
  const xfrm = findChild(findChild(group, 'wpg:grpSpPr') ?? {}, 'a:xfrm')
  if (!xfrm) return null
  const num = (node: XNode | undefined, key: string, dflt: number): number => {
    const v = parseInt(attrsOf(node ?? {})[key] ?? '', 10)
    return Number.isFinite(v) ? v : dflt
  }
  const ox = num(findChild(xfrm, 'a:off'), 'x', 0)
  const oy = num(findChild(xfrm, 'a:off'), 'y', 0)
  const ex = num(findChild(xfrm, 'a:ext'), 'cx', 0)
  const ey = num(findChild(xfrm, 'a:ext'), 'cy', 0)
  const chOffX = num(findChild(xfrm, 'a:chOff'), 'x', 0)
  const chOffY = num(findChild(xfrm, 'a:chOff'), 'y', 0)
  const chExtX = num(findChild(xfrm, 'a:chExt'), 'cx', 0)
  const chExtY = num(findChild(xfrm, 'a:chExt'), 'cy', 0)
  const sx = ex > 0 && chExtX > 0 ? ex / chExtX : 1
  const sy = ey > 0 && chExtY > 0 ? ey / chExtY : 1
  return {
    sx: outer.sx * sx,
    sy: outer.sy * sy,
    tx: outer.tx + outer.sx * (ox - chOffX * sx),
    ty: outer.ty + outer.sy * (oy - chOffY * sy),
  }
}

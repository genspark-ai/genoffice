// Legacy VML (v:shape / v:group) geometry and color readers.
import { attrsOf, findChild, type XNode } from './xml-utils'
import { EMU_PER_PX } from './parse-xml-text'
import type { HfImage, Run, TextboxDisplay } from './types'

/** VML length ("46pt", "16in", "2cm", "12px"; unitless = pt) → CSS px */
export function vmlLengthPx(value: string | undefined): number | undefined {
  if (!value) return undefined
  const m = /^\s*(-?[\d.]+)(pt|px|in|mm|cm)?\s*$/.exec(value)
  if (!m) return undefined
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v)) return undefined
  switch (m[2] ?? 'pt') {
    case 'px':
      return v
    case 'in':
      return v * 96
    case 'mm':
      return (v / 25.4) * 96
    case 'cm':
      return (v / 2.54) * 96
    default:
      return (v * 96) / 72
  }
}

/** one declaration of a VML style attribute ("rotation:315;width:4in" → "4in") */
export function vmlStyleProp(style: string, key: string): string | undefined {
  return new RegExp(`(?:^|;)\\s*${key}:([^;]+)`).exec(style)?.[1]?.trim()
}

/** VML style="width:189.9pt;height:626pt" dimension → CSS px */
export function vmlStyleDimPx(style: string, key: 'width' | 'height'): number | undefined {
  const px = vmlLengthPx(vmlStyleProp(style, key))
  return px != null && px > 0 ? Math.round(px) : undefined
}

/** VML rotation ("315", "-45", "41366637fd" = 1/65536 degree units) → degrees clockwise in (0, 360) */
export function vmlRotationDeg(style: string): number | undefined {
  const m = /^(-?[\d.]+)(fd)?$/.exec(vmlStyleProp(style, 'rotation') ?? '')
  if (!m) return undefined
  let deg = parseFloat(m[1]!)
  if (m[2]) deg /= 65536
  if (!Number.isFinite(deg)) return undefined
  deg = Math.round((((deg % 360) + 360) % 360) * 100) / 100
  return deg > 0 ? deg : undefined
}

/**
 * position:absolute placement of a VML shape → header/footer float anchor
 * fields. mso-position-*-relative "text"/"column" is the paragraph box (the
 * margin box in a header); a keyword alignment wins over the margin offsets.
 */
export function vmlFloatAnchor(
  style: string,
  out: Pick<HfImage, 'posH' | 'posV' | 'posXPx' | 'posYPx' | 'posHRel' | 'posVRel'>,
): void {
  const relH = vmlStyleProp(style, 'mso-position-horizontal-relative')
  const relV = vmlStyleProp(style, 'mso-position-vertical-relative')
  out.posHRel = relH === 'page' ? 'page' : 'margin'
  out.posVRel =
    relV === 'page' ? 'page' : relV === 'text' || relV === 'line' ? 'paragraph' : 'margin'
  const posH = vmlStyleProp(style, 'mso-position-horizontal')
  const posV = vmlStyleProp(style, 'mso-position-vertical')
  if (posH === 'left' || posH === 'center' || posH === 'right') out.posH = posH
  else out.posXPx = Math.round(vmlLengthPx(vmlStyleProp(style, 'margin-left')) ?? 0)
  if (posV === 'top' || posV === 'center' || posV === 'bottom') out.posV = posV
  else out.posYPx = Math.round(vmlLengthPx(vmlStyleProp(style, 'margin-top')) ?? 0)
}

/** VML fixed-point or fraction ("32768f" = 0.5, ".5", "1") → number */
export function vmlFraction(value: string | undefined): number | undefined {
  const m = /^\s*(-?[\d.]+)(f)?\s*$/.exec(value ?? '')
  if (!m) return undefined
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v)) return undefined
  return m[2] ? v / 65536 : v
}

/** HTML color names VML attributes use ("silver", "blue"…) */
const VML_NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  silver: 'C0C0C0',
  gray: '808080',
  grey: '808080',
  maroon: '800000',
  olive: '808000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  fuchsia: 'FF00FF',
  lime: '00FF00',
  aqua: '00FFFF',
  cyan: '00FFFF',
  orange: 'FFA500',
}

/** VML color attr ("#dbe5f1", "#aaa", "#dbe5f1 [3204]", "silver") → hex without '#' */
export function vmlColorHex(value: string | undefined): string | undefined {
  if (!value) return undefined
  const v = value.trim()
  const m6 = /^#?([0-9a-fA-F]{6})/.exec(v)
  if (m6) return m6[1]
  const m3 = /^#([0-9a-fA-F]{3})(?![0-9a-fA-F])/.exec(v)
  if (m3) {
    return m3[1]
      .split('')
      .map((c) => c + c)
      .join('')
  }
  return VML_NAMED_COLORS[v.split(/[\s[]/, 1)[0]!.toLowerCase()]
}

/** VML WordArt: a shape carrying its text in a v:textpath string attribute */
export const VML_WORDART_RE = /<v:textpath[^>]*\bstring="/

/**
 * A v:imagedata that actually references a picture. WPS writes a bare
 * <v:imagedata o:title=""/> on ordinary textbox/geometry shapes; only an
 * r:id makes the pict a picture.
 */
export const VML_PICT_RID_RE = /<v:imagedata[^>]*\br:id="/

/** px per group-coordinate unit for children of a v:group (drawing canvas) */
export interface VmlGroupScale {
  sx: number
  sy: number
}

/** a group-child origin, px from the host paragraph */
export interface VmlOrigin {
  x: number
  y: number
}

export function vmlGroupScale(
  group: XNode,
  parentScale: VmlGroupScale | null = null,
): VmlGroupScale | null {
  const a = attrsOf(group)
  const style = a['style'] ?? ''
  // a nested group sizes in its parent's unitless coordinates
  const wPx = vmlShapeDimPx(style, 'width', parentScale)
  const hPx = vmlShapeDimPx(style, 'height', parentScale)
  const cs = /^\s*(-?\d+)[,\s]+(-?\d+)/.exec(a['coordsize'] ?? '')
  const cw = cs ? parseInt(cs[1]!, 10) : NaN
  const ch = cs ? parseInt(cs[2]!, 10) : NaN
  if (!wPx || !hPx || !(cw > 0) || !(ch > 0)) return null
  return { sx: wPx / cw, sy: hPx / ch }
}

/** style left/top of a group child → px from the paragraph (group origin + scaled coordinate) */
export function vmlCoordPx(
  style: string,
  key: 'left' | 'top',
  scale: VmlGroupScale,
  origin: VmlOrigin,
): number {
  const m = new RegExp(`(?:^|;)\\s*${key}:(-?[\\d.]+)(pt|px|in|mm|cm)?(?=;|$)`).exec(style)
  const base = key === 'left' ? origin.x : origin.y
  if (!m) return base
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v)) return base
  if (m[2]) {
    const px =
      m[2] === 'px'
        ? v
        : m[2] === 'in'
          ? v * 96
          : m[2] === 'mm'
            ? (v / 25.4) * 96
            : m[2] === 'cm'
              ? (v / 2.54) * 96
              : (v * 96) / 72
    return base + px
  }
  return base + v * (key === 'left' ? scale.sx : scale.sy)
}

/**
 * VML path attribute → normalized (0..1) SVG path. Straight-edge subset only
 * (m/l absolute, t/r relative, x close, e end) — curve or arc commands bail so
 * the caller draws nothing instead of a wrong solid box.
 */
export function vmlPathToNormD(path: string, cw: number, ch: number): string | undefined {
  if (!(cw > 0) || !(ch > 0)) return undefined
  const norm = (v: number, c: number): number => Math.round((v / c) * 10000) / 10000
  const parts: string[] = []
  let i = 0
  let cx = 0
  let cy = 0
  const readPairs = (): number[] | null => {
    // an omitted coordinate between separators means 0 ("m,l,21600...")
    const m = /^[-\d.,\s]+/.exec(path.slice(i))
    if (!m) return null
    i += m[0].length
    const nums = m[0]
      .trim()
      .split(/[,\s]/)
      .map((tok) => (tok === '' ? 0 : parseFloat(tok)))
    if (nums.some((v) => !Number.isFinite(v))) return null
    return nums.length > 0 && nums.length % 2 === 0 ? nums : null
  }
  while (i < path.length) {
    const c = path[i]!
    if (c === ' ' || c === ',') {
      i++
      continue
    }
    // nf/ns are fill/stroke hints, not geometry — before the 'e' check, or
    // an 'nf' would read as fatal
    if (path.startsWith('nf', i) || path.startsWith('ns', i)) {
      i += 2
      continue
    }
    if (c === 'e') {
      i++
      continue
    }
    if (c === 'x') {
      parts.push('Z')
      i++
      continue
    }
    if (c === 'm' || c === 'l' || c === 't' || c === 'r') {
      i++
      const nums = readPairs()
      if (!nums) return undefined
      const rel = c === 't' || c === 'r'
      const move = c === 'm' || c === 't'
      for (let k = 0; k < nums.length; k += 2) {
        cx = rel ? cx + nums[k]! : nums[k]!
        cy = rel ? cy + nums[k + 1]! : nums[k + 1]!
        parts.push(`${k === 0 && move ? 'M' : 'L'} ${norm(cx, cw)} ${norm(cy, ch)}`)
      }
      continue
    }
    return undefined
  }
  return parts.length > 1 ? parts.join(' ') : undefined
}

/** shape dimension → px: explicit units directly, unitless via the group scale */
export function vmlShapeDimPx(
  style: string,
  key: 'width' | 'height',
  scale: VmlGroupScale | null,
): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${key}:([0-9.]+)(pt|px|in|mm|cm)?`).exec(style)
  if (!m) return undefined
  if (m[2] || !scale) return vmlStyleDimPx(style, key)
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v) || v <= 0) return undefined
  return Math.round(v * (key === 'width' ? scale.sx : scale.sy))
}

/**
 * WordArt degrade: v:textpath shapes (shapetype 136 family) render their
 * string as plain styled text — no path warp / 3D, but the text is visible at
 * roughly the declared size and position instead of an opaque chip.
 */
export function vmlWordArtBox(shape: XNode): TextboxDisplay | null {
  const tp = findChild(shape, 'v:textpath')
  if (!tp) return null
  const text = attrsOf(tp)['string']
  if (!text || text.trim() === '') return null
  const shapeAttrs = attrsOf(shape)
  const style = shapeAttrs['style'] ?? ''
  const box: TextboxDisplay = {
    paras: [],
    readOnly: true,
    insetTopPx: 0,
    insetRightPx: 0,
    insetBottomPx: 0,
    insetLeftPx: 0,
  }
  const w = vmlStyleDimPx(style, 'width')
  if (w) box.widthPx = w
  const h = vmlStyleDimPx(style, 'height')
  if (h) box.heightPx = h
  // floating WordArt keeps the flow like other absolute shapes
  if (/position:\s*absolute/.test(style)) {
    box.floating = true
    const marginPx = (key: string): number => {
      const pt = parseFloat(new RegExp(`(?:^|;)\\s*${key}:(-?[\\d.]+)pt`).exec(style)?.[1] ?? '')
      return Number.isFinite(pt) ? (pt / 72) * 96 : 0
    }
    box.offsetXEmu = Math.round(marginPx('margin-left') * EMU_PER_PX)
    box.offsetYEmu = Math.round(marginPx('margin-top') * EMU_PER_PX)
  }
  const tpStyle = attrsOf(tp)['style'] ?? ''
  const family = /font-family:\s*"?([^;"]+)"?/.exec(tpStyle)?.[1]?.trim()
  const sizePt = parseFloat(/font-size:\s*([\d.]+)pt/.exec(tpStyle)?.[1] ?? '')
  // fill becomes the *text* color: fillcolor, else the v:fill color/color2
  const fillNode = findChild(shape, 'v:fill')
  const fillAttrs = fillNode ? attrsOf(fillNode) : {}
  const fill =
    shapeAttrs['filled'] === 'f'
      ? undefined
      : (vmlColorHex(shapeAttrs['fillcolor']) ??
        vmlColorHex(fillAttrs['color']) ??
        vmlColorHex(fillAttrs['color2']))
  if (shapeAttrs['stroked'] !== 'f') {
    const strokeColor = vmlColorHex(shapeAttrs['strokecolor']) ?? '000000'
    const weightPt = parseFloat(
      /^([\d.]+)(?:pt)?$/.exec(shapeAttrs['strokeweight'] ?? '')?.[1] ?? '',
    )
    box.textOutline = {
      colorHex: strokeColor,
      widthPx:
        Number.isFinite(weightPt) && weightPt > 0
          ? Math.round((weightPt / 72) * 96 * 100) / 100
          : 1,
    }
  }
  const heightPt = h ? (h / 96) * 72 : NaN
  const run: Run = { text }
  // fitshape sizes the glyphs to the box; the declared font-size matches it in
  // practice (box height ≈ font-size × line factor), so prefer the declared pt
  let pt = Number.isFinite(sizePt) && sizePt > 0 ? sizePt : heightPt > 0 ? heightPt / 1.4 : NaN
  // fitpath compresses long strings into the box; approximate by shrinking the
  // font until the single line fits the declared width (~0.62 em per glyph)
  const widthPt = w ? (w / 96) * 72 : NaN
  if (Number.isFinite(pt) && widthPt > 0 && text.length > 0) {
    pt = Math.max(6, Math.min(pt, widthPt / (0.62 * text.length)))
  }
  if (Number.isFinite(pt) && pt > 0) run.sizeHalfPoints = Math.round(pt * 2)
  box.nowrap = true
  if (family) run.fontAscii = family
  if (fill) run.color = fill
  if (/font-weight:\s*bold/.test(tpStyle)) run.bold = true
  if (/font-style:\s*italic/.test(tpStyle)) run.italic = true
  box.paras.push({ runs: [run], align: 'center' })
  return box
}

/** the shape element of a VML pict, its shapetype templates stripped */
function vmlPictShape(pict: string): RegExpExecArray | null {
  const body = pict.replace(/<v:shapetype\b[\s\S]*?<\/v:shapetype>/g, '')
  return /<v:(oval|rect|roundrect|shape)\b([^>]*?)(\/?)>([\s\S]*?<\/v:\1>)?/.exec(body)
}

/** v:shapetype templates of an XML part, by id (shapes point at them with type="#id") */
export function vmlShapeTypeTable(xml: string): Map<string, Record<string, string>> {
  const out = new Map<string, Record<string, string>>()
  for (const m of xml.matchAll(/<v:shapetype\b([^>]*)>/g)) {
    const a = parseTagAttrs(m[1]!)
    if (a['id']) out.set(a['id'], a)
  }
  return out
}

/** default geometry of the well-known o:spt ids a shape may carry without its shapetype */
const VML_SPT_PATHS: Record<string, string> = {
  '1': 'm,l,21600r21600,l21600,xe',
  '4': 'm10800,l,10800,10800,21600,21600,10800xe',
  '5': 'm10800,l,21600r21600,xe',
  '6': 'm,l,21600r21600,xe',
  '110': 'm10800,l,10800,10800,21600,21600,10800xe',
  '202': 'm,l,21600r21600,l21600,xe',
}

function parseTagAttrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const m of tag.matchAll(/([\w:.-]+)="([^"]*)"/g)) out[m[1]!] = m[2]!
  return out
}

/**
 * Textless VML shape (v:oval / v:rect / v:roundrect / straight-path v:shape)
 * with a solid or linear-gradient fill and stroke → SVG data URL at the
 * shape's declared size. Attributes and geometry the shape omits come from
 * the v:shapetype it points at (in the pict or in `shapeTypes`). Bails (null)
 * on groups, pictures, text and curved paths so partial art never renders.
 */
export function vmlShapeSvg(
  pict: string,
  shapeTypes?: Pick<ReadonlyMap<string, Record<string, string>>, 'get'>,
): { dataUrl: string; widthPx: number; heightPx: number; style: string } | null {
  if (/<v:(group|imagedata|textpath|textbox)\b|<w:txbxContent/.test(pict)) return null
  const m = vmlPictShape(pict)
  if (!m) return null
  const [, rawKind, attrText] = m
  const own = parseTagAttrs(attrText!)
  const typeId = /^#(.+)$/.exec(own['type'] ?? '')?.[1]
  const typeAttrs =
    (typeId && (vmlShapeTypeTable(pict).get(typeId) ?? shapeTypes?.get(typeId))) || {}
  const a: Record<string, string> = { ...typeAttrs, ...own }
  const style = own['style'] ?? ''
  if (/visibility:\s*hidden/.test(style)) return null
  const w = vmlStyleDimPx(style, 'width')
  const h = vmlStyleDimPx(style, 'height')
  if (!w || !h) return null
  const children = m[4] ?? ''
  const fillNode = parseTagAttrs(/<v:fill\b([^>]*)>/.exec(children)?.[1] ?? '')
  const strokeNode = parseTagAttrs(/<v:stroke\b([^>]*)>/.exec(children)?.[1] ?? '')
  const off = (v: string | undefined) => v === 'f' || v === 'false' || v === '0'
  const filled = !off(a['filled']) && !off(fillNode['on'])
  const color1 = vmlColorHex(a['fillcolor']) ?? 'FFFFFF'
  const color2 = vmlColorHex(fillNode['color2'])
  const fillType = fillNode['type']
  if (fillType && fillType !== 'solid' && fillType !== 'gradient') return null
  const gradient = filled && fillType === 'gradient' && color2 ? vmlGradient(fillNode) : null
  let defs = ''
  let fill = 'none'
  if (filled && gradient) {
    const stops = gradient.stops
      .map(([o, c]) => `<stop offset="${o}" stop-color="#${c === 1 ? color1 : color2}"/>`)
      .join('')
    defs = `<defs><linearGradient id="g" x1="${gradient.x1}" y1="${gradient.y1}" x2="${gradient.x2}" y2="${gradient.y2}">${stops}</linearGradient></defs>`
    fill = 'url(#g)'
  } else if (filled) fill = `#${color1}`
  const opacity = vmlFraction(fillNode['opacity'])
  const stroked = !off(a['stroked']) && !off(strokeNode['on'])
  const sw = stroked ? (vmlLengthPx(a['strokeweight']) ?? 1) : 0
  const paint =
    ` fill="${fill}"` +
    (opacity != null && opacity < 1 ? ` fill-opacity="${opacity}"` : '') +
    (stroked ? ` stroke="#${vmlColorHex(a['strokecolor']) ?? '000000'}" stroke-width="${sw}"` : '')
  // o:spt 2 / 3 shapes drawn as v:shape keep their preset geometry (the
  // shapetype's own path is a curve the path converter cannot take)
  const spt = own['path'] ? undefined : a['o:spt']
  const kind =
    rawKind === 'shape' && spt === '3'
      ? 'oval'
      : rawKind === 'shape' && spt === '2'
        ? 'roundrect'
        : rawKind
  // the stroke is centered on the geometry edge: inset it so the img box keeps all of it
  const ix = sw / 2
  const iw = Math.max(0, w - sw)
  const ih = Math.max(0, h - sw)
  let body: string
  if (kind === 'oval') {
    body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${iw / 2}" ry="${ih / 2}"${paint}/>`
  } else if (kind === 'rect' || kind === 'roundrect') {
    // arcsize: corner radius as a fraction of half the shorter side (default 0.2)
    const arc = kind === 'roundrect' ? (vmlFraction(a['arcsize']) ?? 0.2) : 0
    const r = Math.round(arc * (Math.min(iw, ih) / 2) * 100) / 100
    body = `<rect x="${ix}" y="${ix}" width="${iw}" height="${ih}"${r > 0 ? ` rx="${r}"` : ''}${paint}/>`
  } else {
    const path = a['path'] ?? VML_SPT_PATHS[a['o:spt'] ?? '']
    const cs = /^\s*(-?\d+)[,\s]+(-?\d+)/.exec(a['coordsize'] ?? '21600,21600')
    const d =
      path && cs ? vmlPathToNormD(path, parseInt(cs[1]!, 10), parseInt(cs[2]!, 10)) : undefined
    if (!d) return null
    let axis = 0
    const placed = d
      .split(' ')
      .map((tok) => {
        const n = Number(tok)
        if (!Number.isFinite(n)) {
          axis = 0
          return tok
        }
        const v = axis++ % 2 === 0 ? ix + n * iw : ix + n * ih
        return String(Math.round(v * 100) / 100)
      })
      .join(' ')
    body = `<path d="${placed}"${paint}/>`
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}">${defs}${body}</svg>`
  return {
    dataUrl: `data:image/svg+xml,${encodeURIComponent(svg)}`,
    widthPx: w,
    heightPx: h,
    style,
  }
}

/**
 * v:fill type="gradient" → SVG linearGradient line (bounding-box fractions)
 * and stops (1 = fillcolor, 2 = color2). VML angles count counterclockwise
 * from the bottom edge, so angle 0 runs top→bottom with color2 at the top
 * and fillcolor at the bottom. A focus of ±50% is axial (color2 or fillcolor
 * in the middle) and, like LibreOffice's import, anything in 25%..75% is
 * approximated the same way; beyond that a linear fill with swapped ends.
 */
function vmlGradient(fillNode: Record<string, string>): {
  x1: number
  y1: number
  x2: number
  y2: number
  stops: Array<[number, 1 | 2]>
} {
  const angle = parseFloat(fillNode['angle'] ?? '0') || 0
  const focus = vmlFraction(fillNode['focus']?.replace('%', '')) ?? 0
  const f = Math.abs(focus) > 1 ? focus / 100 : focus
  const rad = ((90 - angle) * Math.PI) / 180
  const dx = Math.cos(rad) / 2
  const dy = Math.sin(rad) / 2
  const r = (v: number) => Math.round(v * 1000) / 1000
  const line = { x1: r(0.5 - dx), y1: r(0.5 - dy), x2: r(0.5 + dx), y2: r(0.5 + dy) }
  const axial = Math.abs(f) >= 0.25 && Math.abs(f) <= 0.75
  if (axial) {
    const stops: Array<[number, 1 | 2]> =
      f > 0
        ? [
            [0, 1],
            [0.5, 2],
            [1, 1],
          ]
        : [
            [0, 2],
            [0.5, 1],
            [1, 2],
          ]
    return { ...line, stops }
  }
  let swap = angle < 0
  if (Math.abs(f) > 0.5) swap = !swap
  const stops: Array<[number, 1 | 2]> = swap
    ? [
        [0, 1],
        [1, 2],
      ]
    : [
        [0, 2],
        [1, 1],
      ]
  return { ...line, stops }
}

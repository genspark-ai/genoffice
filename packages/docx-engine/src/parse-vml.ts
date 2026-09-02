// Legacy VML (v:shape / v:group) geometry and color readers.
import { attrsOf, findChild, type XNode } from './xml-utils'
import { EMU_PER_PX } from './parse-xml-text'
import type { Run, TextboxDisplay } from './types'

/** VML style="width:189.9pt;height:626pt" dimension → CSS px */
function vmlStyleDimPx(style: string, key: 'width' | 'height'): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${key}:([0-9.]+)(pt|px|in|mm|cm)?`).exec(style)
  if (!m) return undefined
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v) || v <= 0) return undefined
  const unit = m[2] ?? 'pt'
  const px =
    unit === 'px'
      ? v
      : unit === 'in'
        ? v * 96
        : unit === 'mm'
          ? (v / 25.4) * 96
          : unit === 'cm'
            ? (v / 2.54) * 96
            : (v * 96) / 72
  return Math.round(px)
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

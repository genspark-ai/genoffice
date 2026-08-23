/**
 * Color node resolution (srgbClr/schemeClr/sysClr/prstClr + lumMod/lumOff/tint/shade/alpha modifiers).
 * Extracted from parse.ts into a shared module: used by both parse (run/fill colors) and
 * placeholder (lstStyle defRPr default colors) to avoid a circular dependency.
 */
import { type Theme, resolveSchemeColor } from './theme'
import { asXmlNode, type XmlNode } from './xml-utils'

/**
 * Resolve a color node + modifiers: srgbClr/schemeClr/sysClr, supporting alpha
 * (→#RRGGBBAA) and lumMod/lumOff (tint/shade, common for theme colors).
 */
export function resolveColorNode(
  node: unknown,
  theme: Theme | undefined,
  phClr?: string,
): string | undefined {
  if (!node) return undefined
  const n = asXmlNode(node)
  let base: string | undefined
  let mods: XmlNode | undefined
  if (n['a:srgbClr']) {
    mods = asXmlNode(n['a:srgbClr'])
    base = '#' + String(mods['@_val']).toUpperCase()
  } else if (n['a:schemeClr']) {
    mods = asXmlNode(n['a:schemeClr'])
    base = resolveSchemeColor(String(mods['@_val']), theme, phClr)
  } else if (n['a:sysClr']) {
    mods = asXmlNode(n['a:sysClr'])
    base = '#' + String(mods['@_lastClr'] ?? '000000').toUpperCase()
  } else if (n['a:prstClr']) {
    mods = asXmlNode(n['a:prstClr'])
    base = PRESET_COLORS[String(mods['@_val'])]
  }
  if (!base) return undefined
  return applyColorMods(base, mods)
}

/** ECMA-376 ST_PresetColorVal (CSS/X11 values plus the dk/lt/med aliases). */
const PRESET_COLORS: Record<string, string> = {
  aliceBlue: '#F0F8FF',
  antiqueWhite: '#FAEBD7',
  aqua: '#00FFFF',
  aquamarine: '#7FFFD4',
  azure: '#F0FFFF',
  beige: '#F5F5DC',
  bisque: '#FFE4C4',
  black: '#000000',
  blanchedAlmond: '#FFEBCD',
  blue: '#0000FF',
  blueViolet: '#8A2BE2',
  brown: '#A52A2A',
  burlyWood: '#DEB887',
  cadetBlue: '#5F9EA0',
  chartreuse: '#7FFF00',
  chocolate: '#D2691E',
  coral: '#FF7F50',
  cornflowerBlue: '#6495ED',
  cornsilk: '#FFF8DC',
  crimson: '#DC143C',
  cyan: '#00FFFF',
  darkBlue: '#00008B',
  darkCyan: '#008B8B',
  darkGoldenrod: '#B8860B',
  darkGray: '#A9A9A9',
  darkGreen: '#006400',
  darkGrey: '#A9A9A9',
  darkKhaki: '#BDB76B',
  darkMagenta: '#8B008B',
  darkOliveGreen: '#556B2F',
  darkOrange: '#FF8C00',
  darkOrchid: '#9932CC',
  darkRed: '#8B0000',
  darkSalmon: '#E9967A',
  darkSeaGreen: '#8FBC8F',
  darkSlateBlue: '#483D8B',
  darkSlateGray: '#2F4F4F',
  darkSlateGrey: '#2F4F4F',
  darkTurquoise: '#00CED1',
  darkViolet: '#9400D3',
  deepPink: '#FF1493',
  deepSkyBlue: '#00BFFF',
  dimGray: '#696969',
  dimGrey: '#696969',
  dkBlue: '#00008B',
  dkCyan: '#008B8B',
  dkGoldenrod: '#B8860B',
  dkGray: '#A9A9A9',
  dkGreen: '#006400',
  dkGrey: '#A9A9A9',
  dkKhaki: '#BDB76B',
  dkMagenta: '#8B008B',
  dkOliveGreen: '#556B2F',
  dkOrange: '#FF8C00',
  dkOrchid: '#9932CC',
  dkRed: '#8B0000',
  dkSalmon: '#E9967A',
  dkSeaGreen: '#8FBC8F',
  dkSlateBlue: '#483D8B',
  dkSlateGray: '#2F4F4F',
  dkSlateGrey: '#2F4F4F',
  dkTurquoise: '#00CED1',
  dkViolet: '#9400D3',
  dodgerBlue: '#1E90FF',
  firebrick: '#B22222',
  floralWhite: '#FFFAF0',
  forestGreen: '#228B22',
  fuchsia: '#FF00FF',
  gainsboro: '#DCDCDC',
  ghostWhite: '#F8F8FF',
  gold: '#FFD700',
  goldenrod: '#DAA520',
  gray: '#808080',
  green: '#008000',
  greenYellow: '#ADFF2F',
  grey: '#808080',
  honeydew: '#F0FFF0',
  hotPink: '#FF69B4',
  indianRed: '#CD5C5C',
  indigo: '#4B0082',
  ivory: '#FFFFF0',
  khaki: '#F0E68C',
  lavender: '#E6E6FA',
  lavenderBlush: '#FFF0F5',
  lawnGreen: '#7CFC00',
  lemonChiffon: '#FFFACD',
  lightBlue: '#ADD8E6',
  lightCoral: '#F08080',
  lightCyan: '#E0FFFF',
  lightGoldenrodYellow: '#FAFAD2',
  lightGray: '#D3D3D3',
  lightGreen: '#90EE90',
  lightGrey: '#D3D3D3',
  lightPink: '#FFB6C1',
  lightSalmon: '#FFA07A',
  lightSeaGreen: '#20B2AA',
  lightSkyBlue: '#87CEFA',
  lightSlateGray: '#778899',
  lightSlateGrey: '#778899',
  lightSteelBlue: '#B0C4DE',
  lightYellow: '#FFFFE0',
  lime: '#00FF00',
  limeGreen: '#32CD32',
  linen: '#FAF0E6',
  ltBlue: '#ADD8E6',
  ltCoral: '#F08080',
  ltCyan: '#E0FFFF',
  ltGoldenrodYellow: '#FAFAD2',
  ltGray: '#D3D3D3',
  ltGreen: '#90EE90',
  ltGrey: '#D3D3D3',
  ltPink: '#FFB6C1',
  ltSalmon: '#FFA07A',
  ltSeaGreen: '#20B2AA',
  ltSkyBlue: '#87CEFA',
  ltSlateGray: '#778899',
  ltSlateGrey: '#778899',
  ltSteelBlue: '#B0C4DE',
  ltYellow: '#FFFFE0',
  magenta: '#FF00FF',
  maroon: '#800000',
  medAquamarine: '#66CDAA',
  medBlue: '#0000CD',
  mediumAquamarine: '#66CDAA',
  mediumBlue: '#0000CD',
  mediumOrchid: '#BA55D3',
  mediumPurple: '#9370DB',
  mediumSeaGreen: '#3CB371',
  mediumSlateBlue: '#7B68EE',
  mediumSpringGreen: '#00FA9A',
  mediumTurquoise: '#48D1CC',
  mediumVioletRed: '#C71585',
  medOrchid: '#BA55D3',
  medPurple: '#9370DB',
  medSeaGreen: '#3CB371',
  medSlateBlue: '#7B68EE',
  medSpringGreen: '#00FA9A',
  medTurquoise: '#48D1CC',
  medVioletRed: '#C71585',
  midnightBlue: '#191970',
  mintCream: '#F5FFFA',
  mistyRose: '#FFE4E1',
  moccasin: '#FFE4B5',
  navajoWhite: '#FFDEAD',
  navy: '#000080',
  oldLace: '#FDF5E6',
  olive: '#808000',
  oliveDrab: '#6B8E23',
  orange: '#FFA500',
  orangeRed: '#FF4500',
  orchid: '#DA70D6',
  paleGoldenrod: '#EEE8AA',
  paleGreen: '#98FB98',
  paleTurquoise: '#AFEEEE',
  paleVioletRed: '#DB7093',
  papayaWhip: '#FFEFD5',
  peachPuff: '#FFDAB9',
  peru: '#CD853F',
  pink: '#FFC0CB',
  plum: '#DDA0DD',
  powderBlue: '#B0E0E6',
  purple: '#800080',
  red: '#FF0000',
  rosyBrown: '#BC8F8F',
  royalBlue: '#4169E1',
  saddleBrown: '#8B4513',
  salmon: '#FA8072',
  sandyBrown: '#F4A460',
  seaGreen: '#2E8B57',
  seaShell: '#FFF5EE',
  sienna: '#A0522D',
  silver: '#C0C0C0',
  skyBlue: '#87CEEB',
  slateBlue: '#6A5ACD',
  slateGray: '#708090',
  slateGrey: '#708090',
  snow: '#FFFAFA',
  springGreen: '#00FF7F',
  steelBlue: '#4682B4',
  tan: '#D2B48C',
  teal: '#008080',
  thistle: '#D8BFD8',
  tomato: '#FF6347',
  turquoise: '#40E0D0',
  violet: '#EE82EE',
  wheat: '#F5DEB3',
  white: '#FFFFFF',
  whiteSmoke: '#F5F5F5',
  yellow: '#FFFF00',
  yellowGreen: '#9ACD32',
}

/** Apply lumMod/lumOff/tint/shade/satMod/alpha modifiers (percentages, in units of 1/1000%). */
export function applyColorMods(hex: string, mods: XmlNode | undefined): string {
  let { r, g, b } = hexToRgb(hex)
  const pct = (k: string): number | undefined => {
    const v = asXmlNode(mods?.[k])['@_val']
    return v != null ? (parseInt(String(v), 10) || 0) / 100000 : undefined
  }
  const lumMod = pct('a:lumMod')
  const lumOff = pct('a:lumOff')
  const shade = pct('a:shade')
  const tint = pct('a:tint')
  // satMod (frequent in theme gradient templates): HSL saturation multiplier
  const satMod = pct('a:satMod')
  if (satMod != null) {
    const { h, s, l } = rgbToHsl(r, g, b)
    const rgb2 = hslToRgb(h, Math.max(0, Math.min(1, s * satMod)), l)
    r = rgb2.r
    g = rgb2.g
    b = rgb2.b
  }
  if (lumMod != null) {
    r *= lumMod
    g *= lumMod
    b *= lumMod
  }
  if (lumOff != null) {
    r += 255 * lumOff
    g += 255 * lumOff
    b += 255 * lumOff
  }
  // tint/shade blend in linear-gamma space (PowerPoint semantics: black + 55% tint
  // renders ≈#B1, not #73; verified against PowerPoint for Mac gradient stops)
  const lin = (v: number) => Math.pow(Math.max(v, 0) / 255, 2.2)
  const unlin = (v: number) => 255 * Math.pow(Math.max(v, 0), 1 / 2.2)
  if (shade != null) {
    r = unlin(lin(r) * shade)
    g = unlin(lin(g) * shade)
    b = unlin(lin(b) * shade)
  }
  if (tint != null) {
    r = unlin(lin(r) * tint + (1 - tint))
    g = unlin(lin(g) * tint + (1 - tint))
    b = unlin(lin(b) * tint + (1 - tint))
  }
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)))
  let out =
    '#' +
    [clamp(r), clamp(g), clamp(b)]
      .map((v) => v.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  const alpha = pct('a:alpha')
  if (alpha != null && alpha < 1) {
    out += Math.round(alpha * 255)
      .toString(16)
      .padStart(2, '0')
      .toUpperCase()
  }
  return out
}

export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const h = hex.replace(/^#/, '')
  return {
    r: parseInt(h.slice(0, 2), 16) || 0,
    g: parseInt(h.slice(2, 4), 16) || 0,
    b: parseInt(h.slice(4, 6), 16) || 0,
  }
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  r /= 255
  g /= 255
  b /= 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return { h: 0, s: 0, l }
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6
  else if (max === g) h = ((b - r) / d + 2) / 6
  else h = ((r - g) / d + 4) / 6
  return { h, s, l }
}

/** Scale HSL luminance by a factor (legacy chart-style monochrome ramps). */
export function scaleLuminance(hex: string, factor: number): string {
  const { r, g, b } = hexToRgb(hex)
  const { h, s, l } = rgbToHsl(r, g, b)
  const rgb = hslToRgb(h, s, Math.max(0, Math.min(0.97, l * factor)))
  const to2 = (v: number) =>
    Math.max(0, Math.min(255, Math.round(v)))
      .toString(16)
      .padStart(2, '0')
  return ('#' + to2(rgb.r) + to2(rgb.g) + to2(rgb.b)).toUpperCase()
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  if (s === 0) {
    const v = l * 255
    return { r: v, g: v, b: v }
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const f = (t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  return { r: f(h + 1 / 3) * 255, g: f(h) * 255, b: f(h - 1 / 3) * 255 }
}

/// Cell colors as styles.xml stores them: a literal #RRGGBB, or a theme slot
/// with an optional tint that Excel re-resolves when the document theme
/// changes. The tint math is Excel's HLS algorithm, so resolved values match
/// what Excel shows.

export interface ThemeColor {
  /// theme attribute index: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 accent1-6, 10 hlink, 11 folHlink
  theme: number
  /// -1 (black) … 1 (white); 0 or absent = the slot color itself
  tint?: number | undefined
}

export type StyleColor = string | ThemeColor

/// Slot names by theme index (Excel swaps the dk/lt pairs relative to clrScheme order).
export const THEME_SLOT_NAMES = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const

export type ThemeSlotName = (typeof THEME_SLOT_NAMES)[number]

const SLOT_ALIASES: Record<string, number> = {
  bg1: 0,
  background1: 0,
  tx1: 1,
  text1: 1,
  bg2: 2,
  background2: 2,
  tx2: 3,
  text2: 3,
  hyperlink: 10,
  followedhyperlink: 11,
}

/// Office theme (Excel 2013+), theme index order.
export const DEFAULT_THEME_PALETTE: readonly string[] = [
  '#FFFFFF',
  '#000000',
  '#E7E6E6',
  '#44546A',
  '#4472C4',
  '#ED7D31',
  '#A5A5A5',
  '#FFC000',
  '#5B9BD5',
  '#70AD47',
  '#0563C1',
  '#954F72',
]

export const THEME_SHORTHAND_PATTERN =
  /^(lt1|dk1|lt2|dk2|accent[1-6]|hlink|folHlink|bg1|bg2|tx1|tx2|background[12]|text[12]|hyperlink|followedHyperlink)(?:\s*([+-])\s*(\d{1,3})%)?$/i

export function themeSlotIndex(name: string): number | undefined {
  const lower = name.toLowerCase()
  const exact = THEME_SLOT_NAMES.findIndex((slot) => slot.toLowerCase() === lower)
  if (exact !== -1) return exact
  return SLOT_ALIASES[lower]
}

export function themeSlotName(index: number): ThemeSlotName {
  return THEME_SLOT_NAMES[index] ?? 'accent1'
}

export function isThemeColor(color: StyleColor | null | undefined): color is ThemeColor {
  return typeof color === 'object' && color !== null && typeof color.theme === 'number'
}

/// Accepts `#RRGGBB`, `accent1`, `accent1+40%`, `dk2-25%` or `{theme: 4 | 'accent1', tint?}`.
export function normalizeStyleColor(
  input: string | { theme: number | string; tint?: number | undefined },
): StyleColor {
  if (typeof input === 'string') {
    if (/^#[0-9A-Fa-f]{6}$/.test(input)) return `#${input.slice(1).toUpperCase()}`
    const match = THEME_SHORTHAND_PATTERN.exec(input.trim())
    if (!match) throw new Error(`Unknown color "${input}"`)
    const theme = themeSlotIndex(match[1]!)!
    if (match[2] === undefined) return { theme }
    const tint = (Number(match[3]) / 100) * (match[2] === '-' ? -1 : 1)
    return withTint(theme, tint)
  }
  const theme = typeof input.theme === 'number' ? input.theme : themeSlotIndex(input.theme)
  if (theme === undefined || !Number.isInteger(theme) || theme < 0 || theme > 11) {
    throw new Error(`Unknown theme color "${String(input.theme)}"`)
  }
  return withTint(theme, input.tint)
}

function withTint(theme: number, tint: number | undefined): ThemeColor {
  if (tint === undefined || tint === 0) return { theme }
  if (!(tint >= -1 && tint <= 1)) throw new Error(`Tint ${tint} is outside -1..1`)
  return { theme, tint: Math.round(tint * 1e6) / 1e6 }
}

/// #RRGGBB the color shows as under `palette` (theme index order).
export function resolveStyleColor(
  color: StyleColor,
  palette: readonly string[] = DEFAULT_THEME_PALETTE,
): string {
  if (typeof color === 'string') return isValidHexColor(color) ? color.toUpperCase() : '#000000'
  const base = palette[color.theme] ?? DEFAULT_THEME_PALETTE[color.theme] ?? '#000000'
  return applyTint(base, color.tint ?? 0)
}

/// ECMA-376 tint: lighten towards white for positive, darken towards black for
/// negative, on the HLS luminance channel. Channels truncate on the way back,
/// which lands on Office's palette swatches within one step per channel.
/// Malformed hex falls back to black, matching the missing-palette fallback
/// in resolveStyleColor, so NaN channels never leak into outputs.
export function applyTint(hex: string, tint: number): string {
  if (!isValidHexColor(hex)) return '#000000'
  if (tint === 0) return hex.toUpperCase()
  const [h, l, s] = rgbToHls(hex)
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  return hlsToRgb(h, Math.min(1, Math.max(0, lum)), s)
}

function isValidHexColor(hex: string): boolean {
  return /^#[0-9A-Fa-f]{6}$/.test(hex)
}

function rgbToHls(hex: string): [number, number, number] {
  if (!isValidHexColor(hex)) return [0, 0, 0]
  const r = parseInt(hex.slice(1, 3), 16) / 255
  const g = parseInt(hex.slice(3, 5), 16) / 255
  const b = parseInt(hex.slice(5, 7), 16) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, l, 0]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  return [h / 6, l, s]
}

function hlsToRgb(h: number, l: number, s: number): string {
  const channel = (t: number): number => {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    let x = t
    if (x < 0) x += 1
    if (x > 1) x -= 1
    if (x < 1 / 6) return p + (q - p) * 6 * x
    if (x < 1 / 2) return q
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6
    return p
  }
  const toHex = (v: number): string =>
    Math.min(255, Math.floor(v * 255 + 1e-9))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase()
  if (s === 0) {
    const grey = toHex(l)
    return `#${grey}${grey}${grey}`
  }
  return `#${toHex(channel(h + 1 / 3))}${toHex(channel(h))}${toHex(channel(h - 1 / 3))}`
}

/// Agent-facing echo of a color: the resolved rgb plus the theme slot it came from.
export interface StyleColorEcho {
  rgb: string
  theme?: ThemeSlotName | undefined
  tint?: number | undefined
}

export function echoStyleColor(color: StyleColor, palette?: readonly string[]): StyleColorEcho {
  const rgb = resolveStyleColor(color, palette)
  if (typeof color === 'string') return { rgb }
  return {
    rgb,
    theme: themeSlotName(color.theme),
    ...(color.tint ? { tint: color.tint } : {}),
  }
}

export function describeStyleColor(color: StyleColor): string {
  if (typeof color === 'string') return color
  const tint = color.tint ?? 0
  const name = themeSlotName(color.theme)
  if (tint === 0) return name
  return `${name}${tint > 0 ? '+' : '-'}${Math.round(Math.abs(tint) * 100)}%`
}

// ── fills ──────────────────────────────────────────────────────────────

/// ST_PatternType minus `none` (a cleared fill is `fill: null`).
export const PATTERN_TYPES = [
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'gray125',
  'gray0625',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
] as const

export type PatternType = (typeof PATTERN_TYPES)[number]

export interface PatternFill {
  pattern: PatternType
  /// pattern foreground; for `solid` the cell color
  fg: StyleColor
  bg?: StyleColor | undefined
}

export interface GradientStop {
  /// 0 … 1 along the gradient
  position: number
  color: StyleColor
}

export interface GradientFill {
  gradient: {
    type?: 'linear' | 'path' | undefined
    /// linear: clockwise degrees, 0 = left→right, 90 = top→bottom
    angle?: number | undefined
    /// path: inner rectangle edges as fractions of the cell
    left?: number | undefined
    right?: number | undefined
    top?: number | undefined
    bottom?: number | undefined
    stops: GradientStop[]
  }
}

export type FillSpec = PatternFill | GradientFill

export function isGradientFill(fill: FillSpec): fill is GradientFill {
  return 'gradient' in fill
}

/// The single color a renderer without pattern/gradient support should paint.
export function fillDisplayColor(fill: FillSpec): StyleColor | undefined {
  if (isGradientFill(fill)) return fill.gradient.stops[0]?.color
  return fill.fg
}

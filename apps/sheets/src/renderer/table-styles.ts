import type { WorkbookFile } from '../shared/desktop-api'

/// Excel's built-in table styles, resolved against the workbook theme the
/// same way the file reader does (accent cycle (n-1) % 7 with 0 = dk1,
/// variants in blocks of 7; dk1 members tint 0.05 lighter).

type FileTable = WorkbookFile['sheets'][number]['tables'][number]

export type TableStylePalette = Partial<
  Pick<
    FileTable,
    | 'headerFill'
    | 'headerFontColor'
    | 'stripeFill'
    | 'secondRowStripeFill'
    | 'columnStripeFill'
    | 'wholeTableFill'
    | 'totalRowFill'
    | 'totalRowFontColor'
    | 'totalRowBorderColor'
    | 'totalRowBorderStyle'
    | 'bodyFontColor'
    | 'borderColor'
    | 'wholeTableBorderColor'
    | 'wholeTableBorderStyle'
    | 'innerHorizontalBorderColor'
    | 'innerHorizontalBorderStyle'
    | 'innerVerticalBorderColor'
    | 'innerVerticalBorderStyle'
  >
>

export const TABLE_PALETTE_KEYS = [
  'headerFill',
  'headerFontColor',
  'stripeFill',
  'secondRowStripeFill',
  'columnStripeFill',
  'secondColumnStripeFill',
  'wholeTableFill',
  'firstColumnFill',
  'lastColumnFill',
  'totalRowFill',
  'totalRowFontColor',
  'totalRowBorderColor',
  'totalRowBorderStyle',
  'bodyFontColor',
  'firstHeaderCellFontColor',
  'borderColor',
  'wholeTableBorderColor',
  'wholeTableBorderStyle',
  'innerHorizontalBorderColor',
  'innerHorizontalBorderStyle',
  'innerVerticalBorderColor',
  'innerVerticalBorderStyle',
  'headerBottomBorderColor',
  'headerBottomBorderStyle',
] as const satisfies readonly (keyof FileTable)[]

export type TableStyleFamily = 'Light' | 'Medium' | 'Dark'
export const TABLE_STYLE_COUNTS: Record<TableStyleFamily, number> = {
  Light: 21,
  Medium: 28,
  Dark: 11,
}
export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2'

export function tableStyleNames(family: TableStyleFamily): string[] {
  return Array.from(
    { length: TABLE_STYLE_COUNTS[family] },
    (_, index) => `TableStyle${family}${index + 1}`,
  )
}

export function isBuiltinTableStyle(name: string): boolean {
  const match = /^TableStyle(Light|Medium|Dark)([1-9][0-9]?)$/.exec(name)
  return match !== null && Number(match[2]) <= TABLE_STYLE_COUNTS[match[1] as TableStyleFamily]
}

const DEFAULT_ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
const WHITE = '#FFFFFF'

type Rgb = readonly [number, number, number]

function parseHex(hex: string): Rgb | null {
  const match = /^#?([0-9a-f]{6})$/i.exec(hex.trim())
  if (!match) return null
  const value = Number.parseInt(match[1]!, 16)
  return [(value >> 16) & 0xff, (value >> 8) & 0xff, value & 0xff]
}

function toHex([red, green, blue]: Rgb): string {
  return `#${[red, green, blue].map((part) => part.toString(16).padStart(2, '0').toUpperCase()).join('')}`
}

function rgbToHsl([red, green, blue]: Rgb): [number, number, number] {
  const r = red / 255
  const g = green / 255
  const b = blue / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToRgb(h: number, s: number, l: number): Rgb {
  if (s === 0) {
    const gray = Math.round(l * 255)
    return [gray, gray, gray]
  }
  const hue2rgb = (p: number, q: number, t: number): number => {
    let tt = t
    if (tt < 0) tt += 1
    if (tt > 1) tt -= 1
    if (tt < 1 / 6) return p + (q - p) * 6 * tt
    if (tt < 1 / 2) return q
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6
    return p
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [
    Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, h) * 255),
    Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  ]
}

/// Excel's tint: luminance toward white (tint > 0) or black (tint < 0).
export function applyTint(hex: string, tint: number): string {
  const rgb = parseHex(hex)
  if (!rgb) return hex
  if (tint === 0) return toHex(rgb)
  const [h, s, l] = rgbToHsl(rgb)
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  return toHex(hslToRgb(h, s, Math.min(1, Math.max(0, lum))))
}

interface ThemeBase {
  readonly dark1: string
  readonly accent: (index: number) => string
}

function themeBase(themeColors: readonly string[] | undefined): ThemeBase {
  const slot = (index: number): string | undefined => {
    const value = themeColors?.[index]
    return value && parseHex(value) ? value : undefined
  }
  return {
    dark1: slot(1) ?? '#000000',
    accent: (index) => {
      const wrapped = ((index - 1) % 6) + 1
      return slot(3 + wrapped) ?? DEFAULT_ACCENTS[wrapped - 1]!
    },
  }
}

export function builtinTablePalette(
  styleName: string | undefined,
  themeColors: readonly string[] | undefined,
): TableStylePalette {
  if (!styleName) return {}
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(styleName)
  const family = (match?.[1] ?? 'Medium') as TableStyleFamily
  const number = Number(match?.[2] ?? (family === 'Medium' ? 2 : 1))
  const theme = themeBase(themeColors)
  const accentIndex = Math.max(0, number - 1) % 7
  const neutral = accentIndex === 0
  const base = neutral ? theme.dark1 : theme.accent(accentIndex)
  const band = (color: string, fromDark1: boolean, tint: number): string =>
    applyTint(color, fromDark1 ? tint + 0.05 : tint)
  const variant = Math.floor(Math.max(0, number - 1) / 7)
  const palette: TableStylePalette = {}
  if (family === 'Light') {
    if (variant === 0) {
      palette.headerFontColor = applyTint(base, -0.25)
      palette.stripeFill = band(base, neutral, 0.8)
      palette.totalRowFontColor = applyTint(base, -0.25)
      palette.totalRowBorderColor = base
      palette.totalRowBorderStyle = 'thin'
      palette.borderColor = base
    } else if (variant === 1) {
      palette.headerFill = base
      palette.headerFontColor = WHITE
      palette.totalRowBorderColor = base
      palette.totalRowBorderStyle = 'double'
    } else {
      palette.headerFontColor = theme.dark1
      palette.stripeFill = band(base, neutral, 0.8)
      palette.totalRowBorderColor = base
      palette.totalRowBorderStyle = 'double'
      palette.wholeTableBorderColor = base
      palette.wholeTableBorderStyle = 'thin'
      palette.innerHorizontalBorderColor = base
      palette.innerHorizontalBorderStyle = 'thin'
      palette.innerVerticalBorderColor = base
      palette.innerVerticalBorderStyle = 'thin'
    }
  } else if (family === 'Medium') {
    if (variant === 1) {
      palette.headerFill = base
      palette.headerFontColor = WHITE
      palette.stripeFill = band(base, neutral, 0.6)
      palette.secondRowStripeFill = band(base, neutral, 0.8)
      palette.wholeTableFill = band(base, neutral, 0.8)
      palette.totalRowFill = base
      palette.totalRowFontColor = WHITE
    } else if (variant === 2) {
      palette.headerFill = base
      palette.headerFontColor = WHITE
      palette.stripeFill = applyTint(theme.dark1, 0.85)
      palette.totalRowBorderColor = theme.dark1
      palette.totalRowBorderStyle = 'double'
    } else if (variant === 3) {
      palette.headerFill = band(base, neutral, 0.8)
      palette.headerFontColor = theme.dark1
      palette.stripeFill = band(base, neutral, 0.6)
      palette.wholeTableFill = band(base, neutral, 0.8)
      palette.totalRowFill = band(base, neutral, 0.8)
      palette.totalRowBorderColor = base
      palette.totalRowBorderStyle = 'medium'
    } else {
      palette.headerFill = base
      palette.headerFontColor = WHITE
      palette.stripeFill = band(base, neutral, 0.8)
      palette.totalRowBorderColor = base
      palette.totalRowBorderStyle = 'double'
    }
  } else if (variant === 1) {
    const bandBase = neutral ? theme.dark1 : theme.accent(2 * (number - 8) - 1)
    const headerBase = neutral ? theme.dark1 : theme.accent(2 * (number - 8))
    palette.headerFill = headerBase
    palette.headerFontColor = WHITE
    palette.stripeFill = band(bandBase, neutral, 0.6)
    palette.secondRowStripeFill = band(bandBase, neutral, 0.8)
    palette.wholeTableFill = band(bandBase, neutral, 0.8)
    palette.totalRowFill = band(bandBase, neutral, 0.8)
    palette.totalRowBorderColor = theme.dark1
    palette.totalRowBorderStyle = 'double'
  } else {
    palette.headerFill = theme.dark1
    palette.headerFontColor = WHITE
    if (neutral) {
      palette.wholeTableFill = applyTint(theme.dark1, 0.45)
      palette.stripeFill = applyTint(theme.dark1, 0.25)
      palette.totalRowFill = applyTint(theme.dark1, 0.15)
    } else {
      palette.wholeTableFill = base
      palette.stripeFill = applyTint(base, -0.25)
      palette.totalRowFill = applyTint(base, -0.5)
    }
    palette.bodyFontColor = WHITE
    palette.totalRowFontColor = WHITE
  }
  if (palette.stripeFill) palette.columnStripeFill = palette.stripeFill
  return palette
}

export interface TableThemeOptions {
  readonly bandedRows: boolean
  readonly bandedColumns: boolean
  readonly firstColumn: boolean
  readonly lastColumn: boolean
}

interface ThemeItem {
  bg?: { rgb: string }
  cl?: { rgb: string }
  bl?: 0 | 1
  bd?: Record<string, { s: number; cl: { rgb: string } }>
}

export interface TableThemeJson {
  name: string
  wholeStyle?: ThemeItem
  headerRowStyle?: ThemeItem
  headerColumnStyle?: ThemeItem
  firstRowStyle?: ThemeItem
  secondRowStyle?: ThemeItem
  firstColumnStyle?: ThemeItem
  secondColumnStyle?: ThemeItem
  lastColumnStyle?: ThemeItem
}

// Univer BorderStyleTypes: THIN = 1, DOUBLE = 7, MEDIUM = 8.
const BORDER_STYLE: Record<string, number> = { thin: 1, medium: 8, double: 7 }

/// Univer range theme for a session table (its range excludes the totals
/// band, which is styled directly on its cells).
export function tableThemeJson(
  name: string,
  palette: TableStylePalette,
  options: TableThemeOptions,
): TableThemeJson {
  const fill = (rgb: string | undefined): ThemeItem => (rgb ? { bg: { rgb } } : {})
  const theme: TableThemeJson = { name }
  const whole: ThemeItem = fill(palette.wholeTableFill)
  if (palette.bodyFontColor) whole.cl = { rgb: palette.bodyFontColor }
  if (palette.innerHorizontalBorderColor || palette.innerVerticalBorderColor) {
    const edge = (color: string | undefined, style: string | undefined) =>
      color ? { s: BORDER_STYLE[style ?? 'thin'] ?? 1, cl: { rgb: color } } : undefined
    const h = edge(palette.innerHorizontalBorderColor, palette.innerHorizontalBorderStyle)
    const v = edge(palette.innerVerticalBorderColor, palette.innerVerticalBorderStyle)
    whole.bd = {
      ...(h ? { t: h, b: h } : {}),
      ...(v ? { l: v, r: v } : {}),
    }
  }
  if (Object.keys(whole).length > 0) theme.wholeStyle = whole
  const header: ThemeItem = { ...fill(palette.headerFill), bl: 1 }
  if (palette.headerFontColor) header.cl = { rgb: palette.headerFontColor }
  if (palette.borderColor) {
    header.bd = {
      t: { s: BORDER_STYLE['medium']!, cl: { rgb: palette.borderColor } },
      b: { s: BORDER_STYLE['thin']!, cl: { rgb: palette.borderColor } },
    }
  }
  theme.headerRowStyle = header
  if (options.bandedRows && palette.stripeFill) {
    theme.firstRowStyle = fill(palette.stripeFill)
    theme.secondRowStyle = fill(palette.secondRowStripeFill ?? palette.wholeTableFill)
  }
  if (options.bandedColumns && palette.columnStripeFill) {
    theme.firstColumnStyle = fill(palette.columnStripeFill)
    theme.secondColumnStyle = fill(palette.wholeTableFill)
  }
  if (options.firstColumn) theme.headerColumnStyle = { bl: 1 }
  if (options.lastColumn) theme.lastColumnStyle = { bl: 1 }
  return theme
}

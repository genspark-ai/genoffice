/**
 * Document theme engine (Page Layout > Themes / Colors / Fonts).
 *
 * Preset palettes are stored in theme index order [lt1, dk1, lt2, dk2,
 * accent1-6, hlink, folHlink] — the order the sidecar ships
 * `WorkbookFile.themeColors` in and the save's theme writer expects.
 * `applyTint` mirrors the sidecar's HSL-luminance tint transform, so styles
 * carrying theme/tint provenance re-resolve to the exact colors a reopen
 * would produce.
 */
import type { WorkbookCellStyle } from '../shared/desktop-api'

export interface ThemeColorScheme {
  readonly id: string
  readonly name: string
  /// #RRGGBB in theme index order, 12 entries.
  readonly values: readonly string[]
}

export interface ThemeFontScheme {
  readonly id: string
  readonly name: string
  readonly major: string
  readonly minor: string
}

export interface ThemePreset {
  readonly id: string
  readonly name: string
  readonly colors: ThemeColorScheme
  readonly fonts: ThemeFontScheme
}

/// dk1/lt1/dk2/lt2/accents/hlink (clrScheme semantics) → theme index order.
const scheme = (
  id: string,
  name: string,
  dk1: string,
  lt1: string,
  dk2: string,
  lt2: string,
  accents: readonly [string, string, string, string, string, string],
  hlink: string,
): ThemeColorScheme => ({
  id,
  name,
  values: [lt1, dk1, lt2, dk2, ...accents, hlink, '954F72'].map((value) => `#${value}`),
})

export const COLOR_SCHEMES: readonly ThemeColorScheme[] = [
  scheme(
    'office',
    'Office',
    '000000',
    'FFFFFF',
    '44546A',
    'E7E6E6',
    ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'],
    '0563C1',
  ),
  scheme(
    'ember',
    'Ember',
    '2B1B14',
    'FFFFFF',
    '632B1A',
    'F7EBE6',
    ['C43E1C', 'E97132', 'FFC000', '8A3B12', 'D98F73', 'A33517'],
    'C43E1C',
  ),
  scheme(
    'indigo',
    'Indigo',
    '1F2A44',
    'FFFFFF',
    '3B4C77',
    'E8ECF6',
    ['2E4FA3', '5B79C7', '8FA6DE', 'E97132', '31A5A0', '7030A0'],
    '2E4FA3',
  ),
  scheme(
    'forest',
    'Forest',
    '1E2B20',
    'FFFFFF',
    '375E43',
    'E9F2EB',
    ['217346', '4EA72E', '92D050', 'FFC000', '3E8E8B', '70AD47'],
    '217346',
  ),
  scheme(
    'cream',
    'Cream',
    '3B3226',
    'FBF6EC',
    '6E5F49',
    'F0E7D5',
    ['C0A062', '8A6D3B', 'B65C33', '6E8B5E', '4E7CA1', '9E5D74'],
    '8A6D3B',
  ),
  scheme(
    'rose',
    'Rose',
    '3A1F33',
    'FFFFFF',
    '6E3A62',
    'F8ECF4',
    ['B44582', 'D96FA8', 'E9A6C9', '7030A0', 'E97132', '4E7CA1'],
    'B44582',
  ),
  scheme(
    'graphite',
    'Graphite',
    'F2F2F2',
    '1E1E1E',
    'D9D9D9',
    '333333',
    ['4FC3F7', 'FFB74D', '81C784', 'BA68C8', 'E57373', '90A4AE'],
    '4FC3F7',
  ),
  scheme(
    'midnight',
    'Midnight',
    'EAF2FF',
    '0F1C2E',
    'BCD0EC',
    '1E3350',
    ['4A9EDE', '63C7B2', 'F2C14E', 'E4718D', '9B8CDE', '7FB069'],
    '6FB3EC',
  ),
]

export const FONT_SCHEMES: readonly ThemeFontScheme[] = [
  { id: 'office', name: 'Office', major: 'Calibri Light', minor: 'Calibri' },
  { id: 'arial', name: 'Arial', major: 'Arial', minor: 'Arial' },
  { id: 'georgia', name: 'Georgia', major: 'Georgia', minor: 'Georgia' },
  { id: 'candara', name: 'Candara', major: 'Candara', minor: 'Candara' },
  { id: 'segoe', name: 'Segoe UI', major: 'Segoe UI', minor: 'Segoe UI' },
  { id: 'trebuchet', name: 'Trebuchet MS', major: 'Trebuchet MS', minor: 'Calibri' },
]

const FONT_BY_THEME: Record<string, string> = {
  office: 'office',
  ember: 'trebuchet',
  indigo: 'segoe',
  forest: 'candara',
  cream: 'georgia',
  rose: 'trebuchet',
  graphite: 'segoe',
  midnight: 'office',
}

export const THEME_PRESETS: readonly ThemePreset[] = COLOR_SCHEMES.map((colors) => ({
  id: colors.id,
  name: colors.name,
  colors,
  fonts: FONT_SCHEMES.find((fonts) => fonts.id === FONT_BY_THEME[colors.id]) ?? FONT_SCHEMES[0]!,
}))

function hexToRgb(hex: string): [number, number, number] | null {
  const value = hex.replace(/^#/, '')
  if (!/^[0-9A-Fa-f]{6}$/.test(value)) return null
  return [
    parseInt(value.slice(0, 2), 16),
    parseInt(value.slice(2, 4), 16),
    parseInt(value.slice(4, 6), 16),
  ]
}

function rgbToHsl([red, green, blue]: readonly [number, number, number]): [number, number, number] {
  const r = red / 255
  const g = green / 255
  const b = blue / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const luminance = (max + min) / 2
  if (max === min) return [0, 0, luminance]
  const delta = max - min
  const saturation = luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min)
  let hue: number
  if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0)
  else if (max === g) hue = (b - r) / delta + 2
  else hue = (r - g) / delta + 4
  return [hue * 60, saturation, luminance]
}

function hslToRgb(hue: number, saturation: number, luminance: number): [number, number, number] {
  const channel = (offset: number): number => {
    const k = (offset + hue / 30) % 12
    const a = saturation * Math.min(luminance, 1 - luminance)
    const value = luminance - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))
    return Math.round(value * 255)
  }
  return [channel(0), channel(8), channel(4)]
}

/// Excel's tint transform: scale HSL luminance toward black (tint < 0) or
/// white (tint > 0) — same math as the sidecar's color resolver.
export function applyTint(hex: string, tint: number): string {
  const rgb = hexToRgb(hex)
  if (!rgb || tint === 0) return hex.toUpperCase()
  const [hue, saturation, luminance] = rgbToHsl(rgb)
  const adjusted = tint < 0 ? luminance * (1 + tint) : luminance * (1 - tint) + tint
  const [red, green, blue] = hslToRgb(hue, saturation, Math.min(Math.max(adjusted, 0), 1))
  const hex2 = (value: number): string => value.toString(16).padStart(2, '0').toUpperCase()
  return `#${hex2(red)}${hex2(green)}${hex2(blue)}`
}

/// Re-resolves theme-provenance colors and theme-scheme fonts against a new
/// palette / font pair, returning a new styles array (untouched styles are
/// reused by reference).
export function rethemeStyles(
  styles: readonly WorkbookCellStyle[],
  palette: readonly string[],
  fonts: { major: string; minor: string } | null,
): WorkbookCellStyle[] {
  return styles.map((style) => {
    let next = style
    const resolve = (slot: number | undefined, tint: number | undefined): string | null => {
      if (slot === undefined) return null
      const base = palette[slot]
      if (base === undefined) return null
      return applyTint(base, tint ?? 0)
    }
    const fontColor = resolve(style.fontColorTheme, style.fontColorTint)
    if (fontColor !== null && fontColor !== style.fontColor) {
      next = { ...next, fontColor }
    }
    const fillColor = resolve(style.fillColorTheme, style.fillColorTint)
    if (fillColor !== null && fillColor !== style.fillColor) {
      next = { ...next, fillColor }
    }
    if (fonts !== null && style.fontScheme !== undefined) {
      const family = style.fontScheme === 'major' ? fonts.major : fonts.minor
      if (family !== style.fontFamily) next = { ...next, fontFamily: family }
    }
    return next
  })
}

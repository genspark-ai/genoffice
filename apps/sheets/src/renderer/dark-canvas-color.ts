/**
 * Univer inverts canvas colors in dark mode with a hue-preserving matrix. For
 * the sheet grid that is right, but the formula bar is a small document whose
 * page is white: inverted white is pure #000, which reads as a black hole next
 * to the #1e1e1e chrome. Docs solves the same problem for its dark paper by
 * compressing the inverted color onto the paper tone (editor/dark-page.ts);
 * this service does the same for every canvas color Univer paints, using
 * Univer's supported dependency-override seam rather than its private
 * `_invertAlgo` field.
 *
 * Authored white lands exactly on the chrome floor, authored black stays
 * white and hues keep their direction. Theme tokens are already dark in the
 * dark theme, so they are left alone.
 */
import type { DependencyOverride } from '@univerjs/core'
import { ThemeService } from '@univerjs/core'
import { CanvasColorService, ICanvasColorService } from '@univerjs/engine-render'

/** the chrome's dark background: `#1e1e1e` */
export const SHEETS_DARK_FLOOR = [0x1e, 0x1e, 0x1e] as const

export function compressRgb(
  rgb: readonly number[],
  floor: readonly number[] = SHEETS_DARK_FLOOR,
): [number, number, number] {
  return [0, 1, 2].map((i) => Math.round(floor[i]! + (rgb[i]! * (255 - floor[i]!)) / 255)) as [
    number,
    number,
    number,
  ]
}

const HEX_RE = /^#?([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i
const RGB_RE =
  /^rgba?\(\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/i

const hex2 = (value: number) =>
  Math.max(0, Math.min(255, Math.round(value)))
    .toString(16)
    .padStart(2, '0')

/**
 * Compress one inverted CSS color onto the dark floor. Keeps the alpha and
 * the hex / rgb spelling; values Univer never inverts (theme tokens,
 * keywords) come back unchanged.
 */
export function compressCssColor(color: string): string {
  const s = color.trim()
  const hex = HEX_RE.exec(s)
  if (hex) {
    let h = hex[1]!
    // alpha is the trailing nibble/byte of a #rgba / #rrggbbaa color
    const alpha = h.length === 4 || h.length === 8 ? h.slice(-(h.length === 4 ? 1 : 2)) : ''
    if (h.length <= 4) h = [...h.slice(0, 3)].map((c) => c + c).join('')
    const n = parseInt(h.slice(0, 6), 16)
    const [r, g, b] = compressRgb([(n >> 16) & 255, (n >> 8) & 255, n & 255])
    return `#${hex2(r)}${hex2(g)}${hex2(b)}${alpha}`
  }
  const rgb = RGB_RE.exec(s)
  if (rgb) {
    const [r, g, b] = compressRgb([Number(rgb[1]), Number(rgb[2]), Number(rgb[3])])
    if (rgb[4] === undefined) return `rgb(${r},${g},${b})`
    return `rgba(${r},${g},${b}, ${rgb[4]})`
  }
  return color
}

/** Univer's CanvasColorService with its inverted output compressed onto #1e1e1e. */
export class CompressingCanvasColorService extends CanvasColorService {
  private readonly _compressed = new Map<string, string>()

  constructor(private readonly _theme: ThemeService) {
    super(_theme)
  }

  override getRenderColor(color: string): string {
    if (!this._theme.darkMode) return color
    if (this._theme.isValidThemeColor(color)) return super.getRenderColor(color)
    const hit = this._compressed.get(color)
    if (hit !== undefined) return hit
    const out = compressCssColor(super.getRenderColor(color))
    this._compressed.set(color, out)
    return out
  }
}

/** register the service through Univer's public dependency-override seam */
export const sheetsDarkCanvasOverride: DependencyOverride = [
  [
    ICanvasColorService,
    {
      useFactory: (theme: ThemeService) => new CompressingCanvasColorService(theme),
      deps: [ThemeService],
    },
  ],
]

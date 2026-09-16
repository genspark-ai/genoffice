import type { FillSpec, PatternType, StyleColor } from '../domain/style-color'
import { PATTERN_TYPES } from '../domain/style-color'

/// Read-only view of xl/styles.xml for echoing colors back the way they are
/// stored (theme slot + tint, pattern, gradient) — the sidecar only reports
/// the resolved rgb.
export interface StylesheetFormats {
  /// cellXfs index → fill / font list indexes
  xfs: readonly { fillId: number; fontId: number }[]
  /// fills list; null for the `none` pattern or an unreadable fill
  fills: readonly (FillSpec | null)[]
  /// fonts list → explicit font color
  fontColors: readonly (StyleColor | undefined)[]
}

export function parseStylesheetFormats(stylesXml: string): StylesheetFormats {
  const xfs = elements(sectionInner(stylesXml, 'cellXfs'), 'xf').map((xf) => ({
    fillId: Number(attribute(xf, 'fillId') ?? 0),
    fontId: Number(attribute(xf, 'fontId') ?? 0),
  }))
  const fills = elements(sectionInner(stylesXml, 'fills'), 'fill').map(parseFill)
  const fontColors = elements(sectionInner(stylesXml, 'fonts'), 'font').map((font) => {
    const color = /<color\b[^>]*\/?>/.exec(font)?.[0]
    return color === undefined ? undefined : parseColor(color)
  })
  return { xfs, fills, fontColors }
}

function parseFill(fillXml: string): FillSpec | null {
  const gradient = /<gradientFill\b([^>]*)>([\s\S]*?)<\/gradientFill>/.exec(fillXml)
  if (gradient) {
    const attrs = gradient[1] ?? ''
    const stops = [
      ...gradient[2]!.matchAll(/<stop\b[^>]*position="([^"]*)"[^>]*>([\s\S]*?)<\/stop>/g),
    ]
      .map((stop) => {
        const color = parseColor(/<color\b[^>]*\/?>/.exec(stop[2] ?? '')?.[0] ?? '')
        return color === undefined ? null : { position: Number(stop[1]), color }
      })
      .filter((stop): stop is { position: number; color: StyleColor } => stop !== null)
    if (stops.length < 2) return null
    const type = attribute(attrs, 'type') === 'path' ? 'path' : undefined
    const num = (name: string): number | undefined => {
      const value = attribute(attrs, name)
      return value === undefined ? undefined : Number(value)
    }
    return {
      gradient: {
        ...(type ? { type } : {}),
        ...(type ? {} : num('degree') !== undefined ? { angle: num('degree') } : {}),
        ...(type && num('left') !== undefined ? { left: num('left') } : {}),
        ...(type && num('right') !== undefined ? { right: num('right') } : {}),
        ...(type && num('top') !== undefined ? { top: num('top') } : {}),
        ...(type && num('bottom') !== undefined ? { bottom: num('bottom') } : {}),
        stops,
      },
    }
  }
  const pattern = /<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/.exec(fillXml)
  if (!pattern) return null
  const patternType = attribute(pattern[1] ?? '', 'patternType')
  if (patternType === undefined || patternType === 'none') return null
  if (!(PATTERN_TYPES as readonly string[]).includes(patternType)) return null
  const inner = pattern[2] ?? ''
  const fg = parseColor(/<fgColor\b[^>]*\/?>/.exec(inner)?.[0] ?? '')
  const bg = parseColor(/<bgColor\b[^>]*\/?>/.exec(inner)?.[0] ?? '')
  // a pattern without a foreground is Excel's "automatic" color: not a choice to echo
  if (fg === undefined) return null
  return { pattern: patternType as PatternType, fg, ...(bg === undefined ? {} : { bg }) }
}

/// CT_Color → StyleColor; indexed / auto / system colors have no stable echo
function parseColor(colorXml: string): StyleColor | undefined {
  if (colorXml === '') return undefined
  const theme = attribute(colorXml, 'theme')
  if (theme !== undefined) {
    const index = Number(theme)
    if (!Number.isInteger(index) || index < 0 || index > 11) return undefined
    const tint = attribute(colorXml, 'tint')
    const value = tint === undefined ? 0 : Number(tint)
    return value === 0 || !Number.isFinite(value)
      ? { theme: index }
      : { theme: index, tint: Math.round(value * 1e6) / 1e6 }
  }
  const rgb = attribute(colorXml, 'rgb')
  if (rgb === undefined) return undefined
  const hex = rgb.length === 8 ? rgb.slice(2) : rgb
  return /^[0-9A-Fa-f]{6}$/.test(hex) ? `#${hex.toUpperCase()}` : undefined
}

function sectionInner(xml: string, tag: string): string {
  return new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1] ?? ''
}

function elements(inner: string, tag: string): string[] {
  return [
    ...inner.matchAll(new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`, 'g')),
  ].map((match) => match[0])
}

function attribute(element: string, name: string): string | undefined {
  return new RegExp(`(?<![\\w:.-])${name}="([^"]*)"`).exec(element)?.[1]
}

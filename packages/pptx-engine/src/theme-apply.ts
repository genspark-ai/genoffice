/**
 * Theme application (Design tab theme gallery) — rewrites <a:clrScheme> and
 * <a:fontScheme> of every ppt/theme/theme*.xml in the package, in place.
 *
 * Semantics match PowerPoint: content referencing schemeClr / +mj-lt / +mn-lt
 * follows the new theme; content with explicit srgbClr / explicit fonts stays
 * as-is. After rewriting, the caller (main process) must save→reopen to reparse
 * the inheritance chain so elements' resolved colors/fonts refresh.
 *
 * Exception: theme parts are the sole exemption from the "never write back
 * layout/master/theme" rule — the whole point of the theme gallery is to replace
 * the theme; layout/master themselves are still untouched.
 */
import type { OpenedPptx } from './index'
import { escapeXmlAttr } from './xml-utils'

/** One theme definition: 12 scheme colors + optional heading/body fonts. */
export interface ThemeSpec {
  name: string
  /** dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink → #RRGGBB (upper or lower case) */
  colors: Record<string, string>
  /** Heading (title) Latin font; original theme kept if unset */
  majorFont?: string
  /** Body Latin font; original theme kept if unset */
  minorFont?: string
}

const SCHEME_KEYS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const

const hex6 = (c: string) => c.replace(/^#/, '').slice(0, 6).toUpperCase()

/** Rewrite a theme XML's clrScheme/fontScheme; returned as-is if the scheme structure is missing. */
export function patchThemeXml(xml: string, spec: ThemeSpec): string {
  let out = xml
  // Colors: replace the whole <a:dk1>…</a:dk1> section per key (sysClr/srgbClr both become srgbClr)
  for (const key of SCHEME_KEYS) {
    const color = spec.colors[key]
    if (!color) continue
    out = out.replace(
      new RegExp(`<a:${key}>[\\s\\S]*?</a:${key}>`),
      `<a:${key}><a:srgbClr val="${hex6(color)}"/></a:${key}>`,
    )
  }
  out = out.replace(/(<a:clrScheme name=")[^"]*(")/, `$1${escapeXmlAttr(spec.name)}$2`)
  // Fonts: replace only the major/minor Latin typeface; East Asian / cs kept
  if (spec.majorFont) {
    out = out.replace(
      /(<a:majorFont>[\s\S]*?<a:latin[^>]*typeface=")[^"]*(")/,
      `$1${escapeXmlAttr(spec.majorFont)}$2`,
    )
  }
  if (spec.minorFont) {
    out = out.replace(
      /(<a:minorFont>[\s\S]*?<a:latin[^>]*typeface=")[^"]*(")/,
      `$1${escapeXmlAttr(spec.minorFont)}$2`,
    )
  }
  return out
}

const THEME_PART_RE = /^ppt\/theme\/theme\d+\.xml$/

/** Apply the theme to all theme parts in the package; returns the number rewritten. */
export function applyThemeToArchive(opened: OpenedPptx, spec: ThemeSpec): number {
  const { archive } = opened
  let patched = 0
  for (const path of [...archive.entries.keys()]) {
    if (!THEME_PART_RE.test(path)) continue
    const xml = archive.readText(path)
    if (!xml) continue
    const next = patchThemeXml(xml, spec)
    if (next !== xml) {
      archive.entries.set(path, Buffer.from(next, 'utf8'))
      patched++
    }
  }
  return patched
}

// ── Explicit color remapping ──────────────────────────────────────────
// Real-world decks (especially AI-generated/exported ones) use almost exclusively
// explicit srgbClr colors that never reference the scheme — swapping only the theme
// parts changes nothing visually. Here the deck's existing explicit colors are
// remapped wholesale onto the new theme palette:
//   · neutrals (low saturation / near black-white) map onto the new theme's
//     dk1↔lt1 axis, bucketed by lightness;
//   · chromatic colors are clustered by hue, and clusters are assigned to
//     accent1..6 in order of frequency;
//   · output replaces only hue/base saturation, keeping the original lightness —
//     shading levels and contrast structure stay intact.

interface Hsl {
  /** 0..360 */
  h: number
  /** 0..1 */
  s: number
  /** 0..1 */
  l: number
}

function hexToHsl(hex: string): Hsl {
  const h6 = hex.replace(/^#/, '')
  const r = (parseInt(h6.slice(0, 2), 16) || 0) / 255
  const g = (parseInt(h6.slice(2, 4), 16) || 0) / 255
  const b = (parseInt(h6.slice(4, 6), 16) || 0) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (d === 0) return { h: 0, s: 0, l }
  const s = d / (1 - Math.abs(2 * l - 1))
  let h: number
  if (max === r) h = 60 * (((g - b) / d) % 6)
  else if (max === g) h = 60 * ((b - r) / d + 2)
  else h = 60 * ((r - g) / d + 4)
  if (h < 0) h += 360
  return { h, s: Math.min(1, s), l }
}

function hslToHex({ h, s, l }: Hsl): string {
  const c = (1 - Math.abs(2 * l - 1)) * s
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1))
  const m = l - c / 2
  const [r, g, b] =
    h < 60
      ? [c, x, 0]
      : h < 120
        ? [x, c, 0]
        : h < 180
          ? [0, c, x]
          : h < 240
            ? [0, x, c]
            : h < 300
              ? [x, 0, c]
              : [c, 0, x]
  const to255 = (v: number) => Math.max(0, Math.min(255, Math.round((v + m) * 255)))
  return [to255(r), to255(g), to255(b)]
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

const hueDist = (a: number, b: number): number => {
  const d = Math.abs(a - b) % 360
  return d > 180 ? 360 - d : d
}

/** Low saturation / near black-white counts as neutral (mapped along the dk1↔lt1 axis) */
const isNeutral = ({ s, l }: Hsl) => s < 0.15 || l > 0.95 || l < 0.07

const SRGB_RE = /<a:srgbClr val="([0-9A-Fa-f]{6})"/g

/** Count explicit srgbClr colors in a batch of slide XML (uppercase hex → occurrence count). */
export function collectExplicitColors(xmls: string[]): Map<string, number> {
  const counts = new Map<string, number>()
  for (const xml of xmls) {
    for (const m of xml.matchAll(SRGB_RE)) {
      const hex = m[1]!.toUpperCase()
      counts.set(hex, (counts.get(hex) ?? 0) + 1)
    }
  }
  return counts
}

/**
 * Mapping table from explicit colors to new theme colors. Identity mappings are omitted.
 * Cluster assignment is deterministic: iterated by (occurrence count desc, hex asc).
 */
export function buildColorMap(counts: Map<string, number>, spec: ThemeSpec): Map<string, string> {
  const neutralFor = (l: number): string => {
    const key = l > 0.85 ? 'lt1' : l > 0.55 ? 'lt2' : l > 0.28 ? 'dk2' : 'dk1'
    return hex6(spec.colors[key] ?? (l > 0.5 ? 'FFFFFF' : '000000'))
  }
  const accents = [1, 2, 3, 4, 5, 6]
    .map((i) => spec.colors[`accent${i}`])
    .filter((c): c is string => !!c)
    .map((c) => hexToHsl(c))

  const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1))
  const clusters: Array<{ hue: number }> = []
  const map = new Map<string, string>()
  for (const [hex] of sorted) {
    const hsl = hexToHsl(hex)
    let out: string
    if (isNeutral(hsl)) {
      out = neutralFor(hsl.l)
    } else if (accents.length === 0) {
      continue
    } else {
      let idx = clusters.findIndex((c) => hueDist(c.hue, hsl.h) <= 35)
      if (idx < 0) {
        clusters.push({ hue: hsl.h })
        idx = clusters.length - 1
      }
      const target = accents[idx % accents.length]!
      // Target hue + target base saturation, keeping original lightness (shading levels unchanged)
      out = hslToHex({ h: target.h, s: target.s, l: hsl.l })
    }
    if (out !== hex) map.set(hex, out)
  }
  return map
}

/** Recolor a slide XML per the mapping table (only <a:srgbClr val> touched). */
export function recolorXml(xml: string, map: Map<string, string>): string {
  if (map.size === 0) return xml
  return xml.replace(SRGB_RE, (full, hex: string) => {
    const to = map.get(hex.toUpperCase())
    return to ? full.replace(hex, to) : full
  })
}

/**
 * Remap all slides' explicit colors onto the theme palette (rewriting archive
 * entries directly). Requires the model to have no unpersisted edits (caller should
 * save→reopen first), otherwise dirty elements would overwrite with old colors on
 * save. Returns the number of distinct colors mapped.
 */
export function remapDeckColors(opened: OpenedPptx, spec: ThemeSpec): number {
  const { archive, deck } = opened
  const paths = deck.slides.map((s) => s.path)
  const xmls = paths.map((p) => archive.readText(p)).filter((x): x is string => !!x)
  const map = buildColorMap(collectExplicitColors(xmls), spec)
  if (map.size === 0) return 0
  for (const path of paths) {
    const xml = archive.readText(path)
    if (!xml) continue
    const next = recolorXml(xml, map)
    if (next !== xml) archive.entries.set(path, Buffer.from(next, 'utf8'))
  }
  return map.size
}

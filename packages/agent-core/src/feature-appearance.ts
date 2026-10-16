import type { FeatureRecord } from './features-skill'

/**
 * The appearance kind's contract: what a feature may change about the shell's
 * own chrome, plus the markdown reading theme.
 *
 * Chrome is a choice from host-defined names, never a raw value: a mod picks
 * "violet" or "surface-subtle", the literals live in the host's stylesheet, and
 * dark mode and contrast stay the host's business.
 *
 * The markdown reading theme (`markdown`) re-skins the md reader's paper the
 * way a Typora theme does. It is either a host palette name, or a small custom
 * object ({ paper, ink, accent?, font? }) the validator checks for readability
 * before it is ever applied. Display only, screen media only: the document's
 * text, its saving, and the print/export output are never touched by it.
 */

export const APPEARANCE_BACKGROUNDS = [
  'default',
  'surface',
  'surface-subtle',
  'accent-tint',
] as const
export type AppearanceBackground = (typeof APPEARANCE_BACKGROUNDS)[number]

export const APPEARANCE_ACCENTS = ['default', 'blue', 'violet', 'teal', 'amber', 'rose'] as const
export type AppearanceAccent = (typeof APPEARANCE_ACCENTS)[number]

/**
 * The markdown reader's paper themes: the app's own picker palette (drawn in
 * apps/markdown/renderer/markdown-themes.css — Notion, Drake, Blubook, the Nord
 * palette, Newsprint; the license notes of the reference designs live there)
 * plus `default`, which means the paper follows the light/dark document theme —
 * or the mod appearance, when a feature declares one. A theme the host has not
 * drawn is not in the list; for a look of one's own, a feature declares the
 * custom object instead.
 */
export const MD_READ_THEMES = [
  'default',
  'notion-light',
  'notion-dark',
  'drake-dark',
  'blubook',
  'nord',
  'newsprint',
] as const
export type MarkdownReadTheme = (typeof MD_READ_THEMES)[number]

/** The body font families a custom reading theme may pick from. */
export const MD_THEME_FONTS = ['sans', 'serif', 'mono'] as const
export type MarkdownThemeFont = (typeof MD_THEME_FONTS)[number]

export interface MarkdownCustomTheme {
  /** the paper color, a #rgb or #rrggbb hex */
  paper: string
  /** the body text color; the validator refuses an unreadable pairing */
  ink: string
  /** the link color; absent keeps the host's link color */
  accent?: string
  /** the body font family; absent keeps the host default */
  font?: MarkdownThemeFont
}

export interface FeatureAppearance {
  /** the chrome behind the content areas */
  background?: AppearanceBackground
  /** the accent family: buttons, chips, focus rings, selection tints */
  accent?: AppearanceAccent
  /** how the markdown reader's paper is skinned; absent means the host default */
  markdown?: MarkdownReadTheme | MarkdownCustomTheme
}

export interface AppearanceValidationError {
  field: string
  message: string
}

export type AppearanceValidation =
  { ok: true; appearance: FeatureAppearance } | { ok: false; errors: AppearanceValidationError[] }

/** minimum WCAG contrast ratio between a custom theme's paper and its ink */
const INK_CONTRAST_MIN = 4.5
/** minimum ratio between the paper and a custom link color */
const ACCENT_CONTRAST_MIN = 3

/** Parse a #rgb / #rrggbb hex into lowercase #rrggbb; undefined otherwise. */
function normalizeHex(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const m = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.exec(value.trim())
  if (!m) return undefined
  const hex = m[1]!
  return hex.length === 3
    ? '#' +
        [...hex]
          .map((ch) => ch + ch)
          .join('')
          .toLowerCase()
    : '#' + hex.toLowerCase()
}

function channelLuminance(channel: number): number {
  const c = channel / 255
  return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}

/** WCAG relative luminance of a #rrggbb hex (normalized input). */
function relativeLuminance(hex: string): number {
  return (
    0.2126 * channelLuminance(parseInt(hex.slice(1, 3), 16)) +
    0.7152 * channelLuminance(parseInt(hex.slice(3, 5), 16)) +
    0.0722 * channelLuminance(parseInt(hex.slice(5, 7), 16))
  )
}

/** WCAG contrast ratio between two #rgb / #rrggbb hex colors (1 to 21). */
export function contrastRatio(a: string, b: string): number {
  const la = normalizeHex(a)
  const lb = normalizeHex(b)
  if (!la || !lb) return 1
  const x = relativeLuminance(la)
  const y = relativeLuminance(lb)
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05)
}

/**
 * The CSS custom properties a custom reading theme sets on the root, and the
 * color scheme the paper's lightness implies. Pure, so the host applies them
 * and the tests can read them: the derived tones (borders, code backgrounds)
 * are mixed in the stylesheet, not here.
 */
export function markdownCustomThemeVars(theme: MarkdownCustomTheme): {
  scheme: 'light' | 'dark'
  vars: Record<string, string>
} {
  const paper = normalizeHex(theme.paper)!
  const ink = normalizeHex(theme.ink)!
  return {
    scheme: relativeLuminance(paper) < 0.25 ? 'dark' : 'light',
    vars: {
      '--md-custom-paper': paper,
      '--md-custom-ink': ink,
      ...(theme.accent ? { '--md-custom-accent': normalizeHex(theme.accent)! } : {}),
    },
  }
}

/** The names a `markdown` value may take, for the assistant's schema text. */
export function markdownThemeSummary(): string {
  return MD_READ_THEMES.join(', ') + ', or { paper, ink, accent?, font? }'
}

/** Validate an appearance payload from the assistant or a mod file. */
export function validateAppearance(value: unknown): AppearanceValidation {
  const errors: AppearanceValidationError[] = []
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'appearance', message: 'needs an appearance object' }] }
  }
  const input = value as Record<string, unknown>
  const unknown = Object.keys(input).filter(
    (key) => key !== 'background' && key !== 'accent' && key !== 'markdown',
  )
  for (const key of unknown) {
    errors.push({ field: 'appearance.' + key, message: 'is not an appearance field' })
  }
  if (input.background !== undefined) {
    if (!(APPEARANCE_BACKGROUNDS as readonly string[]).includes(String(input.background))) {
      errors.push({
        field: 'appearance.background',
        message: 'must be one of ' + APPEARANCE_BACKGROUNDS.join(', '),
      })
    }
  }
  if (input.accent !== undefined) {
    if (!(APPEARANCE_ACCENTS as readonly string[]).includes(String(input.accent))) {
      errors.push({
        field: 'appearance.accent',
        message: 'must be one of ' + APPEARANCE_ACCENTS.join(', '),
      })
    }
  }
  let markdown: MarkdownReadTheme | MarkdownCustomTheme | undefined
  if (input.markdown !== undefined) {
    if (typeof input.markdown === 'string') {
      if (!(MD_READ_THEMES as readonly string[]).includes(input.markdown)) {
        errors.push({
          field: 'appearance.markdown',
          message: 'must be one of ' + markdownThemeSummary(),
        })
      } else {
        markdown = input.markdown as MarkdownReadTheme
      }
    } else if (typeof input.markdown === 'object' && !Array.isArray(input.markdown)) {
      markdown = validateCustomTheme(input.markdown as Record<string, unknown>, errors)
    } else {
      errors.push({
        field: 'appearance.markdown',
        message: 'must be one of ' + markdownThemeSummary(),
      })
    }
  }
  if (
    input.background === undefined &&
    input.accent === undefined &&
    input.markdown === undefined &&
    errors.length === 0
  ) {
    errors.push({
      field: 'appearance',
      message: 'needs a background, an accent or a markdown theme',
    })
  }
  if (errors.length) return { ok: false, errors }
  return {
    ok: true,
    appearance: {
      ...(input.background === undefined
        ? {}
        : { background: input.background as AppearanceBackground }),
      ...(input.accent === undefined ? {} : { accent: input.accent as AppearanceAccent }),
      ...(markdown ? { markdown } : {}),
    },
  }
}

/** Validate the custom reading-theme object; a bad pairing is refused with the ratio. */
function validateCustomTheme(
  input: Record<string, unknown>,
  errors: AppearanceValidationError[],
): MarkdownCustomTheme | undefined {
  const unknown = Object.keys(input).filter(
    (key) => key !== 'paper' && key !== 'ink' && key !== 'accent' && key !== 'font',
  )
  for (const key of unknown) {
    errors.push({ field: 'appearance.markdown.' + key, message: 'is not a theme field' })
  }
  const paper = normalizeHex(input.paper)
  const ink = normalizeHex(input.ink)
  if (!paper) {
    errors.push({ field: 'appearance.markdown.paper', message: 'needs a #rgb or #rrggbb hex' })
  }
  if (!ink) {
    errors.push({ field: 'appearance.markdown.ink', message: 'needs a #rgb or #rrggbb hex' })
  }
  let accent: string | undefined
  if (input.accent !== undefined) {
    accent = normalizeHex(input.accent)
    if (!accent) {
      errors.push({ field: 'appearance.markdown.accent', message: 'needs a #rgb or #rrggbb hex' })
    }
  }
  let font: MarkdownThemeFont | undefined
  if (input.font !== undefined) {
    if (
      typeof input.font !== 'string' ||
      !(MD_THEME_FONTS as readonly string[]).includes(input.font)
    ) {
      errors.push({
        field: 'appearance.markdown.font',
        message: 'must be one of ' + MD_THEME_FONTS.join(', '),
      })
    } else {
      font = input.font as MarkdownThemeFont
    }
  }
  if (paper && ink) {
    const ratio = contrastRatio(paper, ink)
    if (ratio < INK_CONTRAST_MIN) {
      errors.push({
        field: 'appearance.markdown',
        message:
          'the paper and ink are too close (' +
          ratio.toFixed(1) +
          ':1); a reading theme needs at least ' +
          INK_CONTRAST_MIN +
          ':1',
      })
    }
  }
  if (paper && accent) {
    const ratio = contrastRatio(paper, accent)
    if (ratio < ACCENT_CONTRAST_MIN) {
      errors.push({
        field: 'appearance.markdown.accent',
        message:
          'the accent is too close to the paper (' +
          ratio.toFixed(1) +
          ':1); links need at least ' +
          ACCENT_CONTRAST_MIN +
          ':1',
      })
    }
  }
  if (errors.length) return undefined
  return {
    paper: paper!,
    ink: ink!,
    ...(accent ? { accent } : {}),
    ...(font ? { font } : {}),
  }
}

/** Which appearance applies: the most recently changed feature that declares one. */
export function activeAppearance(features: readonly FeatureRecord[]): FeatureAppearance | null {
  let best: FeatureRecord | null = null
  for (const feature of features) {
    if (!feature.appearance) continue
    if (!best || feature.updatedAt > best.updatedAt) best = feature
  }
  return best?.appearance ?? null
}

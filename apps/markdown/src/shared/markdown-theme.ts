/**
 * Markdown document themes: a paper-scoped palette selected in the markdown
 * app's ribbon (`data-markdown-theme` on the document root; the variable
 * blocks live in renderer/markdown-themes.css, loaded after styles.css so a
 * theme's `.doc-page` scope wins over the light/dark document theme).
 *
 * Licenses of the reference designs each palette is derived from:
 * - notion-light / notion-dark: MIT, github.com/adrian-fuertes/typora-notion-theme
 * - drake-dark: MIT, github.com/liangjingkanji/DrakeTyporaTheme
 * - blubook: Apache-2.0, github.com/FishionYu/typora-blubook-theme
 * - nord: the Nord palette (nordtheme.com) — clean-room reimplementation,
 *   not a port of the GPL typora theme
 * - newsprint: palette values from Typora's official newsprint theme
 *   (github.com/typora/typora-default-themes carries no license file) —
 *   reimplementation, rules written from scratch
 */
export const MARKDOWN_THEME_IDS = [
  'default',
  'notion-light',
  'notion-dark',
  'drake-dark',
  'blubook',
  'nord',
  'newsprint',
] as const

export type MarkdownTheme = (typeof MARKDOWN_THEME_IDS)[number]

export function normalizeMarkdownTheme(value: unknown): MarkdownTheme {
  return (MARKDOWN_THEME_IDS as readonly string[]).includes(String(value))
    ? (String(value) as MarkdownTheme)
    : 'default'
}

/** Dark themes flip the paper regardless of the light/dark document theme.
 *  Blubook stays light: its signature is a light paper with deep-blue code fences. */
export function isMarkdownThemeDark(theme: MarkdownTheme): boolean {
  return theme === 'notion-dark' || theme === 'drake-dark' || theme === 'nord'
}

/**
 * Which paper the editor shows, from the two people who may ask: the user's
 * ribbon pick wins outright; only a `default` pick defers to the mod
 * appearance (a palette id, or a custom object that renders as 'custom').
 */
export function resolvePaperTheme(
  user: MarkdownTheme,
  mod: MarkdownTheme | import('@genoffice/agent-core').MarkdownCustomTheme | undefined,
): MarkdownTheme | 'custom' {
  if (user !== 'default') return user
  if (typeof mod === 'string' && mod !== 'default') return mod
  if (mod && typeof mod === 'object') return 'custom'
  return 'default'
}

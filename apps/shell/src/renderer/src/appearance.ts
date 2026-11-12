/** user-adjustable look of the shell; stored per machine, applied before first paint */
export type UiFont = 'system' | 'google-sans' | 'sans' | 'serif' | 'mono'
export type Contrast = 'default' | 'strong'
/** how much tool/step detail the AI chat panels show while the agent works */
export type ConversationSteps = 'default' | 'verbose'

export interface ThemeColors {
  /** '#rrggbb' */
  background: string
  /** '#rrggbb' */
  foreground: string
  /** '#rrggbb' */
  accent: string
}

export interface ThemePreset extends ThemeColors {
  id: string
  name: string
}

export interface Appearance {
  conversationSteps: ConversationSteps
  contrast: Contrast
  font: UiFont
  /** colors used while the effective theme is dark */
  dark: ThemeColors
  /** colors used while the effective theme is light */
  light: ThemeColors
  /** false = opaque surfaces instead of glass over the window material */
  transparency: boolean
  /** ask before deleting instead of the undo toast */
  confirmDelete: boolean
}

export const DARK_PRESETS: readonly ThemePreset[] = [
  {
    id: 'default-dark',
    name: 'Default Dark',
    background: '#070709',
    foreground: '#ffffff',
    accent: '#2f71e3',
  },
  {
    id: 'midnight',
    name: 'Midnight',
    background: '#0a0f1e',
    foreground: '#e6edf7',
    accent: '#5b8cff',
  },
  {
    id: 'graphite',
    name: 'Graphite',
    background: '#121212',
    foreground: '#ececec',
    accent: '#8ab4f8',
  },
  { id: 'forest', name: 'Forest', background: '#0b1210', foreground: '#e4efe9', accent: '#3dd68c' },
  { id: 'ember', name: 'Ember', background: '#120c0a', foreground: '#f3e9e4', accent: '#f5a623' },
]

export const LIGHT_PRESETS: readonly ThemePreset[] = [
  {
    id: 'default-light',
    name: 'Default Light',
    background: '#ffffff',
    foreground: '#232425',
    accent: '#0f7fff',
  },
  { id: 'paper', name: 'Paper', background: '#faf8f5', foreground: '#2b2623', accent: '#b5541c' },
  { id: 'mist', name: 'Mist', background: '#f3f6fa', foreground: '#1b2430', accent: '#2f71e3' },
  { id: 'mint', name: 'Mint', background: '#f3faf6', foreground: '#14261d', accent: '#12a06b' },
]

export const DEFAULT_APPEARANCE: Appearance = {
  conversationSteps: 'default',
  contrast: 'default',
  font: 'system',
  dark: pick(DARK_PRESETS[0]!),
  light: pick(LIGHT_PRESETS[0]!),
  transparency: true,
  confirmDelete: false,
}

function pick(p: ThemeColors): ThemeColors {
  return { background: p.background, foreground: p.foreground, accent: p.accent }
}

/** id of the preset whose three colors match exactly, else 'custom' */
export function matchPreset(kind: 'dark' | 'light', colors: ThemeColors): string {
  const list = kind === 'dark' ? DARK_PRESETS : LIGHT_PRESETS
  const hit = list.find(
    (p) =>
      p.background === colors.background &&
      p.foreground === colors.foreground &&
      p.accent === colors.accent,
  )
  return hit ? hit.id : 'custom'
}

export const STORAGE_KEY = 'genoffice.appearance'

export const FONT_STACKS: Record<UiFont, string | null> = {
  system: null,
  'google-sans':
    "'Google Sans', 'Google Sans Text', 'Product Sans', 'Inter', 'Segoe UI', system-ui, var(--ui-cjk), 'Microsoft YaHei', sans-serif",
  sans: "'Inter', 'Segoe UI', system-ui, var(--ui-cjk), 'Microsoft YaHei', sans-serif",
  serif: "'Georgia', 'Noto Serif', var(--ui-cjk), 'Microsoft YaHei', serif",
  mono: "'Cascadia Mono', 'Consolas', ui-monospace, var(--ui-cjk), 'Microsoft YaHei', monospace",
}

/** '#abc' / 'abc' / '#aabbcc' → '#aabbcc' (lower-case), anything else → null */
export function parseHex(input: string): string | null {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(input.trim())
  if (!m) return null
  let hex = m[1]!.toLowerCase()
  if (hex.length === 3)
    hex = hex
      .split('')
      .map((c) => c + c)
      .join('')
  return `#${hex}`
}

function luminance(hex: string): number {
  const channel = (i: number): number => {
    const v = parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16) / 255
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * channel(0) + 0.7152 * channel(1) + 0.0722 * channel(2)
}

/** WCAG contrast ratio of two '#rrggbb' colors, 1..21 */
export function contrastRatio(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x) as [number, number]
  return (hi + 0.05) / (lo + 0.05)
}

/** text color that stays readable on an accent-filled control */
export function textOnAccent(accent: string): string {
  const onWhite = contrastRatio(accent, '#ffffff')
  // white reads better on mid-tone brand colors even when black wins by a hair
  return onWhite >= 4.4 || onWhite >= contrastRatio(accent, '#000000') ? '#ffffff' : '#000000'
}

function channels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ]
}

function toHex(n: number): string {
  return Math.max(0, Math.min(255, Math.round(n)))
    .toString(16)
    .padStart(2, '0')
}

/** linear mix of two '#rrggbb' colors; t=0 → a, t=1 → b */
export function mix(a: string, b: string, t: number): string {
  const [ar, ag, ab] = channels(a)
  const [br, bg, bb] = channels(b)
  return `#${toHex(ar + (br - ar) * t)}${toHex(ag + (bg - ag) * t)}${toHex(ab + (bb - ab) * t)}`
}

/** '#rrggbb' + alpha 0..1 → 'rgb(r g b / a%)' */
export function alpha(hex: string, a: number): string {
  const [r, g, b] = channels(hex)
  return `rgb(${r} ${g} ${b} / ${Math.round(a * 1000) / 10}%)`
}

function readColors(raw: unknown, fallback: ThemeColors): ThemeColors {
  const o = (raw && typeof raw === 'object' ? raw : {}) as Partial<
    Record<keyof ThemeColors, unknown>
  >
  const read = (v: unknown, d: string): string => (typeof v === 'string' ? (parseHex(v) ?? d) : d)
  return {
    background: read(o.background, fallback.background),
    foreground: read(o.foreground, fallback.foreground),
    accent: read(o.accent, fallback.accent),
  }
}

export function loadAppearance(storage: Pick<Storage, 'getItem'> | undefined): Appearance {
  try {
    const raw = storage?.getItem(STORAGE_KEY)
    if (!raw) return cloneAppearance(DEFAULT_APPEARANCE)
    const data = JSON.parse(raw) as Record<string, unknown>
    let dark = readColors(data.dark, DEFAULT_APPEARANCE.dark)
    let light = readColors(data.light, DEFAULT_APPEARANCE.light)
    // legacy single accent (before per-theme colors) applies to both themes
    if (typeof data.accent === 'string') {
      const legacy = parseHex(data.accent)
      if (legacy) {
        if (!data.dark) dark = { ...dark, accent: legacy }
        if (!data.light) light = { ...light, accent: legacy }
      }
    }
    return {
      conversationSteps: data.conversationSteps === 'verbose' ? 'verbose' : 'default',
      contrast: data.contrast === 'strong' ? 'strong' : 'default',
      font:
        typeof data.font === 'string' && Object.hasOwn(FONT_STACKS, data.font)
          ? (data.font as UiFont)
          : 'system',
      dark,
      light,
      transparency: data.transparency !== false,
      confirmDelete: data.confirmDelete === true,
    }
  } catch {
    return cloneAppearance(DEFAULT_APPEARANCE)
  }
}

export function cloneAppearance(a: Appearance): Appearance {
  return { ...a, dark: { ...a.dark }, light: { ...a.light } }
}

export function saveAppearance(
  storage: Pick<Storage, 'setItem' | 'removeItem'> | undefined,
  value: Appearance,
): void {
  try {
    const isDefault = JSON.stringify(value) === JSON.stringify(DEFAULT_APPEARANCE)
    if (isDefault) storage?.removeItem(STORAGE_KEY)
    else storage?.setItem(STORAGE_KEY, JSON.stringify(value))
  } catch {
    /* storage unavailable: the choice just won't persist */
  }
}

/** the theme the page is rendering right now: explicit data-theme wins, else the OS preference */
export function effectiveTheme(root: HTMLElement): 'dark' | 'light' {
  const attr = root.getAttribute('data-theme')
  if (attr === 'dark' || attr === 'light') return attr
  try {
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
  } catch {
    return 'light'
  }
}

/** every CSS custom property that applyAppearance may set, so a re-apply can clear stale ones */
const THEME_VARS = [
  '--surface',
  '--surface-subtle',
  '--chrome-bg',
  '--canvas',
  '--wall',
  '--bg-content',
  '--border',
  '--border-subtle',
  '--border-strong',
  '--border-hover',
  '--text',
  '--text-primary',
  '--text-secondary',
  '--text-tertiary',
  '--text-muted',
  '--text-dim',
  '--hover',
  '--pressed',
  '--active-bg',
  '--bg-hover',
  '--bg-hover-subtle',
  '--bg-hover-strong',
  '--bg-hover-accent',
  '--icon-muted',
  '--color-btn-primary',
  '--color-btn-primary-hover',
  '--color-btn-primary-text',
  '--color-brand-secondary',
  '--color-text-primary',
  '--color-text-secondary',
  '--color-text-tertiary',
  '--color-bg-page',
  '--color-bg-subtle',
  '--color-border-default',
  '--color-border-strong',
  '--color-border-brand',
  '--glass-bg',
  '--shell-border',
  '--segment-active-bg',
  '--accent',
  '--accent-hover',
  '--accent-text',
] as const

/** CSS custom properties for one theme's colors; pure so it can be unit-tested */
export function deriveThemeVars(
  kind: 'dark' | 'light',
  colors: ThemeColors,
  contrast: Contrast,
): Record<string, string> {
  const { background: B, foreground: F, accent: A } = colors
  const strong = contrast === 'strong'
  const dark = kind === 'dark'
  // text ramps: how far from the background toward the foreground
  const secondary = mix(B, F, strong ? 0.86 : dark ? 0.62 : 0.72)
  const tertiary = mix(B, F, strong ? 0.72 : dark ? 0.52 : 0.5)
  // hairlines: foreground at low alpha reads correctly on both light and dark
  const line = strong ? 1.9 : 1
  const vars: Record<string, string> = {
    '--surface': B,
    '--surface-subtle': mix(B, F, dark ? 0.06 : 0.025),
    '--chrome-bg': mix(B, F, dark ? 0.03 : 0),
    '--canvas': mix(B, F, dark ? 0.05 : 0.1),
    '--bg-content': dark ? B : mix(B, F, 0.015),
    '--border': alpha(F, (dark ? 0.08 : 0.11) * line),
    '--border-subtle': alpha(F, (dark ? 0.06 : 0.07) * line),
    '--border-strong': alpha(F, (dark ? 0.16 : 0.17) * line),
    '--border-hover': alpha(F, (dark ? 0.16 : 0.15) * line),
    '--text': F,
    '--text-primary': F,
    '--text-secondary': secondary,
    '--text-tertiary': tertiary,
    '--text-muted': tertiary,
    '--text-dim': secondary,
    '--hover': alpha(F, 0.05),
    '--pressed': alpha(F, 0.09),
    '--active-bg': alpha(F, 0.08),
    '--bg-hover': alpha(F, 0.05),
    '--bg-hover-subtle': alpha(F, 0.04),
    '--bg-hover-strong': alpha(F, 0.1),
    '--bg-hover-accent': alpha(F, 0.11),
    '--icon-muted': mix(B, F, 0.38),
    '--color-btn-primary': F,
    '--color-btn-primary-hover': mix(F, B, 0.12),
    '--color-btn-primary-text': B,
    '--color-brand-secondary': A,
    '--color-text-primary': F,
    '--color-text-secondary': secondary,
    '--color-text-tertiary': tertiary,
    '--color-bg-page': B,
    '--color-bg-subtle': mix(B, F, dark ? 0.06 : 0.04),
    '--color-border-default': alpha(F, (dark ? 0.08 : 0.09) * line),
    '--color-border-strong': alpha(F, (dark ? 0.16 : 0.17) * line),
    '--color-border-brand': A,
    '--accent': A,
    '--accent-hover': dark ? mix(A, '#ffffff', 0.15) : mix(A, '#000000', 0.08),
    '--accent-text': textOnAccent(A),
  }
  if (dark) {
    vars['--wall'] = mix(B, F, 0.09)
    vars['--glass-bg'] = alpha(F, 0.03)
    vars['--shell-border'] = alpha(F, 0.12 * (strong ? 1.6 : 1))
    vars['--segment-active-bg'] = alpha(F, 0.08)
  }
  return vars
}

/** push the choices onto <html>; call again whenever the effective theme flips */
export function applyAppearance(root: HTMLElement, value: Appearance): void {
  const kind = effectiveTheme(root)
  const vars = deriveThemeVars(kind, value[kind], value.contrast)
  for (const name of THEME_VARS) {
    const v = vars[name]
    if (v === undefined) root.style.removeProperty(name)
    else root.style.setProperty(name, v)
  }
  // --wall only exists for dark; light keeps the stylesheet value
  if (kind === 'light') root.style.removeProperty('--wall')
  const font = FONT_STACKS[value.font]
  if (font === null) root.style.removeProperty('--ui-font')
  else root.style.setProperty('--ui-font', font)
  if (value.transparency) root.removeAttribute('data-transparency')
  else root.setAttribute('data-transparency', 'off')
  root.setAttribute('data-contrast', value.contrast)
  root.setAttribute('data-conv-steps', value.conversationSteps)
}

// ── live store: one source of truth for Settings, boot and cross-window sync ──

type Listener = (value: Appearance) => void
let current: Appearance = cloneAppearance(DEFAULT_APPEARANCE)
const listeners = new Set<Listener>()

export function getAppearance(): Appearance {
  return current
}

export function subscribeAppearance(fn: Listener): () => void {
  listeners.add(fn)
  return () => listeners.delete(fn)
}

/** merge a change, persist it, repaint, and notify subscribers */
export function updateAppearance(patch: Partial<Appearance>): Appearance {
  current = { ...current, ...patch }
  saveAppearance(window.localStorage, current)
  applyAppearance(document.documentElement, current)
  for (const fn of listeners) fn(current)
  return current
}

export function resetAppearance(): Appearance {
  return updateAppearance(cloneAppearance(DEFAULT_APPEARANCE))
}

/**
 * Boot: load, paint before first render, and keep painting when the effective
 * theme flips (Settings toggle, main-process broadcast, or the OS scheme in
 * "system" mode). Returns a disposer.
 */
export function startAppearance(root: HTMLElement): () => void {
  current = loadAppearance(window.localStorage)
  applyAppearance(root, current)
  const repaint = (): void => applyAppearance(root, current)
  const observer = new MutationObserver((records) => {
    if (records.some((r) => r.attributeName === 'data-theme')) repaint()
  })
  observer.observe(root, { attributes: true, attributeFilter: ['data-theme'] })
  let mq: MediaQueryList | null = null
  try {
    mq = window.matchMedia('(prefers-color-scheme: dark)')
    mq.addEventListener('change', repaint)
  } catch {
    mq = null
  }
  return () => {
    observer.disconnect()
    mq?.removeEventListener('change', repaint)
  }
}

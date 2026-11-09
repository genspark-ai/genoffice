/** user-adjustable look of the shell; stored per machine, applied before first paint */
export type UiFont = 'system' | 'sans' | 'serif' | 'mono'

export interface Appearance {
  /** null = the theme's built-in accent */
  accent: string | null
  font: UiFont
  /** false = opaque surfaces instead of glass over the window material */
  transparency: boolean
  /** ask before deleting instead of the undo toast */
  confirmDelete: boolean
}

export const DEFAULT_APPEARANCE: Appearance = {
  accent: null,
  font: 'system',
  transparency: true,
  confirmDelete: false,
}

const STORAGE_KEY = 'genoffice.appearance'

export const FONT_STACKS: Record<UiFont, string | null> = {
  system: null,
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
  return contrastRatio(accent, '#ffffff') >= contrastRatio(accent, '#000000')
    ? '#ffffff'
    : '#000000'
}

export function loadAppearance(storage: Pick<Storage, 'getItem'> | undefined): Appearance {
  try {
    const raw = storage?.getItem(STORAGE_KEY)
    if (!raw) return { ...DEFAULT_APPEARANCE }
    const data = JSON.parse(raw) as Partial<Appearance>
    return {
      accent: typeof data.accent === 'string' ? parseHex(data.accent) : null,
      font:
        typeof data.font === 'string' && Object.hasOwn(FONT_STACKS, data.font)
          ? data.font
          : 'system',
      transparency: data.transparency !== false,
      confirmDelete: data.confirmDelete === true,
    }
  } catch {
    return { ...DEFAULT_APPEARANCE }
  }
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

/** push the choices onto <html>; unset values fall back to the stylesheet */
export function applyAppearance(root: HTMLElement, value: Appearance): void {
  const set = (name: string, v: string | null): void => {
    if (v === null) root.style.removeProperty(name)
    else root.style.setProperty(name, v)
  }
  set('--accent', value.accent)
  set('--accent-hover', value.accent)
  set('--accent-text', value.accent ? textOnAccent(value.accent) : null)
  set('--ui-font', FONT_STACKS[value.font])
  if (value.transparency) root.removeAttribute('data-transparency')
  else root.setAttribute('data-transparency', 'off')
}

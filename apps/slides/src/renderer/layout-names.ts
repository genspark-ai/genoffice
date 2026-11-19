import type { StringKey } from './i18n/locale'

/** PowerPoint's canonical layout names → localized labels (the built-in set; unknown names show as-is) */
const LAYOUT_NAME_KEYS: Record<string, StringKey> = {
  'Title Slide': 'ribbonLayoutTitleSlide',
  'Title and Content': 'ribbonLayoutTitleAndContent',
  'Section Header': 'ribbonLayoutSectionHeader',
  'Two Content': 'ribbonLayoutTwoContent',
  'Title Only': 'ribbonLayoutTitleOnly',
  Blank: 'ribbonLayoutBlank',
}

export function layoutLabel(name: string, tf: (key: StringKey) => string): string {
  const key = LAYOUT_NAME_KEYS[name]
  return key ? tf(key) : name
}

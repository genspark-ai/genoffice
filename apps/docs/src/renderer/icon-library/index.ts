import type { StringKey } from '../i18n/locale'

/**
 * The icons of Insert > Icons. The drawings are Fluent UI System Icons (MIT),
 * picked by tools/docs-icon-library/list.mjs and built into icons.json, which
 * loads only when the dialog first opens.
 */
export const ICON_CATEGORIES = [
  'people',
  'business',
  'commerce',
  'communication',
  'technology',
  'analytics',
  'office',
  'education',
  'medical',
  'transport',
  'buildings',
  'nature',
  'food',
  'sports',
  'media',
  'arrows',
  'signs',
  'interface',
  'celebration',
  'home',
  'security',
] as const

export type IconCategory = (typeof ICON_CATEGORIES)[number]

export const ICON_CATEGORY_LABEL: Record<IconCategory, StringKey> = {
  people: 'ribbonIconCatPeople',
  business: 'ribbonIconCatBusiness',
  commerce: 'ribbonIconCatCommerce',
  communication: 'ribbonIconCatCommunication',
  technology: 'ribbonIconCatTechnology',
  analytics: 'ribbonIconCatAnalytics',
  office: 'ribbonIconCatOffice',
  education: 'ribbonIconCatEducation',
  medical: 'ribbonIconCatMedical',
  transport: 'ribbonIconCatTransport',
  buildings: 'ribbonIconCatBuildings',
  nature: 'ribbonIconCatNature',
  food: 'ribbonIconCatFood',
  sports: 'ribbonIconCatSports',
  media: 'ribbonIconCatMedia',
  arrows: 'ribbonIconCatArrows',
  signs: 'ribbonIconCatSigns',
  interface: 'ribbonIconCatInterface',
  celebration: 'ribbonIconCatCelebration',
  home: 'ribbonIconCatHome',
  security: 'ribbonIconCatSecurity',
}

export interface LibraryIcon {
  name: string
  category: IconCategory
  /** English name, shown as the tooltip and written as the picture's alt text */
  label: string
  keywords: string
  /** the `d` of the icon's one path, in a square of ICON_VIEW_BOX units */
  path: string
}

export const ICON_VIEW_BOX = 24

/** Word inserts an icon one inch square */
export const ICON_INSERT_PX = 96

let loading: Promise<LibraryIcon[]> | null = null

export function loadIconLibrary(): Promise<LibraryIcon[]> {
  loading ??= import('./icons.json').then((data) =>
    (data.default.icons as string[][]).map(([name, category, label, keywords, path]) => ({
      name,
      category: category as IconCategory,
      label,
      keywords,
      path,
    })),
  )
  return loading
}

/**
 * Icons whose name or keywords have a word starting with every word of the
 * query. The keywords are English; `categoryNames` are the category names in
 * the language of the UI, where a query word may sit anywhere (a language
 * written without spaces has no word starts).
 */
export function searchIcons(
  icons: LibraryIcon[],
  query: string,
  categoryNames?: Partial<Record<IconCategory, string>>,
): LibraryIcon[] {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (terms.length === 0) return icons
  return icons.filter((icon) => {
    const words = `${icon.label} ${icon.name.replace(/_/g, ' ')} ${icon.keywords}`
      .toLowerCase()
      .split(/[\s-]+/)
    const category = categoryNames?.[icon.category]?.toLowerCase() ?? ''
    return terms.every(
      (term) => words.some((word) => word.startsWith(term)) || category.includes(term),
    )
  })
}

/** the icon as an SVG file: black, as Word's icons are until recoloured */
export function iconSvg(icon: LibraryIcon, sizePx = ICON_INSERT_PX): string {
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${sizePx}" height="${sizePx}" ` +
    `viewBox="0 0 ${ICON_VIEW_BOX} ${ICON_VIEW_BOX}"><path d="${icon.path}"/></svg>`
  )
}

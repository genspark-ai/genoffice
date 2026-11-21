import type { SectionSettings, ThemeColors, ThemeFonts } from '@genoffice/docx-engine'

/** Format > Document > Default…: the page setup new blank documents start with */
export interface PageDefault {
  /** margins as the layout uses them (the gutter folded in), twips */
  section: Pick<
    SectionSettings,
    | 'marginTop'
    | 'marginRight'
    | 'marginBottom'
    | 'marginLeft'
    | 'headerDist'
    | 'footerDist'
    | 'gutter'
    | 'gutterAtTop'
    | 'vAlign'
    | 'columns'
  >
  mirror: boolean
  titlePg: boolean
  evenOdd: boolean
}

/** Format > Font > Default…: the font the default paragraph style of new blank documents carries */
export interface FontDefault {
  latin?: string
  eastAsia?: string
  sizeHalfPoints?: number
  bold?: boolean
  italic?: boolean
  /** hex without '#' */
  color?: string
}

/** Design > Set as Default: the look new blank documents start with */
export interface DesignDefault {
  fonts?: ThemeFonts
  colors?: ThemeColors
  /** the default paragraph style's spacing */
  spacing?: { beforeTwips?: number; afterTwips: number; line: number }
  page?: PageDefault
  font?: FontDefault
}

const KEY = 'aidocs.designDefault'

export function loadDesignDefault(): DesignDefault | null {
  try {
    const raw = localStorage.getItem(KEY)
    const parsed: unknown = raw ? JSON.parse(raw) : null
    return parsed && typeof parsed === 'object' ? (parsed as DesignDefault) : null
  } catch {
    return null
  }
}

function store(value: DesignDefault): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(value))
  } catch {
    /* storage unavailable: the default just does not stick */
  }
}

/** the look replaces the kept one; the kept page setup and font stay */
export function saveDesignDefault(look: Omit<DesignDefault, 'page' | 'font'>): void {
  const kept = loadDesignDefault()
  store({
    ...look,
    ...(kept?.page ? { page: kept.page } : {}),
    ...(kept?.font ? { font: kept.font } : {}),
  })
}

/** the page setup replaces the kept one; the kept look stays */
export function savePageDefault(page: PageDefault): void {
  store({ ...loadDesignDefault(), page })
}

/** the font replaces the kept one; the kept look and page setup stay */
export function saveFontDefault(font: FontDefault): void {
  store({ ...loadDesignDefault(), font })
}

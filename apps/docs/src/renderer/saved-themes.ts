import { themeOfXml, type ThemeFileTheme } from '@genoffice/docx-engine'
import type { SavedThemeEntry } from '../shared/ipc'
import type { ThemePick } from './commands/menu-commands'

/** a theme file of the themes folder, as the Themes gallery lists it under Custom */
export interface SavedTheme {
  /** file name inside the themes folder */
  file: string
  /** the file's name without its extension: what the card is called */
  name: string
  theme: ThemeFileTheme
}

export const themePickOf = ({ fonts, colors, format }: ThemeFileTheme): ThemePick => ({
  fonts,
  colors,
  format,
})

/** the files that hold a theme; one that does not is left out */
export const savedThemesOf = (entries: readonly SavedThemeEntry[]): SavedTheme[] =>
  entries.flatMap(({ name, xml }) => {
    const theme = themeOfXml(xml)
    return theme ? [{ file: name, name: name.replace(/\.thmx$/i, ''), theme }] : []
  })

export const loadSavedThemes = (): Promise<SavedTheme[]> =>
  window.desktop
    .listThemes()
    .then(savedThemesOf)
    .catch((): SavedTheme[] => [])

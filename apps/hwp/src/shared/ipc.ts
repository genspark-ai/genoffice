import type { Lang } from '@genoffice/i18n'

export const HWP_CHANNELS = {
  consumePending: 'hwp:consume-pending',
  fileRenamed: 'hwp:file-renamed',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
} as const

export type UiTheme = 'light' | 'dark' | 'system'

export interface HwpApi {
  consumePending(): Promise<string | null>
  onFileRenamed(handler: (newPath: string) => void): () => void
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  onChromePressed(handler: () => void): () => void
}

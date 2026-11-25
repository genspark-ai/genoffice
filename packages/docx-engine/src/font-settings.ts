import { loadDocxZip } from './zip-load'
import { parseStyles } from './parse-styles'
import {
  mergeDefaultFontsXml,
  upsertStyleXml,
  type DefaultFonts,
  type StyleUpsert,
} from './style-upsert'
import {
  THEME_PART_PATH,
  editedThemeXml,
  readThemeColors,
  readThemeFonts,
  readThemeFormat,
  type ThemeEdit,
  type ThemeFormat,
} from './theme'
import type { ParsedDoc, ThemeColors, ThemeFonts } from './types'

export type { ThemeEdit }

/** the theme a save of `edit` would leave in the file, read back the way a parse reads it */
export function previewTheme(
  themeXml: string | undefined,
  parsed: Pick<ParsedDoc, 'themeFonts' | 'themeColors'>,
  edit: ThemeEdit,
): { themeFonts: ThemeFonts | null; themeColors: ThemeColors | null } {
  if (!edit.fonts && !edit.colors) {
    return { themeFonts: parsed.themeFonts ?? null, themeColors: parsed.themeColors ?? null }
  }
  const xml = editedThemeXml(themeXml, edit)
  const themeFonts = readThemeFonts(xml)
  const { eaLang, bidiLang } = parsed.themeFonts ?? {}
  if (themeFonts && eaLang) themeFonts.eaLang = eaLang
  if (themeFonts && bidiLang) themeFonts.bidiLang = bidiLang
  return { themeFonts, themeColors: readThemeColors(xml) }
}

/**
 * The document's bytes with `edit` written into its theme part and nothing else touched:
 * parsing them resolves every theme reference in the body the way the saved file will.
 */
export async function themedDocxBytes(bytes: Uint8Array, edit: ThemeEdit): Promise<Uint8Array> {
  const zip = await loadDocxZip(bytes)
  zip.file(THEME_PART_PATH, editedThemeXml(await zip.file(THEME_PART_PATH)?.async('string'), edit))
  return zip.generateAsync({ type: 'uint8array', compression: 'STORE' })
}

/** the theme name and format scheme the file has; null without a theme part */
export async function readDocxThemeFormat(bytes: Uint8Array): Promise<ThemeFormat | null> {
  const xml = await (await loadDocxZip(bytes)).file(THEME_PART_PATH)?.async('string')
  return xml === undefined ? null : readThemeFormat(xml)
}

/** Resolve live style inheritance without reparsing the document body or its media. */
export async function previewFontSettings(
  parsed: ParsedDoc,
  upserts: StyleUpsert[],
  defaults?: DefaultFonts,
  theme?: ThemeEdit,
) {
  const zip = await loadDocxZip(parsed.internal.originalBytes)
  let xml =
    (await zip.file('word/styles.xml')?.async('string')) ??
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
  for (const up of upserts) xml = upsertStyleXml(xml, up)
  if (defaults) xml = mergeDefaultFontsXml(xml, defaults)
  zip.file('word/styles.xml', xml)
  const live = previewTheme(await zip.file(THEME_PART_PATH)?.async('string'), parsed, theme ?? {})
  const resolved = await parseStyles(zip, live.themeColors, live.themeFonts)
  return { ...resolved, ...live }
}

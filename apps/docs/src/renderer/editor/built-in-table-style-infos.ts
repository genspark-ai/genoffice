import { builtInTableStyleInfos, type StyleInfo, type ThemeColors } from '@genoffice/docx-engine'

/** Word's built-in table styles read once per theme (the gallery previews and the apply command share it) */
let cached: { theme: ThemeColors | null; infos: Promise<Map<string, StyleInfo>> } | null = null

export function loadBuiltInTableStyleInfos(
  theme: ThemeColors | null,
): Promise<Map<string, StyleInfo>> {
  if (!cached || cached.theme !== theme) cached = { theme, infos: builtInTableStyleInfos(theme) }
  return cached.infos
}

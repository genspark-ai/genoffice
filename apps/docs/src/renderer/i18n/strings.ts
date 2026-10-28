import type { Lang } from '@genoffice/i18n'
import { aiEn, loadAiStrings, type AiDict } from './strings-ai'
import { appEn, loadAppStrings, type AppDict } from './strings-app'
import { editorStrings } from './strings-editor'
import { ribbonEn, loadRibbonStrings, type RibbonDict } from './strings-ribbon'
import { tableStrings } from './strings-table'
import { zoteroStrings } from './strings-zotero'

/** merged per-locale dictionary; zh still anchors the key set (see the shard types) */
export type StringsDict = AppDict &
  RibbonDict &
  typeof tableStrings.zh &
  typeof editorStrings.zh &
  AiDict &
  typeof zoteroStrings.zh

/** eager en fallback, so t() is meaningful before any lazy locale chunk lands */
export const enStrings: StringsDict = {
  ...appEn,
  ...ribbonEn,
  ...tableStrings.en,
  ...editorStrings.en,
  ...aiEn,
  ...zoteroStrings.en,
}

const dicts: Partial<Record<Lang, StringsDict>> = { en: enStrings }

/** sync registry read for the translator; undefined until loadStrings(lang) resolves */
export function getStrings(lang: Lang): StringsDict | undefined {
  return dicts[lang]
}

/** merge one locale's shards into the registry; en resolves immediately, others lazy-load */
export async function loadStrings(lang: Lang): Promise<void> {
  if (dicts[lang]) return
  const [app, ai, ribbon] = await Promise.all([
    loadAppStrings(lang),
    loadAiStrings(lang),
    loadRibbonStrings(lang),
  ])
  dicts[lang] = {
    ...app,
    ...ribbon,
    ...tableStrings[lang],
    ...editorStrings[lang],
    ...ai,
    ...zoteroStrings[lang],
  }
}

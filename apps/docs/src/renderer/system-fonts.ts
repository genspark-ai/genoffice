import { useCallback, useState } from 'react'
import { setFontAliasLookup } from './line-metrics'

let cached: readonly string[] | null = null
let pending: Promise<readonly string[]> | null = null

interface LocalFontFace {
  readonly family: string
  readonly fullName: string
  readonly postscriptName: string
  readonly style?: string
}

/** localized-name → English family name, from the system font table.
    A font like 方正仿宋_GBK registers its Chinese full name alongside the
    English family FZFangSong-Z02; Chromium's font matching on Windows can
    miss the localized name, so the renderer substitutes the English family
    it maps to (genoffice#2449). */
let aliasCache: ReadonlyMap<string, string> | null = null

export function fontAliasOf(name: string): string | null {
  return aliasCache?.get(name) ?? null
}

async function queryFamilies(): Promise<readonly string[]> {
  const query = (window as { queryLocalFonts?: () => Promise<readonly LocalFontFace[]> })
    .queryLocalFonts
  if (!query) return []
  try {
    const fonts = await query.call(window)
    const families = new Set<string>()
    const aliases = new Map<string, string>()
    for (const font of fonts) {
      if (font.family) families.add(font.family)
      // a full name that differs from its family is a localized alias
      if (font.fullName && font.family && font.fullName !== font.family) {
        aliases.set(font.fullName, font.family)
      }
      if (font.postscriptName && font.family && font.postscriptName !== font.family) {
        aliases.set(font.postscriptName, font.family)
      }
    }
    aliasCache = aliases
    setFontAliasLookup((name) => aliases.get(name) ?? null)
    return [...families].sort((a, b) => a.localeCompare(b))
  } catch {
    return []
  }
}

function loadSystemFontFamilies(): Promise<readonly string[]> {
  if (cached) return Promise.resolve(cached)
  pending ??= queryFamilies().then((families) => {
    cached = families
    return families
  })
  return pending
}

/// Empty until load() runs — call it from the picker's open click so the
/// Local Font Access API sees user activation; cached for the page lifetime,
/// and on failure the pickers just keep the built-in list. Returns every
/// family: the pickers dedupe the candidate names where the two sections
/// merge (docs keeps every candidate visible, so it never partitions).
/** Eagerly populate the alias map (and the family list): called at document
    open so the localized-name lookup is ready before the first render. The
    query needs user activation — the file-open click provides it. */
export function loadSystemFontAliases(): Promise<void> {
  return loadSystemFontFamilies().then(() => undefined)
}

export function useSystemFontFamilies(): {
  readonly families: readonly string[]
  readonly load: () => void
} {
  const [families, setFamilies] = useState<readonly string[]>(cached ?? [])
  const load = useCallback(() => {
    void loadSystemFontFamilies().then(setFamilies)
  }, [])
  return { families, load }
}

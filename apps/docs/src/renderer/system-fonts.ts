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

/** A face whose full name is its family plus a style suffix ("PingFang SC
    Semibold", "Helvetica Neue Light", "Arial Bold"). The browser resolves
    those names itself, so aliasing them to their family would drop the
    weight a document declared (genoffice#2449). */
function isWeightQualified(font: LocalFontFace): boolean {
  return !!font.style && font.fullName === `${font.family} ${font.style}`
}

async function queryFamilies(): Promise<readonly string[] | null> {
  const query = (window as { queryLocalFonts?: () => Promise<readonly LocalFontFace[]> })
    .queryLocalFonts
  if (!query) return []
  try {
    const fonts = await query.call(window)
    const families = new Set<string>()
    // localized-name -> English family name, from the system font table.
    // A font like \u65b9\u6b63\u4eff\u5b8b_GBK registers its Chinese full name alongside
    // the English family FZFangSong-Z02; Chromium's font matching on
    // Windows can miss the localized name, so the renderer substitutes the
    // English family it maps to (genoffice#2449).
    const aliases = new Map<string, string>()
    for (const font of fonts) {
      if (font.family) families.add(font.family)
      // a full name that differs from its family is a localized alias, unless
      // it only carries a weight suffix the browser already resolves
      if (
        font.fullName &&
        font.family &&
        font.fullName !== font.family &&
        !isWeightQualified(font)
      ) {
        aliases.set(font.fullName, font.family)
      }
      if (font.postscriptName && font.family && font.postscriptName !== font.family) {
        aliases.set(font.postscriptName, font.family)
      }
    }
    setFontAliasLookup((name) => aliases.get(name) ?? null)
    return [...families].sort((a, b) => a.localeCompare(b))
  } catch {
    // the query needs transient user activation; a document open may not have
    // any, so the query rejects. A rejected query must not be cached: the
    // pickers would keep an empty system-font list for the rest of the
    // page's life, and the alias map would stay empty (genoffice#2449).
    return null
  }
}

function loadSystemFontFamilies(): Promise<readonly string[]> {
  if (cached) return Promise.resolve(cached)
  pending ??= queryFamilies().then((families) => {
    if (families) {
      cached = families
    } else {
      // the query was rejected: drop the pending promise so a later query
      // (the font picker's own click) can run and fill the cache
      pending = null
    }
    return families ?? []
  })
  return pending
}

/// Empty until load() runs — call it from the picker's open click so the
/// Local Font Access API sees user activation; a successful query is cached
/// for the page lifetime, a rejected one is not cached at all, so the pickers
/// keep the built-in list until a query succeeds. Returns every family: the
/// pickers dedupe the candidate names where the two sections merge (docs
/// keeps every candidate visible, so it never partitions).
/** Eagerly populate the alias map (and the family list): the query needs
    transient user activation, which a document open often lacks, so it never
    blocks the open — a rejected query is not cached and the next query that
    does succeed (the font picker's own click) fills the cache instead
    (genoffice#2449). */
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

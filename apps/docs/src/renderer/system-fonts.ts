import { useCallback, useState } from 'react'
import { setFontAliasLookup } from './line-metrics'

let cached: readonly string[] | null = null
let pending: Promise<readonly string[]> | null = null

/** declared family name -> the other spellings of the face behind it, the
 *  canonical (en-US) one first. The map is filled from the main process, which
 *  reads the installed faces' name tables (genoffice#2449). */
const aliasMap = new Map<string, readonly string[]>()
const aliasListeners = new Set<(names: readonly string[]) => void>()

interface LocalFontFace {
  readonly family: string
}

/** the key the alias map is indexed by: a declared name, without its quotes */
const aliasKey = (name: string): string =>
  name
    .trim()
    .replace(/^["']|["']$/g, '')
    .toLowerCase()

// the lookup reads the live map, so an alias applies the moment it lands
setFontAliasLookup((name) => aliasMap.get(aliasKey(name)) ?? null)

/**
 * Ask the main process for the other spellings of the installed faces behind
 * these names. The answer comes over IPC, so it lands after the first layout:
 * the subscribers re-style the document when it does (genoffice#2449).
 */
export async function resolveFontAliases(names: readonly string[]): Promise<void> {
  if (!names.length) return
  let answered: Record<string, readonly string[]>
  try {
    answered = (await window.desktop.fontAliases(names)) ?? {}
  } catch {
    return
  }
  const added: string[] = []
  for (const [name, aliases] of Object.entries(answered)) {
    if (!aliases?.length) continue
    const known = aliasMap.get(aliasKey(name))
    if (known && known.length === aliases.length && known.every((a, i) => a === aliases[i]))
      continue
    aliasMap.set(aliasKey(name), aliases)
    added.push(name)
  }
  if (!added.length) return
  for (const listener of [...aliasListeners]) listener(added)
}

/**
 * Subscribe to the alias map landing. The callback gets the names that just
 * gained an alias, so a document that declares none of them is not reflowed.
 */
export function onFontAliasesChanged(cb: (names: readonly string[]) => void): () => void {
  aliasListeners.add(cb)
  return () => {
    aliasListeners.delete(cb)
  }
}

async function queryFamilies(): Promise<readonly string[] | null> {
  const query = (window as { queryLocalFonts?: () => Promise<readonly LocalFontFace[]> })
    .queryLocalFonts
  if (!query) return []
  try {
    const fonts = await query.call(window)
    const families = new Set<string>()
    for (const font of fonts) if (font.family) families.add(font.family)
    return [...families].sort((a, b) => a.localeCompare(b))
  } catch {
    // the query needs transient user activation; a document open may not have
    // any, so the query rejects. A rejected query must not be cached: the
    // pickers would keep an empty system-font list for the rest of the
    // page's life (genoffice#2449).
    return null
  }
}

/// Empty until load() runs — call it from the picker's open click so the
// Local Font Access API sees user activation; a successful query is cached
// for the page lifetime, a rejected one is not cached at all, so the pickers
// keep the built-in list until a query succeeds. Returns every family: the
// pickers dedupe the candidate names where the two sections merge (docs
// keeps every candidate visible, so it never partitions).
export function loadSystemFontFamilies(): Promise<readonly string[]> {
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

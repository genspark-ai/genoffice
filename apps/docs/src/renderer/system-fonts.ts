import { useCallback, useState } from 'react'

import { BUILTIN_FONT_FAMILIES } from './font-list'

let cached: readonly string[] | null = null
let pending: Promise<readonly string[]> | null = null

async function queryFamilies(): Promise<readonly string[]> {
  const query = (window as { queryLocalFonts?: () => Promise<{ readonly family: string }[]> })
    .queryLocalFonts
  if (!query) return []
  try {
    const fonts = await query.call(window)
    const families = new Set<string>()
    for (const font of fonts) {
      if (font.family && !BUILTIN_FONT_FAMILIES.includes(font.family)) families.add(font.family)
    }
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
/// and on failure the pickers just keep the built-in list.
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

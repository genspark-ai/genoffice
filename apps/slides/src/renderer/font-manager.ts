import { useCallback, useEffect, useState } from 'react'

export interface CatalogEntry {
  family: string
  script: 'latin' | 'ja' | 'ko' | 'sc' | 'tc'
  installed: boolean
  downloading: boolean
}

let cached: CatalogEntry[] | null = null

/**
 * Downloadable font catalog + install actions. Loaded lazily from the picker's
 * open click (same pattern as useSystemFontFamilies); download/install push new
 * layouts from main via deck-changed, so callers only refresh list state here.
 */
export function useFontCatalog(): {
  readonly catalog: CatalogEntry[]
  readonly busy: ReadonlySet<string>
  readonly failed: ReadonlySet<string>
  readonly load: () => void
  readonly download: (family: string) => Promise<boolean>
  readonly installLocal: () => Promise<string[]>
} {
  const [catalog, setCatalog] = useState<CatalogEntry[]>(cached ?? [])
  const [busy, setBusy] = useState<ReadonlySet<string>>(new Set())
  const [failed, setFailed] = useState<ReadonlySet<string>>(new Set())

  const load = useCallback(() => {
    void window.slidesApi
      .fontCatalog?.()
      .then((c) => {
        cached = c
        setCatalog(c)
      })
      .catch(() => {})
  }, [])

  const download = useCallback(
    async (family: string): Promise<boolean> => {
      setBusy((s) => new Set(s).add(family))
      setFailed((s) => {
        const n = new Set(s)
        n.delete(family)
        return n
      })
      try {
        const r = await window.slidesApi.fontDownload?.(family)
        if (!r?.ok) throw new Error(r?.error)
        return true
      } catch {
        setFailed((s) => new Set(s).add(family))
        return false
      } finally {
        setBusy((s) => {
          const n = new Set(s)
          n.delete(family)
          return n
        })
        load()
      }
    },
    [load],
  )

  const installLocal = useCallback(async (): Promise<string[]> => {
    try {
      const r = await window.slidesApi.fontInstallLocal?.()
      return r?.families ?? []
    } catch {
      return []
    } finally {
      load()
    }
  }, [load])

  useEffect(() => window.slidesApi.onFontsChanged?.(load), [load])

  return { catalog, busy, failed, load, download, installLocal }
}

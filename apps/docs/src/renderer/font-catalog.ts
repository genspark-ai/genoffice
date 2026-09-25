import { useCallback, useState } from 'react'

import type { FontScript } from '@genoffice/electron-utils/font-catalog'
import { registerStoreFonts } from './store-fonts'

export interface CatalogEntry {
  family: string
  script: FontScript
  installed: boolean
  downloading: boolean
}

let cached: CatalogEntry[] | null = null

/**
 * Downloadable font catalog + install actions for docs. Loaded lazily from the
 * picker's open click (same pattern as useSystemFontFamilies); downloads and
 * local imports land in <userData>/fonts and register as FontFaces, so the new
 * families are immediately selectable and render in the editor.
 */
export function useFontCatalog(): {
  catalog: CatalogEntry[]
  busy: ReadonlySet<string>
  failed: ReadonlySet<string>
  load: () => void
  download: (family: string) => Promise<boolean>
  installLocal: () => Promise<string[]>
} {
  const [catalog, setCatalog] = useState<CatalogEntry[]>(cached ?? [])
  const [busy, setBusy] = useState<ReadonlySet<string>>(new Set())
  const [failed, setFailed] = useState<ReadonlySet<string>>(new Set())

  const load = useCallback(() => {
    void window.desktop
      .fontCatalog()
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
        const r = await window.desktop.fontDownload(family)
        if (!r?.ok) throw new Error(r?.error)
        await registerStoreFonts()
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
      const r = await window.desktop.fontInstallLocal()
      const families = r?.families ?? []
      if (families.length) await registerStoreFonts()
      return families
    } catch {
      return []
    } finally {
      load()
    }
  }, [load])

  return { catalog, busy, failed, load, download, installLocal }
}

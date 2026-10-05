/**
 * The downloadable-catalog half of a font picker, shared by every app that
 * offers one.
 *
 * ## What this decides, and what it refuses to
 *
 * A family is tens of megabytes — Noto Serif SC alone is 28 MiB — so a bare
 * selection that quietly starts a download is the failure mode worth designing
 * against. This module therefore **never downloads on its own**: it hands the
 * caller the list *with its size* and waits to be asked. Whether a download
 * needs a click, a confirmation dialog, or is fine inline is the host app's
 * call, and it should be a decision someone can review rather than a default
 * inherited from here.
 *
 * The three pieces of copy this needs already exist in the slides app, so a
 * host copies those across rather than inventing new words; the size itself
 * needs no translation, being a number and a unit symbol.
 */
import { useCallback, useEffect, useState } from 'react'

/** One offerable family, as the main process reports it. */
export interface FontCatalogRow {
  readonly family: string
  readonly script: 'latin' | 'ja' | 'ko' | 'sc' | 'tc'
  readonly license: string
  readonly installed: boolean
  readonly bytes: number
}

/** What a host has to supply; `catalog` and `download` are the two IPC calls. */
export interface FontCatalogApi {
  fontCatalog?: () => Promise<readonly FontCatalogRow[]>
  fontDownload?: (family: string) => Promise<{ ok: boolean; error?: string }>
}

const KIB = 1024
const MIB = 1024 * 1024

/**
 * A download size the way a person reads it.
 *
 * Unit symbols are not language, so this needs no translation and a picker in
 * any locale shows the same number the reader will be billed for. `MiB` is
 * spelled out rather than rendered as "MB": the catalog counts bytes, and a
 * 28 MiB download rendered as "28 MB" is a promise about disk that does not
 * match what gets written.
 */
export function formatFontBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return ''
  if (bytes < KIB) return `${bytes} B`
  if (bytes < MIB) return `${Math.round(bytes / KIB)} KiB`
  const mib = bytes / MIB
  // One decimal under 10 MiB so a 4.2 MiB family is not "4 MiB", whole above
  return `${mib < 10 ? mib.toFixed(1) : Math.round(mib)} MiB`
}

export interface UseFontCatalog {
  readonly rows: readonly FontCatalogRow[]
  readonly loading: boolean
  readonly busy: ReadonlySet<string>
  readonly failed: ReadonlySet<string>
  readonly reload: () => void
  /** ask for a family; resolves to whether it landed */
  readonly download: (family: string) => Promise<boolean>
}

/**
 * Load the catalog and track what is being fetched.
 *
 * `reload` is the call a host makes after a download finishes: the main process
 * registers the new family, so the row's `installed` flag only changes there.
 */
export function useFontCatalog(api: FontCatalogApi): UseFontCatalog {
  const [rows, setRows] = useState<readonly FontCatalogRow[]>([])
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState<ReadonlySet<string>>(new Set())
  const [failed, setFailed] = useState<ReadonlySet<string>>(new Set())
  const [nonce, setNonce] = useState(0)

  useEffect(() => {
    let cancelled = false
    const list = api.fontCatalog
    if (!list) {
      // No download surface in this build: showing nothing beats showing a
      // row whose button cannot work.
      setRows([])
      setLoading(false)
      return
    }
    setLoading(true)
    void list()
      .then((next) => {
        if (!cancelled) setRows([...next])
      })
      .catch(() => {
        if (!cancelled) setRows([])
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [api, nonce])

  const download = useCallback(
    async (family: string): Promise<boolean> => {
      const send = api.fontDownload
      if (!send) return false
      setBusy((previous) => new Set(previous).add(family))
      setFailed((previous) => {
        const next = new Set(previous)
        next.delete(family)
        return next
      })
      try {
        const result = await send(family)
        if (!result?.ok) throw new Error(result?.error ?? 'download failed')
        // the main process registered it; the row's installed flag moved there
        setNonce((n) => n + 1)
        return true
      } catch {
        setFailed((previous) => new Set(previous).add(family))
        return false
      } finally {
        setBusy((previous) => {
          const next = new Set(previous)
          next.delete(family)
          return next
        })
      }
    },
    [api],
  )

  const reload = useCallback(() => setNonce((n) => n + 1), [])

  return { rows, loading, busy, failed, reload, download }
}

/** The families worth offering: published, not already installed. */
export function offerableRows(rows: readonly FontCatalogRow[]): readonly FontCatalogRow[] {
  return rows.filter((row) => !row.installed)
}

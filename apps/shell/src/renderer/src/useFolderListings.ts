import { useCallback, useRef, useState } from 'react'
import type { FolderListing } from '../../shared/home-api'

/**
 * Lazily loaded folder listings keyed by directory. Listing a folder that is
 * already cached is a no-op; `invalidate` drops entries so the next render
 * refetches them (the main process reports changed directories via watch).
 */
export function useFolderListings() {
  const [listings, setListings] = useState<ReadonlyMap<string, FolderListing>>(new Map())
  // dir → whether a reload was requested while its request was in flight (the
  // in-flight answer may predate the change, so it is fetched once more)
  const inflight = useRef(new Map<string, boolean>())

  const load = useCallback((dir: string, force = false) => {
    if (inflight.current.has(dir)) {
      if (force) inflight.current.set(dir, true)
      return
    }
    inflight.current.set(dir, false)
    void window.aiOffice
      .listFolder(dir)
      .then((listing) => {
        setListings((prev) => {
          const next = new Map(prev)
          next.set(dir, listing)
          return next
        })
      })
      .finally(() => {
        const again = inflight.current.get(dir)
        inflight.current.delete(dir)
        if (again) load(dir, true)
      })
  }, [])

  const invalidate = useCallback(
    (dirs: readonly string[]) => {
      setListings((prev) => {
        let changed = false
        const next = new Map(prev)
        for (const dir of dirs) {
          if (next.delete(dir)) changed = true
        }
        return changed ? next : prev
      })
      for (const dir of dirs) load(dir, true)
    },
    [load],
  )

  const reset = useCallback(() => setListings(new Map()), [])

  /** shown or currently loading: the folders a watch event should refresh */
  const tracked = useCallback(
    (dir: string) => listings.has(dir) || inflight.current.has(dir),
    [listings],
  )

  return { listings, load, invalidate, reset, tracked }
}

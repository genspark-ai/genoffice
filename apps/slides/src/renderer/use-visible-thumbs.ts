import { useEffect, useState, type RefObject } from 'react'

/**
 * Indices of thumbnail items near the scroll viewport. Every thumbnail is a live Konva canvas,
 * so a 409-slide deck used to hold 411 canvases (~9 GB); items outside `rootMargin` render a
 * same-size placeholder instead and get their canvas back when scrolled near.
 */
export function useVisibleThumbs(
  rootRef: RefObject<HTMLElement | null>,
  itemSelector: string,
  deps: readonly unknown[],
  initial = 24,
): ReadonlySet<number> {
  const [visible, setVisible] = useState<ReadonlySet<number>>(
    () => new Set(Array.from({ length: initial }, (_, i) => i)),
  )
  useEffect(() => {
    const root = rootRef.current
    if (!root || typeof IntersectionObserver === 'undefined') return
    // The first callback after (re)observing reports every item, so it rebuilds the set from
    // scratch instead of inheriting indices from a previous DOM (collapsed section, other deck)
    let fresh = true
    const io = new IntersectionObserver(
      (entries) => {
        setVisible((prev) => {
          const next = fresh ? new Set<number>() : new Set(prev)
          fresh = false
          for (const e of entries) {
            const i = Number((e.target as HTMLElement).dataset.index)
            if (!Number.isFinite(i)) continue
            if (e.isIntersecting) next.add(i)
            else next.delete(i)
          }
          return next
        })
      },
      { root, rootMargin: '150% 0px' },
    )
    for (const el of root.querySelectorAll(itemSelector)) io.observe(el)
    return () => io.disconnect()
  }, deps)
  return visible
}

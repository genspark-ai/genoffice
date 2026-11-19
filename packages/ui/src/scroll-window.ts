import { useCallback, useLayoutEffect, useRef, useState, type UIEvent } from 'react'

const FALLBACK_ROW_PX = 26
const FALLBACK_VIEW_PX = 320

/// Row-index window of a scrolling list: only rows near the viewport should
/// carry a per-row font preview (Chromium loads every previewed face during
/// layout, which froze the picker on machines with hundreds of fonts). Rows
/// stay mounted, so keyboard navigation and scrollIntoView keep working; the
/// estimate uses the first button's height and the overscan absorbs group labels.
export function useScrollWindow(active: boolean, overscan = 8) {
  const ref = useRef<HTMLDivElement | null>(null)
  const frame = useRef(0)
  const [range, setRange] = useState<readonly [number, number]>([0, overscan * 2])
  const measure = useCallback(() => {
    const el = ref.current
    if (!el) return
    const rowH = el.querySelector<HTMLElement>('button')?.offsetHeight || FALLBACK_ROW_PX
    const viewH = el.clientHeight || FALLBACK_VIEW_PX
    const from = Math.max(0, Math.floor(el.scrollTop / rowH) - overscan)
    const to = Math.ceil((el.scrollTop + viewH) / rowH) + overscan
    setRange((r) => (r[0] === from && r[1] === to ? r : [from, to]))
  }, [overscan])
  useLayoutEffect(() => {
    if (active) measure()
    return () => cancelAnimationFrame(frame.current)
  }, [active, measure])
  const onScroll = useCallback(
    (_e: UIEvent<HTMLElement>) => {
      cancelAnimationFrame(frame.current)
      frame.current = requestAnimationFrame(measure)
    },
    [measure],
  )
  const inWindow = useCallback((i: number) => i >= range[0] && i <= range[1], [range])
  return { ref, onScroll, inWindow }
}

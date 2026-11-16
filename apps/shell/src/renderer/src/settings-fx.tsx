import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'

/**
 * Motion helpers for the settings dialog. The patterns follow React Bits
 * (CountUp, FadeContent) and 21st.dev (animated tabs / segmented control),
 * rebuilt here on the theme tokens with no extra dependencies. Styles live in
 * settings-fx.css; everything degrades to static when motion is reduced.
 */

/** true when motion should be skipped (also where matchMedia is missing, e.g. tests) */
export function prefersReducedMotion(): boolean {
  if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') return true
  return window.matchMedia('(prefers-reduced-motion: reduce)').matches
}

/**
 * Sliding highlight: returns a ref for an absolutely positioned indicator that
 * sits over the container's child matching `activeSelector` and glides to the
 * next one when `activeKey` changes. The container must be position: relative.
 */
export function useSlidingIndicator(
  container: RefObject<HTMLElement | null>,
  activeSelector: string,
  activeKey: unknown,
): RefObject<HTMLSpanElement | null> {
  const indicator = useRef<HTMLSpanElement | null>(null)
  useLayoutEffect(() => {
    const box = container.current
    const pill = indicator.current
    if (!box || !pill) return
    const place = (): void => {
      const active = box.querySelector<HTMLElement>(activeSelector)
      if (!active) {
        pill.style.opacity = '0'
        return
      }
      pill.style.opacity = '1'
      pill.style.transform = `translate(${active.offsetLeft}px, ${active.offsetTop}px)`
      pill.style.width = `${active.offsetWidth}px`
      pill.style.height = `${active.offsetHeight}px`
    }
    place()
    // first placement is instant; only later moves animate
    const raf =
      typeof requestAnimationFrame === 'function'
        ? requestAnimationFrame(() => pill.setAttribute('data-ready', ''))
        : 0
    const ro = typeof ResizeObserver === 'function' ? new ResizeObserver(place) : null
    ro?.observe(box)
    return () => {
      if (raf && typeof cancelAnimationFrame === 'function') cancelAnimationFrame(raf)
      ro?.disconnect()
    }
  }, [container, activeSelector, activeKey])
  return indicator
}

/** number that counts up to `value` (ease-out cubic), static when motion is reduced */
export function CountUp({
  value,
  format,
  durationMs = 700,
}: {
  value: number
  format: (n: number) => string
  durationMs?: number
}) {
  const [shown, setShown] = useState(() => (prefersReducedMotion() ? value : 0))
  useEffect(() => {
    if (prefersReducedMotion() || typeof requestAnimationFrame !== 'function') {
      setShown(value)
      return
    }
    let raf = 0
    const start = performance.now()
    const tick = (now: number): void => {
      const p = Math.min(1, (now - start) / durationMs)
      setShown(value * (1 - Math.pow(1 - p, 3)))
      if (p < 1) raf = requestAnimationFrame(tick)
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [value, durationMs])
  return <span className="fx-count">{format(shown)}</span>
}

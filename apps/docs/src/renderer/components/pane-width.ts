import { useState, type PointerEvent as ReactPointerEvent } from 'react'

/** Resizable side-pane width persisted in localStorage; drag the returned handler's element to resize */
export function usePaneWidth(key: string, fallback: number, min: number, max: number) {
  const [width, setWidth] = useState(() => {
    const n = Number(localStorage.getItem(key))
    return n >= min && n <= max ? n : fallback
  })
  const onResizeStart = (e: ReactPointerEvent<HTMLDivElement>) => {
    const startX = e.clientX
    const startW = width
    const el = e.currentTarget
    el.setPointerCapture(e.pointerId)
    let next = startW
    const move = (ev: PointerEvent) => {
      next = Math.min(max, Math.max(min, startW + ev.clientX - startX))
      setWidth(next)
    }
    const up = () => {
      el.removeEventListener('pointermove', move)
      el.removeEventListener('pointerup', up)
      localStorage.setItem(key, String(next))
    }
    el.addEventListener('pointermove', move)
    el.addEventListener('pointerup', up)
  }
  return { width, onResizeStart }
}

export type LeftPane = 'nav' | null

const LEFT_PANE_KEY = 'aidocs.leftPane'

/** whether the Navigation pane was open last time (legacy aidocs.showNav='1') */
export function loadLeftPane(): LeftPane {
  const v = localStorage.getItem(LEFT_PANE_KEY)
  if (v === 'nav') return v
  if (v === null && localStorage.getItem('aidocs.showNav') === '1') return 'nav'
  return null
}

export function saveLeftPane(pane: LeftPane): void {
  localStorage.setItem(LEFT_PANE_KEY, pane ?? '')
}

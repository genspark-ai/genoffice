/**
 * Office "Collapse the Ribbon": the tab row stays, the command band hides.
 * While collapsed, pressing a tab peeks the band as an overlay above the
 * document; a press elsewhere (or Escape / window blur / the shell tab strip)
 * hides it again. Double-clicking a tab and Ctrl+F1 (⌥⌘R on macOS) toggle the
 * collapsed state, which persists per app in localStorage.
 *
 * Markup contract: the ribbon root carries `rootRef` + `rootClass`, the band
 * element carries `data-ribbon-body`, and `RibbonCollapseButton` renders as a
 * sibling of the band (ribbon-collapse.css anchors it to the band's corner).
 */
import { useCallback, useEffect, useRef, useState, type MouseEvent, type RefObject } from 'react'
import { subscribeChromePressed } from './popover-dismiss'

const IS_MAC = typeof navigator !== 'undefined' && navigator.platform.toLowerCase().includes('mac')

export const RIBBON_TOGGLE_SHORTCUT = IS_MAC ? '⌥⌘R' : 'Ctrl+F1'

export function readRibbonCollapsed(storageKey: string): boolean {
  try {
    return localStorage.getItem(storageKey) === '1'
  } catch {
    return false
  }
}

export function writeRibbonCollapsed(storageKey: string, collapsed: boolean): void {
  try {
    localStorage.setItem(storageKey, collapsed ? '1' : '0')
  } catch {
    /* private mode / quota: the toggle still works for this session */
  }
}

/** Ctrl+F1 (Office on Windows) or ⌥⌘R (Office for Mac); key-repeat while held does not re-toggle. */
export function isRibbonToggleShortcut(e: KeyboardEvent): boolean {
  if (e.repeat) return false
  if (e.key === 'F1' && e.ctrlKey && !e.metaKey && !e.altKey) return true
  return IS_MAC && e.metaKey && e.altKey && !e.ctrlKey && !e.shiftKey && e.code === 'KeyR'
}

const POPOVER_OPEN_CLASS = 'genoffice-popover-open'

/**
 * Hide the peeked band on a press outside the ribbon. Popovers that portal
 * out of the ribbon DOM (and presses that only dismiss an open popover) are
 * told apart via the html-level open-popover class maintained by
 * popover-dismiss.ts: with a popover open the decision is deferred a tick —
 * a target unmounted with its popover was a press inside it, and a popover
 * still open afterwards means the press landed inside that popover.
 */
export function installRibbonPeekDismiss(
  root: () => Element | null,
  close: () => void,
): () => void {
  const onPress = (e: Event) => {
    const target = e.target as Node | null
    const el = root()
    if (!target || !el || el.contains(target)) return
    if (!document.documentElement.classList.contains(POPOVER_OPEN_CLASS)) {
      close()
      return
    }
    setTimeout(() => {
      if (document.documentElement.classList.contains(POPOVER_OPEN_CLASS)) return
      if (!target.isConnected) return
      close()
    }, 0)
  }
  const onKey = (e: KeyboardEvent) => {
    // with a ribbon popover open, Escape belongs to it: hiding the band would leave the
    // popover mounted (html marker stuck, menu back on the next peek)
    if (e.key !== 'Escape') return
    if (document.documentElement.classList.contains(POPOVER_OPEN_CLASS)) return
    close()
  }
  const onBlur = () => close()
  window.addEventListener('pointerdown', onPress, true)
  window.addEventListener('keydown', onKey)
  window.addEventListener('blur', onBlur)
  const offChrome = subscribeChromePressed(close)
  return () => {
    window.removeEventListener('pointerdown', onPress, true)
    window.removeEventListener('keydown', onKey)
    window.removeEventListener('blur', onBlur)
    offChrome?.()
  }
}

export interface RibbonCollapse {
  readonly collapsed: boolean
  /** collapsed and showing the band as an overlay */
  readonly peek: boolean
  readonly rootRef: RefObject<HTMLDivElement | null>
  /** class list for the ribbon root (append to the app's own classes) */
  readonly rootClass: string
  readonly toggle: () => void
  /** call from every tab button's click handler */
  readonly onTabPress: (wasActive: boolean) => void
  /** double-click on a tab toggles, like Office; attach to the tab row */
  readonly onTabsDoubleClick: (e: MouseEvent) => void
}

export function useRibbonCollapse(storageKey: string): RibbonCollapse {
  const [collapsed, setCollapsed] = useState(() => readRibbonCollapsed(storageKey))
  const [peek, setPeek] = useState(false)
  const rootRef = useRef<HTMLDivElement | null>(null)
  const collapsedRef = useRef(collapsed)
  collapsedRef.current = collapsed

  const toggle = useCallback(() => {
    const next = !collapsedRef.current
    writeRibbonCollapsed(storageKey, next)
    setCollapsed(next)
    setPeek(false)
  }, [storageKey])

  const onTabPress = useCallback((wasActive: boolean) => {
    if (!collapsedRef.current) return
    // pressing the already-peeked tab hides the band again (Office)
    setPeek((p) => !(p && wasActive))
  }, [])

  const onTabsDoubleClick = useCallback(
    (e: MouseEvent) => {
      const btn = (e.target as Element | null)?.closest('button')
      if (!btn || btn.classList.contains('qa-btn') || btn.classList.contains('ribbon-tab-file'))
        return
      toggle()
    },
    [toggle],
  )

  useEffect(() => {
    if (!collapsed || !peek) return
    return installRibbonPeekDismiss(
      () => rootRef.current,
      () => setPeek(false),
    )
  }, [collapsed, peek])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!isRibbonToggleShortcut(e)) return
      e.preventDefault()
      toggle()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [toggle])

  const rootClass = `ribbon-collapsible${collapsed ? ' ribbon-collapsed' : ''}${
    collapsed && peek ? ' ribbon-peek' : ''
  }`
  return { collapsed, peek, rootRef, rootClass, toggle, onTabPress, onTabsDoubleClick }
}

export interface RibbonCollapseLabels {
  readonly collapse: string
  readonly pin: string
}

function ChevronUp() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path
        d="M2.5 7.5 6 4l3.5 3.5"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function ChevronDown() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path
        d="M2.5 4.5 6 8l3.5-3.5"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function Pin() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path
        d="M4 1.5h4M5 1.5v3L3.5 6.5v1h5v-1L7 4.5v-3M6 7.5v3"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/** Corner button of the band: "Collapse the Ribbon" when expanded, "Pin the ribbon" while peeking. */
export function RibbonCollapseButton({
  state,
  labels,
}: {
  state: RibbonCollapse
  labels: RibbonCollapseLabels
}) {
  if (state.collapsed && !state.peek) return null
  const label = `${state.collapsed ? labels.pin : labels.collapse} (${RIBBON_TOGGLE_SHORTCUT})`
  return (
    <button
      type="button"
      className="ribbon-collapse-btn"
      data-tip={label}
      aria-label={label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={state.toggle}
    >
      {state.collapsed ? <Pin /> : <ChevronUp />}
    </button>
  )
}

/** Tab-row button for ribbons without tabs (nothing to press to peek): shown only while collapsed. */
export function RibbonExpandButton({ state, label }: { state: RibbonCollapse; label: string }) {
  if (!state.collapsed) return null
  const tip = `${label} (${RIBBON_TOGGLE_SHORTCUT})`
  return (
    <button
      type="button"
      className="ribbon-expand-btn"
      data-tip={tip}
      aria-label={tip}
      onMouseDown={(e) => e.preventDefault()}
      onClick={state.toggle}
    >
      <ChevronDown />
    </button>
  )
}

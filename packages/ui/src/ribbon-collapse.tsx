/**
 * Word for Mac "Collapse ribbon": the tab row stays, the command band hides.
 * The selected tab doubles as the collapse control; while collapsed no tab is
 * selected and pressing any tab expands the band again (it stays expanded —
 * no peek overlay, no pin). Double-clicking a tab and Ctrl+F1 (⌥⌘R on macOS)
 * toggle too. The state persists per app in localStorage.
 *
 * Markup contract: the ribbon root carries `rootRef` + `rootClass`, the band
 * element carries `data-ribbon-body`, tabs render `tabClass` / `tabTip` and
 * call `onTabPress`. Ribbons without tabs use `RibbonCollapseButton` +
 * `RibbonExpandButton` instead (ribbon-collapse.css anchors the former to the
 * band's corner).
 */
import { useCallback, useEffect, useRef, useState, type MouseEvent, type RefObject } from 'react'

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

export interface RibbonCollapseLabels {
  readonly collapse: string
  readonly expand: string
}

export interface RibbonCollapse {
  readonly collapsed: boolean
  readonly rootRef: RefObject<HTMLDivElement | null>
  /** class list for the ribbon root (append to the app's own classes) */
  readonly rootClass: string
  readonly toggle: () => void
  /** call from every tab button's click handler */
  readonly onTabPress: (wasActive: boolean) => void
  /** double-click on a tab toggles, like Office; attach to the tab row */
  readonly onTabsDoubleClick: (e: MouseEvent) => void
  /** `active` only while expanded: the collapsed tab row has no selected tab */
  readonly tabClass: (isActive: boolean) => string
  /** hover tip: the selected tab offers Collapse, every tab offers Expand while collapsed */
  readonly tabTip: (isActive: boolean) => string | undefined
}

export function useRibbonCollapse(
  storageKey: string,
  labels: RibbonCollapseLabels,
): RibbonCollapse {
  const [collapsed, setCollapsed] = useState(() => readRibbonCollapsed(storageKey))
  const rootRef = useRef<HTMLDivElement | null>(null)
  const collapsedRef = useRef(collapsed)
  collapsedRef.current = collapsed

  const toggle = useCallback(() => {
    const next = !collapsedRef.current
    writeRibbonCollapsed(storageKey, next)
    setCollapsed(next)
  }, [storageKey])

  // state before each of the last two presses: a double-click arrives after its
  // two clicks already ran onTabPress, and must end up toggled relative to the
  // state before the first of them
  const pressHistory = useRef<boolean[]>([])

  const onTabPress = useCallback(
    (wasActive: boolean) => {
      pressHistory.current = [...pressHistory.current.slice(-1), collapsedRef.current]
      if (collapsedRef.current || wasActive) toggle()
    },
    [toggle],
  )

  const onTabsDoubleClick = useCallback(
    (e: MouseEvent) => {
      const btn = (e.target as Element | null)?.closest('button')
      if (!btn || btn.classList.contains('qa-btn') || btn.classList.contains('ribbon-tab-file'))
        return
      const before = pressHistory.current[0] ?? collapsedRef.current
      pressHistory.current = []
      if (collapsedRef.current === before) toggle()
    },
    [toggle],
  )

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!isRibbonToggleShortcut(e)) return
      e.preventDefault()
      toggle()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [toggle])

  const tabClass = (isActive: boolean) => (isActive && !collapsed ? 'active' : '')
  const tabTip = (isActive: boolean) => {
    if (collapsed) return `${labels.expand} (${RIBBON_TOGGLE_SHORTCUT})`
    return isActive ? `${labels.collapse} (${RIBBON_TOGGLE_SHORTCUT})` : undefined
  }

  const rootClass = `ribbon-collapsible${collapsed ? ' ribbon-collapsed' : ''}`
  return { collapsed, rootRef, rootClass, toggle, onTabPress, onTabsDoubleClick, tabClass, tabTip }
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

/** Corner button of the band for ribbons without tabs (no selected tab to press): shown only while expanded. */
export function RibbonCollapseButton({ state, label }: { state: RibbonCollapse; label: string }) {
  if (state.collapsed) return null
  const tip = `${label} (${RIBBON_TOGGLE_SHORTCUT})`
  return (
    <button
      type="button"
      className="ribbon-collapse-btn"
      data-tip={tip}
      aria-label={tip}
      onMouseDown={(e) => e.preventDefault()}
      onClick={state.toggle}
    >
      <ChevronUp />
    </button>
  )
}

/** Tab-row button for ribbons without tabs (nothing to press to expand): shown only while collapsed. */
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

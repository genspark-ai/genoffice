/**
 * Word for Mac "Collapse ribbon": the tab row stays, the command band hides.
 * The selected tab doubles as the collapse control; while collapsed no tab is
 * selected and pressing any tab expands the band again (it stays expanded —
 * no peek overlay, no pin). Double-clicking a tab and Ctrl+F1 (⌥⌘R on macOS)
 * toggle too. The state persists per app in localStorage.
 *
 * An optional middle density, "compact" (icon-only commands on a shorter
 * band), is opt-in per call site: it needs `labels.compact` / `labels.expandFull`
 * and is otherwise inert, so ribbons that only pass collapse/expand keep the
 * original two-state behaviour and stay byte-for-byte compatible. ⌥⌘K
 * (Ctrl+Alt+K) cycles full ↔ compact. See #362.
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
export const RIBBON_COMPACT_SHORTCUT = IS_MAC ? '⌥⌘K' : 'Ctrl+Alt+K'

/**
 * How much of the command band to show. `full` and `compact` both keep the band
 * visible (and keep the active tab highlighted); `collapsed` hides it entirely.
 */
export type RibbonDensity = 'full' | 'compact' | 'collapsed'

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

/**
 * Reads the persisted density, honouring the legacy `'1'`/`'0'` values written
 * by the two-state version so an existing collapsed ribbon stays collapsed.
 */
export function readRibbonDensity(storageKey: string): RibbonDensity {
  try {
    const raw = localStorage.getItem(storageKey)
    if (raw === 'collapsed' || raw === 'compact') return raw
    if (raw === '1') return 'collapsed'
    return 'full'
  } catch {
    return 'full'
  }
}

export function writeRibbonDensity(storageKey: string, density: RibbonDensity): void {
  try {
    localStorage.setItem(storageKey, density)
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

/** ⌥⌘K on macOS, Ctrl+Alt+K elsewhere. ⌥⌘R (collapse) and every ⌥⌘<letter> the
 * editors bind — G/M/A/F/E/D — are already taken, hence K. */
export function isRibbonCompactShortcut(e: KeyboardEvent): boolean {
  if (e.repeat || e.shiftKey) return false
  if (IS_MAC) return e.metaKey && e.altKey && !e.ctrlKey && e.code === 'KeyK'
  return e.ctrlKey && e.altKey && !e.metaKey && e.code === 'KeyK'
}

export interface RibbonCollapseLabels {
  readonly collapse: string
  readonly expand: string
  /** opt in: the icon-only middle density */
  readonly compact?: string
  /** opt in: label for leaving `compact` back to `full` */
  readonly expandFull?: string
}

export interface RibbonCollapse {
  readonly collapsed: boolean
  /** current band density; `'full'` when the call site did not opt in */
  readonly density: RibbonDensity
  /** true when the call site passed the compact labels */
  readonly compactEnabled: boolean
  readonly rootRef: RefObject<HTMLDivElement | null>
  /** class list for the ribbon root (append to the app's own classes) */
  readonly rootClass: string
  readonly toggle: () => void
  /** cycle full ↔ compact (no-op when the call site did not opt in) */
  readonly toggleCompact: () => void
  /** call from every tab button's click handler */
  readonly onTabPress: (wasActive: boolean) => void
  /** double-click on a tab toggles, like Office; attach to the tab row */
  readonly onTabsDoubleClick: (e: MouseEvent) => void
  /** `active` whenever the band is visible: a collapsed tab row has no selected tab */
  readonly tabClass: (isActive: boolean) => string
  /** hover tip: the selected tab offers Collapse, every tab offers Expand while collapsed */
  readonly tabTip: (isActive: boolean) => string | undefined
}

export function useRibbonCollapse(
  storageKey: string,
  labels: RibbonCollapseLabels,
): RibbonCollapse {
  const compactEnabled = Boolean(labels.compact && labels.expandFull)
  const [density, setDensity] = useState<RibbonDensity>(() =>
    compactEnabled
      ? readRibbonDensity(storageKey)
      : readRibbonCollapsed(storageKey)
        ? 'collapsed'
        : 'full',
  )
  const rootRef = useRef<HTMLDivElement | null>(null)
  const collapsed = density === 'collapsed'
  const densityRef = useRef(density)
  densityRef.current = density

  const persist = useCallback(
    (next: RibbonDensity) => {
      // the legacy key only ever stored '1'/'0', so keep writing that shape for
      // the two states and use the name for the middle one
      if (!compactEnabled) {
        writeRibbonCollapsed(storageKey, next === 'collapsed')
        return
      }
      writeRibbonDensity(storageKey, next)
    },
    [compactEnabled, storageKey],
  )

  const toggle = useCallback(() => {
    // the band control hides the band or brings it back at full size; compact is
    // not a stopping point for it
    const next: RibbonDensity = densityRef.current === 'collapsed' ? 'full' : 'collapsed'
    persist(next)
    setDensity(next)
  }, [persist])

  const toggleCompact = useCallback(() => {
    if (!compactEnabled) return
    // from collapsed this reveals the band compactly, matching what the user
    // just asked for rather than jumping straight back to full
    const next: RibbonDensity = densityRef.current === 'compact' ? 'full' : 'compact'
    persist(next)
    setDensity(next)
  }, [compactEnabled, persist])

  // state before each of the last two presses: a double-click arrives after its
  // two clicks already ran onTabPress, and must end up toggled relative to the
  // state before the first of them
  const pressHistory = useRef<RibbonDensity[]>([])

  const onTabPress = useCallback(
    (wasActive: boolean) => {
      pressHistory.current = [...pressHistory.current.slice(-1), densityRef.current]
      if (densityRef.current === 'collapsed' || wasActive) toggle()
    },
    [toggle],
  )

  const onTabsDoubleClick = useCallback(
    (e: MouseEvent) => {
      const btn = (e.target as Element | null)?.closest('button')
      if (!btn || btn.classList.contains('qa-btn') || btn.classList.contains('ribbon-tab-file'))
        return
      const before = pressHistory.current[0] ?? densityRef.current
      pressHistory.current = []
      if (densityRef.current === before) toggle()
    },
    [toggle],
  )

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (isRibbonToggleShortcut(e)) {
        e.preventDefault()
        toggle()
        return
      }
      if (compactEnabled && isRibbonCompactShortcut(e)) {
        e.preventDefault()
        toggleCompact()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [toggle, toggleCompact, compactEnabled])

  const tabClass = (isActive: boolean) => (isActive && !collapsed ? 'active' : '')
  const tabTip = (isActive: boolean) => {
    if (collapsed) return `${labels.expand} (${RIBBON_TOGGLE_SHORTCUT})`
    if (!isActive) return undefined
    if (!compactEnabled) return `${labels.collapse} (${RIBBON_TOGGLE_SHORTCUT})`
    // while compact the band toggle would collapse; offer both ways out
    return density === 'compact'
      ? `${labels.expandFull} (${RIBBON_COMPACT_SHORTCUT})`
      : `${labels.collapse} (${RIBBON_TOGGLE_SHORTCUT}) · ${labels.compact} (${RIBBON_COMPACT_SHORTCUT})`
  }

  const rootClass = `ribbon-collapsible${collapsed ? ' ribbon-collapsed' : ''}${
    density === 'compact' ? ' ribbon-compact' : ''
  }`
  return {
    collapsed,
    density,
    compactEnabled,
    rootRef,
    rootClass,
    toggle,
    toggleCompact,
    onTabPress,
    onTabsDoubleClick,
    tabClass,
    tabTip,
  }
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

/**
 * Floating quick-action launcher that appears next to a text selection.
 *
 * Pattern mirrors `apps/slides/src/renderer/components/AiAskPopover.tsx`
 * (anchor rect refreshed on scroll/resize, viewport-clamped, mounted outside
 * any zoom-transformed subtree) but renders an icon-chip row instead of an
 * input — when the user clicks a chip the host app is responsible for
 * opening a full dialog (e.g. `TranslateDialog` for the Translate chip).
 *
 * Behaviour:
 *   - Position is measured in viewport coords via `getAnchorRect()` so the
 *     launcher never inherits the editor's zoom transform.
 *   - The launcher hides when `getAnchorRect()` returns null (selection
 *     collapsed or scrolled out of the visible band).
 *   - Picking a chip does NOT close the launcher: the user may want to
 *     change their mind. Press Esc to dismiss.
 *   - Chip order: Polish / Expand / Shorten / Summarize / Translate — same
 *     as Codex desktop's "Rewrite" + WorkBuddy's "一句话" pattern; Translate
 *     is intentionally last because it opens a dialog rather than a one-shot.
 */

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'

export type AiInlineAction = 'polish' | 'expand' | 'shorten' | 'summarize' | 'translate'

export interface AiInlineLauncherAnchorRect {
  left: number
  top: number
  right: number
  bottom: number
  viewTop?: number
  viewBottom?: number
}

export interface AiInlineLauncherStrings {
  title: string
  polish: string
  expand: string
  shorten: string
  summarize: string
  translate: string
}

export interface AiInlineLauncherProps {
  /** Re-measured on every layout-affecting event; null hides the chip. */
  getAnchorRect: () => AiInlineLauncherAnchorRect | null
  /** Which quick-action chips to show (defaults to all five). */
  actions?: AiInlineAction[]
  /** Chip + tooltip strings; the host supplies localised text. */
  strings: AiInlineLauncherStrings
  /** Chip click handler. The launcher does NOT auto-close on pick. */
  onPick: (action: AiInlineAction) => void
  /** Optional className passthrough for app-specific theming. */
  className?: string
  /**
   * Bump whenever the host knows the underlying selection or viewport
   * changed (e.g. editor's `selectionUpdate` event). The launcher
   * recomputes its anchor rect on each revision so it can follow a
   * selection that lives inside a memoised React subtree.
   */
  revision?: number
}

const GAP = 8
const EDGE = 8
const EST_HEIGHT = 36
const EST_WIDTH = 220

const DEFAULT_ACTIONS: AiInlineAction[] = ['polish', 'expand', 'shorten', 'summarize', 'translate']

export function AiInlineLauncher(props: AiInlineLauncherProps): React.JSX.Element | null {
  const { getAnchorRect, actions = DEFAULT_ACTIONS, strings, onPick, className } = props
  const [pos, setPos] = useState<{ left: number; top: number; placeBelow: boolean } | null>(null)
  const ref = useRef<HTMLDivElement>(null)

  // Re-measure on every render via a tick prop; parents bump it when the
  // underlying selection / viewport state they care about changes.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const _tick = (props as unknown as { revision?: number }).revision
  useLayoutEffect(() => {
    const update = () => {
      const rect = getAnchorRect()
      if (!rect) {
        setPos(null)
        return
      }
      const w = ref.current?.offsetWidth ?? EST_WIDTH
      const h = ref.current?.offsetHeight ?? EST_HEIGHT
      const below = rect.bottom + GAP
      const above = rect.top - GAP - h
      const viewBottom = rect.viewBottom ?? window.innerHeight
      const viewTop = rect.viewTop ?? 0
      const placeBelow = below + h + EDGE <= viewBottom || above < viewTop
      const top = placeBelow ? below : above
      const left = Math.min(
        Math.max(EDGE, rect.left + (rect.right - rect.left) / 2 - w / 2),
        window.innerWidth - w - EDGE,
      )
      setPos({ left, top: Math.max(EDGE, top), placeBelow })
    }
    update()
    const onScroll = () => update()
    window.addEventListener('scroll', onScroll, { passive: true, capture: true })
    window.addEventListener('resize', onScroll)
    return () => {
      window.removeEventListener('scroll', onScroll, { capture: true })
      window.removeEventListener('resize', onScroll)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [getAnchorRect, _tick])

  useEffect(() => {
    if (!pos) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setPos(null)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [pos])

  if (!pos) return null

  return (
    <div
      ref={ref}
      role="toolbar"
      aria-label={strings.title}
      className={['ai-inline-launcher', className].filter(Boolean).join(' ')}
      style={{ left: `${pos.left}px`, top: `${pos.top}px` }}
      data-place={pos.placeBelow ? 'below' : 'above'}
    >
      {actions.map((action) => (
        <button
          key={action}
          type="button"
          className={`ai-inline-launcher-chip ai-inline-launcher-chip--${action}`}
          aria-label={strings[action]}
          title={strings[action]}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => onPick(action)}
        >
          <AiInlineIcon action={action} />
          <span className="ai-inline-launcher-label">{strings[action]}</span>
        </button>
      ))}
    </div>
  )
}

function AiInlineIcon({ action }: { action: AiInlineAction }): React.JSX.Element {
  switch (action) {
    case 'polish':
      return (
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden>
          <path
            d="M11.5 2.5L13.5 4.5L5 13H3V11L11.5 2.5Z"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinejoin="round"
          />
        </svg>
      )
    case 'expand':
      return (
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden>
          <path
            d="M3 5h10M3 8h10M3 11h7"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
          />
        </svg>
      )
    case 'shorten':
      return (
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden>
          <path d="M3 6h10M3 10h6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
        </svg>
      )
    case 'summarize':
      return (
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden>
          <path
            d="M3 3h7M3 6h10M3 9h8M3 12h6"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
          />
        </svg>
      )
    case 'translate':
      return (
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden>
          <path
            d="M2 3h6M5 3v1.5C5 7 3.5 8.5 2 9M6 5.5C5.5 7 4.5 8 3 8.5M9 13l2-5 2 5M9.7 11.5h2.6"
            stroke="currentColor"
            strokeWidth="1.3"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      )
  }
}

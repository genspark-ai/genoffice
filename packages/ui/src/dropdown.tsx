/**
 * Shared select replacement: a framed trigger (current option + caret) opening
 * a themed, anchored popover of options — native <select> popups are OS-drawn
 * and ignore the app theme. Options render arbitrary content (line previews,
 * text) and stay keyboard-operable (arrows / Enter / Escape / Home / End while
 * the trigger keeps focus). The popover is position: fixed + CSS-anchored to
 * the trigger so scroll-container overflow never clips it, with a flip-block
 * fallback near the viewport bottom.
 *
 * Styling comes from dropdown.css (gs-dd* classes, token colors only); apps
 * size the control via `className` on the wrapper.
 */
import React, { useEffect, useId, useMemo, useRef, useState } from 'react'
import { useDismissablePopover } from './popover-dismiss'

export interface DropdownOption<K extends string = string> {
  readonly value: K
  /** Accessible name (aria-label/title); also the visible text when `render` is omitted. */
  readonly label: string
  readonly render?: React.ReactNode
  /** Shown but not pickable (placeholder rows like "Auto", unavailable modes). */
  readonly disabled?: boolean
}

export function nextEnabledIndex(
  options: ReadonlyArray<{ readonly disabled?: boolean }>,
  start: number,
  step: 1 | -1,
): number {
  for (let i = start; i >= 0 && i < options.length; i += step) {
    if (!options[i]!.disabled) return i
  }
  return -1
}

export function reconcileActiveIndex(
  options: ReadonlyArray<{ readonly value?: string; readonly disabled?: boolean }>,
  active: number,
  value?: string,
): number {
  if (active >= 0 && active < options.length && !options[active]!.disabled) return active
  if (value !== undefined) {
    const selected = options.findIndex((option) => option.value === value)
    if (selected >= 0 && !options[selected]!.disabled) return selected
  }
  const start = active >= 0 && active < options.length ? active : 0
  const forward = nextEnabledIndex(options, start, 1)
  return forward >= 0 ? forward : nextEnabledIndex(options, Math.min(start, options.length - 1), -1)
}

export function filterOptions<T extends { readonly value: string; readonly label: string }>(
  options: ReadonlyArray<T>,
  query: string,
): T[] {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean)
  return options.filter((o) => {
    const hay = `${o.value} ${o.label}`.toLowerCase()
    return terms.every((term) => hay.includes(term))
  })
}

export function Dropdown<K extends string>({
  value,
  options,
  onPick,
  className,
  ariaLabel,
  disabled,
  tip,
  ariaRequired,
  ariaInvalid,
  searchPlaceholder,
}: {
  readonly value: K
  readonly options: ReadonlyArray<DropdownOption<K>>
  readonly onPick: (value: K) => void
  /** Extra class on the wrapper (apps set width there). */
  readonly className?: string
  /** Control name for screen readers; defaults to the current option's label. */
  readonly ariaLabel?: string
  readonly disabled?: boolean
  /** ScreenTip text (data-tip on the trigger). */
  readonly tip?: string
  /** Form semantics passthrough (AcroForm widgets etc.). */
  readonly ariaRequired?: boolean
  readonly ariaInvalid?: boolean
  /** Set to put a filter box atop the list (long catalogs); arrows/Enter keep working from it. */
  readonly searchPlaceholder?: string
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(0)
  const [query, setQuery] = useState('')
  const searchRef = useRef<HTMLInputElement>(null)
  const all = options
  const visible = useMemo(
    () => (searchPlaceholder && query.trim() ? filterOptions(all, query) : all),
    [all, query, searchPlaceholder],
  )
  const popRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLSpanElement>(null)
  // Focus stays on the trigger (menu-button pattern), so the open listbox and
  // the option the arrows are on are referenced by id instead of by focus.
  const listId = useId()
  const optionId = (index: number): string => `${listId}-option-${index}`
  // guarded (capture-phase) dismissal: a press on another dropdown's trigger
  // must close this one even though that trigger stops mousedown propagation
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEffect(() => {
    if (!open) return
    // optional chaining on the call: jsdom elements have no scrollIntoView
    popRef.current?.querySelectorAll('.gs-dd-item')[active]?.scrollIntoView?.({ block: 'nearest' })
  }, [open, active])
  useEffect(() => {
    setActive((current) => reconcileActiveIndex(visible, current, value))
  }, [visible, value])
  // No fallback to options[0]: an off-list value (e.g. a document-only font)
  // must read as itself, not masquerade as the first option
  const current = all.find((o) => o.value === value)
  // mirrors the .active class: a disabled row is never the active option
  const activeOption = open && visible[active] && !visible[active]!.disabled ? active : null
  const openList = () => {
    setQuery('')
    const i = all.findIndex((o) => o.value === value)
    setActive(reconcileActiveIndex(all, i, value))
    setOpen(true)
  }
  useEffect(() => {
    if (open && searchPlaceholder) searchRef.current?.focus()
  }, [open, searchPlaceholder])
  const pick = (o: DropdownOption<K>) => {
    if (o.disabled) return
    setOpen(false)
    onPick(o.value)
  }
  const move = (step: 1 | -1) => {
    setActive((i) => {
      const current = i >= 0 && i < visible.length ? i : -1
      const start = current < 0 ? (step === 1 ? 0 : visible.length - 1) : current + step
      const next = nextEnabledIndex(visible, start, step)
      return next < 0 ? reconcileActiveIndex(visible, current, value) : next
    })
  }
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (!open) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === ' ') {
        e.preventDefault()
        openList()
      }
      return
    }
    // a space typed into the filter box is text, and keys during an IME
    // composition belong to the IME: neither is a pick, and neither may reach
    // the hosting modal (an Escape there would close it)
    if (e.target === searchRef.current && (e.key === ' ' || e.nativeEvent.isComposing)) {
      e.stopPropagation()
      return
    }
    if (e.key === 'Escape') setOpen(false)
    else if (e.key === 'ArrowDown') move(1)
    else if (e.key === 'ArrowUp') move(-1)
    else if (e.key === 'Home') setActive(nextEnabledIndex(visible, 0, 1))
    else if (e.key === 'End') setActive(nextEnabledIndex(visible, visible.length - 1, -1))
    else if (e.key === 'Enter' || e.key === ' ') {
      const o = visible[active]
      if (o && !o.disabled) pick(o)
      else {
        const next = reconcileActiveIndex(visible, active, value)
        setActive(next)
        const nextOption = visible[next]
        if (nextOption) pick(nextOption)
      }
    } else return
    e.preventDefault()
    // handled keys stay ours while the list is open: a bubbling Escape would
    // close the hosting modal, bubbling arrows would nudge canvas elements
    e.stopPropagation()
  }
  return (
    <span ref={wrapRef} className={`gs-dd${className ? ` ${className}` : ''}`}>
      <button
        type="button"
        className="gs-dd-btn"
        disabled={disabled}
        data-value={value}
        data-tip={tip}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={open ? listId : undefined}
        aria-activedescendant={activeOption === null ? undefined : optionId(activeOption)}
        aria-label={ariaLabel ?? current?.label ?? value}
        aria-required={ariaRequired}
        aria-invalid={ariaInvalid}
        onClick={() => (open ? setOpen(false) : openList())}
        onKeyDown={onKeyDown}
        onBlur={(e) => {
          // native selects close on focus loss (Tab); staying inside the wrapper
          // (clicking an option focuses it) must not dismiss
          if (!wrapRef.current?.contains(e.relatedTarget as Node)) setOpen(false)
        }}
      >
        <span className="gs-dd-value">{current ? (current.render ?? current.label) : value}</span>
        <span className="gs-dd-caret" aria-hidden="true">
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none">
            <path
              d="M5.5 9.25 12 15.75l6.5-6.5"
              stroke="currentColor"
              strokeWidth="2.6"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </span>
      </button>
      {open && (
        <div
          ref={popRef}
          className={`gs-dd-pop${searchPlaceholder ? ' has-search' : ''}`}
          role="listbox"
          id={listId}
        >
          {searchPlaceholder && (
            <input
              ref={searchRef}
              className="gs-dd-search"
              type="text"
              value={query}
              placeholder={searchPlaceholder}
              aria-label={searchPlaceholder}
              spellCheck={false}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={onKeyDown}
              onBlur={(e) => {
                if (!wrapRef.current?.contains(e.relatedTarget as Node)) setOpen(false)
              }}
            />
          )}
          <div className="gs-dd-list">
            {visible.map((o, i) => (
              <button
                key={o.value}
                id={optionId(i)}
                type="button"
                role="option"
                // menu-button pattern: options never join the tab order (focus
                // lives on the trigger; Tab away closes via its onBlur)
                tabIndex={-1}
                disabled={o.disabled}
                aria-selected={o.value === value}
                aria-label={o.label}
                data-value={o.value}
                title={o.label}
                className={`gs-dd-item${o.value === value ? ' selected' : ''}${i === active && !o.disabled ? ' active' : ''}`}
                onMouseEnter={() => {
                  if (!o.disabled) setActive(i)
                }}
                // menu-button pattern: options never take focus, so picking one
                // can't blur focus-scoped hosts (the PDF text editor commits its
                // draft on focus leaving the edit bar)
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => pick(o)}
              >
                {o.render ?? o.label}
              </button>
            ))}
          </div>
        </div>
      )}
    </span>
  )
}

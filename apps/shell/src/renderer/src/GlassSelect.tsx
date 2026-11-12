import { useCallback, useEffect, useId, useLayoutEffect, useRef, useState } from 'react'
import type { CSSProperties, ReactNode } from 'react'
import { createPortal } from 'react-dom'
import './glass-select.css'

// Replacement for the native <select>: the OS popup can't be themed (that blue list),
// so this draws its own glass listbox. Keyboard: ↑ ↓ Home End Enter Space Esc Tab + type-ahead + optional search.

export interface GlassOption {
  value: string
  label: string
  disabled?: boolean
  render?: ReactNode
}

interface GlassSelectProps {
  /** id for the trigger, so a <label htmlFor> still points at the control */
  id?: string
  value: string
  options: readonly GlassOption[]
  onChange: (value: string) => void
  /** extra classes on the trigger (e.g. 'ap-pill') */
  className?: string
  'aria-label'?: string
  searchPlaceholder?: string
}

/** measuring state: fixed + invisible, so it can't extend the page or trigger scrolling */
const HIDDEN: CSSProperties = { position: 'fixed', top: 0, left: 0, visibility: 'hidden' }
const GAP = 6
const VIEWPORT_MARGIN = 8

export function GlassSelect({
  id,
  value,
  options,
  onChange,
  className,
  'aria-label': ariaLabel,
  searchPlaceholder,
}: GlassSelectProps) {
  const uid = useId()
  const triggerRef = useRef<HTMLButtonElement>(null)
  const menuRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(0)
  const [search, setSearch] = useState('')
  const [style, setStyle] = useState<CSSProperties>(HIDDEN)
  const typed = useRef({ text: '', at: 0 })

  const selected = options.find((o) => o.value === value)
  const filteredOptions =
    searchPlaceholder && search.trim()
      ? options.filter((o) => o.label.toLowerCase().includes(search.trim().toLowerCase()))
      : options
  const enabled = filteredOptions.map((o, i) => (o.disabled ? -1 : i)).filter((i) => i >= 0)

  const close = useCallback((refocus: boolean) => {
    setOpen(false)
    setSearch('')
    if (refocus) triggerRef.current?.focus()
  }, [])

  const openMenu = (): void => {
    const current = filteredOptions.findIndex((o) => o.value === value && !o.disabled)
    setActive(current >= 0 ? current : (enabled[0] ?? 0))
    setSearch('')
    setStyle(HIDDEN)
    setOpen(true)
  }

  const choose = (index: number): void => {
    const opt = filteredOptions[index]
    if (!opt || opt.disabled) return
    if (opt.value !== value) onChange(opt.value)
    close(true)
  }

  // anchor under the trigger, right edges aligned; flip above when there's no room below
  useLayoutEffect(() => {
    if (!open) return
    const trigger = triggerRef.current
    const menu = menuRef.current
    if (!trigger || !menu) return
    const rect = trigger.getBoundingClientRect()
    const height = menu.offsetHeight
    const below = window.innerHeight - rect.bottom - GAP - VIEWPORT_MARGIN
    const above = rect.top - GAP - VIEWPORT_MARGIN
    const flip = height > below && above > below
    const room = Math.max(120, flip ? above : below)
    setStyle({
      position: 'fixed',
      right: Math.max(VIEWPORT_MARGIN, window.innerWidth - rect.right),
      minWidth: Math.max(rect.width, 160),
      maxHeight: Math.min(340, room),
      ...(flip ? { bottom: window.innerHeight - rect.top + GAP } : { top: rect.bottom + GAP }),
    })
  }, [open, search])

  // focus the list or search input once visible
  const positioned = open && style.visibility !== 'hidden'
  useEffect(() => {
    if (positioned) {
      if (searchPlaceholder && searchInputRef.current) {
        searchInputRef.current.focus({ preventScroll: true })
      } else {
        menuRef.current?.focus({ preventScroll: true })
      }
    }
  }, [positioned, searchPlaceholder])

  // dismiss on outside press, resize, or scrolling the page behind
  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: PointerEvent): void => {
      const target = e.target as Node
      if (menuRef.current?.contains(target) || triggerRef.current?.contains(target)) return
      close(false)
    }
    const onScroll = (e: Event): void => {
      if (menuRef.current?.contains(e.target as Node)) return
      close(false)
    }
    const onResize = (): void => close(false)
    document.addEventListener('pointerdown', onPointerDown, true)
    window.addEventListener('scroll', onScroll, true)
    window.addEventListener('resize', onResize)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true)
      window.removeEventListener('scroll', onScroll, true)
      window.removeEventListener('resize', onResize)
    }
  }, [open, close])

  // keep the highlighted option in view by scrolling the list itself
  useEffect(() => {
    const menu = menuRef.current
    if (!open || !menu || style.visibility === 'hidden') return
    const el = menu.querySelector<HTMLElement>(`[data-index="${active}"]`)
    if (!el) return
    if (el.offsetTop < menu.scrollTop) menu.scrollTop = el.offsetTop - 4
    else if (el.offsetTop + el.offsetHeight > menu.scrollTop + menu.clientHeight)
      menu.scrollTop = el.offsetTop + el.offsetHeight - menu.clientHeight + 4
  }, [open, active, style])

  const move = (delta: 1 | -1): void => {
    if (enabled.length === 0) return
    const at = enabled.indexOf(active)
    const next = enabled[(at + delta + enabled.length) % enabled.length]
    if (next !== undefined) setActive(next)
  }

  const onTriggerKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault()
      openMenu()
    }
  }

  const onSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    switch (e.key) {
      case 'Escape':
        e.preventDefault()
        e.stopPropagation()
        close(true)
        return
      case 'ArrowDown':
        e.preventDefault()
        move(1)
        return
      case 'ArrowUp':
        e.preventDefault()
        move(-1)
        return
      case 'Enter':
        e.preventDefault()
        if (enabled.length > 0 && active >= 0 && active < filteredOptions.length) {
          choose(active)
        }
        return
      case 'Tab':
        close(false)
        return
    }
  }

  const onMenuKeyDown = (e: React.KeyboardEvent): void => {
    switch (e.key) {
      case 'Escape':
        e.preventDefault()
        e.stopPropagation()
        close(true)
        return
      case 'ArrowDown':
        e.preventDefault()
        move(1)
        return
      case 'ArrowUp':
        e.preventDefault()
        move(-1)
        return
      case 'Home':
        e.preventDefault()
        if (enabled[0] !== undefined) setActive(enabled[0])
        return
      case 'End':
        e.preventDefault()
        if (enabled.length) setActive(enabled[enabled.length - 1]!)
        return
      case 'Enter':
      case ' ':
        e.preventDefault()
        if (enabled.length > 0 && active >= 0 && active < filteredOptions.length) {
          choose(active)
        }
        return
      case 'Tab':
        close(false)
        return
    }
    // type-ahead: jump to the next option starting with what was typed (when no search input)
    if (!searchPlaceholder && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const now = Date.now()
      const buf = typed.current
      buf.text = now - buf.at > 600 ? e.key.toLowerCase() : buf.text + e.key.toLowerCase()
      buf.at = now
      const start = buf.text.length === 1 ? enabled.indexOf(active) + 1 : enabled.indexOf(active)
      const order = [...enabled.slice(start), ...enabled.slice(0, start)]
      const hit = order.find((i) => filteredOptions[i]?.label.toLowerCase().startsWith(buf.text))
      if (hit !== undefined) setActive(hit)
    }
  }

  const listId = `${uid}-list`

  const menu = open && (
    <div
      ref={menuRef}
      id={listId}
      className="gsel-menu"
      role="listbox"
      tabIndex={-1}
      aria-label={ariaLabel}
      aria-activedescendant={`${uid}-opt-${active}`}
      style={style}
      onKeyDown={onMenuKeyDown}
    >
      {searchPlaceholder && (
        <div className="gsel-search-wrap">
          <svg className="gsel-search-icon" width="12" height="12" viewBox="0 0 16 16" fill="none">
            <circle cx="7" cy="7" r="4.5" stroke="currentColor" strokeWidth="1.5" />
            <path
              d="m10.5 10.5 3.5 3.5"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
            />
          </svg>
          <input
            ref={searchInputRef}
            type="text"
            className="gsel-search-input"
            placeholder={searchPlaceholder}
            value={search}
            onChange={(e) => {
              setSearch(e.target.value)
              setActive(0)
            }}
            onKeyDown={onSearchKeyDown}
          />
        </div>
      )}
      {filteredOptions.length === 0 ? (
        <div className="gsel-empty">No results</div>
      ) : (
        filteredOptions.map((opt, i) => (
          <div
            key={opt.value}
            id={`${uid}-opt-${i}`}
            data-index={i}
            role="option"
            aria-selected={opt.value === value}
            aria-disabled={opt.disabled || undefined}
            className={`gsel-option${i === active ? ' active' : ''}${opt.disabled ? ' disabled' : ''}`}
            onMouseMove={() => !opt.disabled && i !== active && setActive(i)}
            onClick={() => choose(i)}
          >
            <span className="gsel-label">{opt.render ?? opt.label}</span>
            {opt.value === value && (
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
                <path
                  d="m2.8 7.3 2.8 2.8 5.6-6"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            )}
          </div>
        ))
      )}
    </div>
  )

  return (
    <>
      <button
        ref={triggerRef}
        id={id}
        type="button"
        className={`set-select-wrap gsel-trigger${className ? ` ${className}` : ''}`}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={open ? listId : undefined}
        aria-label={ariaLabel}
        onClick={() => (open ? close(false) : openMenu())}
        onKeyDown={onTriggerKeyDown}
      >
        <span className="set-select-text">
          {selected ? (selected.render ?? selected.label) : ''}
        </span>
      </button>
      {menu && createPortal(menu, document.body)}
    </>
  )
}

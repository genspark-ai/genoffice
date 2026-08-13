import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'

/**
 * Minimal Google-Docs-style menu bar: a row of top-level labels, each opening
 * a flyout of items (with optional nested submenus). Closes on outside click
 * or Escape; arrow keys are not implemented (out of scope for this pass).
 */

export interface MenuItem {
  /** divider row */
  divider?: boolean
  label?: string
  shortcut?: string
  onSelect?: () => void
  disabled?: boolean
  checked?: boolean
  /** nested flyout (Format ▸ Text ▸ Bold) */
  submenu?: MenuItem[]
}

export interface MenuSpec {
  key: string
  label: string
  items: MenuItem[]
}

function MenuItemRow({ item, onDone }: { item: MenuItem; onDone: () => void }) {
  const [subOpen, setSubOpen] = useState(false)
  if (item.divider) return <div className="gdh-menu-divider" role="separator" />
  if (item.submenu) {
    return (
      <div
        className="gdh-menu-item gdh-menu-item-parent"
        onMouseEnter={() => setSubOpen(true)}
        onMouseLeave={() => setSubOpen(false)}
      >
        <span className="gdh-menu-check">{item.checked ? '✓' : ''}</span>
        <span className="gdh-menu-label">{item.label}</span>
        <span className="gdh-menu-arrow">›</span>
        {subOpen && (
          <div className="gdh-submenu">
            {item.submenu.map((sub, i) => (
              <MenuItemRow key={sub.label ?? `d${i}`} item={sub} onDone={onDone} />
            ))}
          </div>
        )}
      </div>
    )
  }
  return (
    <button
      type="button"
      className="gdh-menu-item"
      disabled={item.disabled}
      onClick={() => {
        item.onSelect?.()
        onDone()
      }}
    >
      <span className="gdh-menu-check">{item.checked ? '✓' : ''}</span>
      <span className="gdh-menu-label">{item.label}</span>
      {item.shortcut && <span className="gdh-menu-shortcut">{item.shortcut}</span>}
    </button>
  )
}

export function GDocsMenuBar({ menus }: { menus: MenuSpec[] }) {
  const [openKey, setOpenKey] = useState<string | null>(null)
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!openKey) return
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpenKey(null)
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpenKey(null)
    }
    window.addEventListener('mousedown', onDoc)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDoc)
      window.removeEventListener('keydown', onKey)
    }
  }, [openKey])

  return (
    <div className="gdh-menubar" ref={ref}>
      {menus.map((m) => (
        <div key={m.key} className="gdh-menu-wrap">
          <button
            type="button"
            className={`gdh-menu-top ${openKey === m.key ? 'open' : ''}`}
            onClick={() => setOpenKey((k) => (k === m.key ? null : m.key))}
            onMouseEnter={() => setOpenKey((k) => (k ? m.key : k))}
          >
            {m.label}
          </button>
          {openKey === m.key && (
            <div className="gdh-menu-panel">
              {m.items.map((item, i) => (
                <MenuItemRow
                  key={item.label ?? `d${i}`}
                  item={item}
                  onDone={() => setOpenKey(null)}
                />
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  )
}

export function menuDivider(): MenuItem {
  return { divider: true }
}

/** small helper so callers can build item arrays tersely, skipping items whose command doesn't exist */
export function mi(
  label: string,
  onSelect: (() => void) | undefined | false | null,
  opts: { shortcut?: string; checked?: boolean; submenu?: MenuItem[] } = {},
): MenuItem | null {
  if (!onSelect && !opts.submenu) return null
  return {
    label,
    onSelect: onSelect || undefined,
    shortcut: opts.shortcut,
    checked: opts.checked,
    submenu: opts.submenu,
  }
}

export function compact(items: Array<MenuItem | null>): MenuItem[] {
  return items.filter((i): i is MenuItem => i != null)
}

export type { ReactNode }

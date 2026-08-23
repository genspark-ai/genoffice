/**
 * Color tool trigger + the shared Word-style palette (@genoffice/ui ColorPicker),
 * anchored below its trigger via the `.menu-select` / `.sheets-color-pop` CSS
 * anchor pair. Used by the ribbon (swatch-letter trigger via `display`) and by
 * dialogs/panes (plain color-well trigger when `display` is omitted).
 *
 * Same outside-click/Escape/chrome-press closing as MenuSelect; no blur close —
 * opening the native "More Colors" dialog blurs the window and must not tear
 * the panel down while the OS picker is still open.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { ColorPicker } from '@genoffice/ui'
// module-level t (not useI18n): this dropdown also renders inside Univer's
// float-DOM React root, which lives outside our LocaleProvider; the module
// translator is kept in sync on every language switch and re-renders arrive
// with the host component
import { t } from './i18n/locale'

/// Portal wrapper for hosts inside Univer's float DOM (chart editor): the
/// float container is transformed and clips overflow, which breaks the CSS
/// anchor positioning — position the panel manually against the trigger's
/// viewport rect, clamped to the window (same escape as ChartContextMenu).
function PortalPop({
  anchor,
  popRef,
  children,
}: {
  readonly anchor: HTMLElement
  readonly popRef: React.RefObject<HTMLDivElement | null>
  readonly children: React.ReactNode
}): React.JSX.Element {
  useLayoutEffect(() => {
    const el = popRef.current
    if (!el) return
    const r = anchor.getBoundingClientRect()
    const left = Math.min(Math.max(8, r.left), window.innerWidth - el.offsetWidth - 8)
    let top = r.bottom + 4
    if (top + el.offsetHeight > window.innerHeight - 8)
      top = Math.max(8, r.top - el.offsetHeight - 4)
    el.style.left = `${left}px`
    el.style.top = `${top}px`
  }, [anchor, popRef])
  return createPortal(
    <div ref={popRef} className="sheets-color-pop-portal">
      {children}
    </div>,
    document.body,
  )
}

export function ColorDropdown({
  label,
  'data-tip': tip,
  display,
  value,
  auto,
  disabled,
  portal,
  onPick,
}: {
  /// aria-label of the trigger
  readonly label: string
  readonly 'data-tip'?: string
  /// trigger content (glyph + color echo bar); omitted = plain color well
  readonly display?: React.ReactNode
  /// current color (#rrggbb)
  readonly value: string
  /// label of the automatic / no-fill entry; picking it emits null
  readonly auto?: string
  readonly disabled?: boolean
  /// render the panel in a body portal (for transformed/clipping hosts)
  readonly portal?: boolean
  readonly onPick: (hex: string | null) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  const popRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const onDown = (event: MouseEvent): void => {
      const target = event.target as Node
      if (!wrapRef.current?.contains(target) && !popRef.current?.contains(target)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    // capture phase: the grid canvas stops mousedown propagation, so bubble-phase listeners never fire
    window.addEventListener('pointerdown', onDown, true)
    window.addEventListener('keydown', onKey)
    const offChrome = window.desktopApi?.onChromePressed?.(() => setOpen(false))
    return () => {
      window.removeEventListener('pointerdown', onDown, true)
      window.removeEventListener('keydown', onKey)
      offChrome?.()
    }
  }, [open])
  return (
    <div ref={wrapRef} className="menu-select">
      <button
        type="button"
        className={display ? 'color-tool' : 'color-well'}
        data-tip={tip}
        aria-label={label}
        aria-haspopup="dialog"
        aria-expanded={open}
        disabled={disabled}
        style={display ? undefined : { background: value }}
        onClick={() => setOpen((v) => !v)}
      >
        {display}
      </button>
      {open &&
        (() => {
          const picker = (
            <ColorPicker
              className={portal ? undefined : 'sheets-color-pop'}
              value={value}
              strings={{
                auto,
                themeColors: t('appThemeColors'),
                standardColors: t('appStandardColors'),
                moreColors: t('appMoreColors'),
              }}
              onPick={(hex) => {
                onPick(hex ? hex.toLowerCase() : null)
                setOpen(false)
              }}
              moreInputProps={{
                // apply live without closing: unmounting the hidden input would
                // tear down the still-open native color dialog
                onChange: (e) => onPick(e.currentTarget.value.toLowerCase()),
              }}
            />
          )
          return portal && wrapRef.current ? (
            <PortalPop anchor={wrapRef.current} popRef={popRef}>
              {picker}
            </PortalPop>
          ) : (
            picker
          )
        })()}
    </div>
  )
}

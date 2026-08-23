/**
 * Compact color well for the format pane: a swatch trigger opening the shared
 * Word-style palette (@genoffice/ui ColorPicker), fixed + anchored below the
 * trigger so it escapes the pane's scroll clip and opens leftwards (the pane
 * hugs the window's right edge).
 *
 * Outside-click/Escape close only — no blur close, because the "More Colors"
 * native dialog blurs the window while it is open. Custom picks apply live
 * through onPick while that dialog stays open.
 */
import { useEffect, useRef, useState } from 'react'
import { ColorPicker } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import { armColorInput } from '../color-input'

export function ColorWell({
  value,
  label,
  onPick,
}: {
  /** current color (#rrggbb), painted on the trigger */
  readonly value: string
  readonly label: string
  readonly onPick: (hex: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLSpanElement>(null)
  useEffect(() => {
    if (!open) return
    const onDown = (event: MouseEvent): void => {
      if (!wrapRef.current?.contains(event.target as Node)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('pointerdown', onDown, true)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('pointerdown', onDown, true)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])
  return (
    <span ref={wrapRef} className="fp-color-well-wrap">
      <button
        type="button"
        className="fp-color-well"
        data-tip={label}
        aria-label={label}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        <span className="fp-color-well-swatch" style={{ background: value }} />
        <span className="fp-color-well-caret" aria-hidden="true">
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
        <ColorPicker
          className="fp-color-pop"
          value={value}
          strings={{
            themeColors: t('ribbonThemeColorsSection'),
            standardColors: t('ribbonStandardColors'),
            moreColors: t('ribbonMoreColors'),
          }}
          onPick={(hex) => {
            if (hex) onPick(hex)
            setOpen(false)
          }}
          moreInputProps={{
            // re-fires change even when the confirmed color equals the shown one
            onPointerDown: (e) => armColorInput(e.currentTarget),
            // apply live without closing: unmounting the hidden input would
            // tear down the still-open native color dialog
            onChange: (e) => onPick(e.currentTarget.value),
          }}
        />
      )}
    </span>
  )
}

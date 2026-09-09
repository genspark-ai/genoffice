import { useLayoutEffect, useRef, useState } from 'react'
import type { ReactNode, Ref } from 'react'
import { createPortal } from 'react-dom'
import { ColorPicker, useDismissablePopover } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'

const POP_WIDTH = 260

interface PopProps {
  /** current #rrggbb, '' when unset / transparent */
  value: string
  /** show the "None" button (emits null) */
  none?: boolean
  onPick: (hex: string | null) => void
  onClose: () => void
  anchor: HTMLElement | null
  /** the popover root, for the caller's outside-press guard */
  popRef?: Ref<HTMLDivElement>
  /** any value that changes when the anchor is repositioned by its owner (e.g. toolbar coordinates) */
  track?: unknown
  children?: ReactNode
}

/** the shared Word-style palette, portaled to <body> and fixed under `anchor`: it escapes the toolbar's
    and panel's stacking contexts and clipping, and follows the anchor on scroll / resize / `track` */
export function ColorPop(p: PopProps) {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ left: 0, top: 0 })
  useLayoutEffect(() => {
    const place = () => {
      const r = p.anchor?.getBoundingClientRect()
      const h = ref.current?.offsetHeight ?? 0
      if (!r) return
      const left = Math.max(4, Math.min(r.left, window.innerWidth - POP_WIDTH - 4))
      // below the anchor by default; flip above when the palette would run off the bottom
      const below = r.bottom + 4
      const top = below + h > window.innerHeight - 4 ? Math.max(4, r.top - 4 - h) : below
      setPos({ left, top })
    }
    place()
    window.addEventListener('resize', place)
    // capture: the style panel scrolls itself, not the window
    window.addEventListener('scroll', place, true)
    return () => {
      window.removeEventListener('resize', place)
      window.removeEventListener('scroll', place, true)
    }
    // `none` toggles the None row, so the height (and thus the flip) changes with it
  }, [p.anchor, p.none, p.track])

  return createPortal(
    <div
      ref={(el) => {
        ref.current = el
        if (typeof p.popRef === 'function') p.popRef(el)
        else if (p.popRef) p.popRef.current = el
      }}
      className="hx-color-pop"
      style={pos}
      onMouseDown={(e) => e.stopPropagation()}
    >
      {p.children}
      <ColorPicker
        value={p.value || null}
        strings={{
          auto: p.none ? t('colorNone') : undefined,
          themeColors: t('themeColors'),
          standardColors: t('standardColors'),
          moreColors: t('moreColors'),
        }}
        onPick={(hex) => {
          p.onPick(hex ? hex.toLowerCase() : null)
          p.onClose()
        }}
        moreInputProps={{
          // apply live without closing: unmounting the hidden input would tear down the open OS dialog
          onChange: (e) => p.onPick(e.currentTarget.value.toLowerCase()),
        }}
      />
    </div>,
    document.body,
  )
}

interface FieldProps {
  label: string
  value: string
  /** what the swatch paints when it differs from the pickable colour (a gradient layer) */
  swatch?: string
  none?: boolean
  onPick: (hex: string | null) => void
}

/** style-panel field: a wide swatch button that opens the shared palette */
export function ColorField(p: FieldProps) {
  const swatch = p.swatch ?? p.value
  const [open, setOpen] = useState(false)
  const btnRef = useRef<HTMLButtonElement>(null)
  const popRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(open, () => setOpen(false), {
    inside: () => [btnRef.current, popRef.current],
  })

  return (
    <>
      <label>
        <span>{p.label}</span>
        <button
          ref={btnRef}
          type="button"
          className="hx-color-field"
          aria-label={p.label}
          aria-expanded={open}
          onClick={() => setOpen((v) => !v)}
        >
          <span
            className={`hx-color-swatch${swatch ? '' : ' none'}`}
            style={swatch ? { background: swatch } : undefined}
          />
        </button>
      </label>
      {open && (
        <ColorPop
          popRef={popRef}
          anchor={btnRef.current}
          value={p.value}
          none={p.none}
          onPick={p.onPick}
          onClose={() => setOpen(false)}
        />
      )}
    </>
  )
}

import { useLayoutEffect, useRef } from 'react'
import type { ReactElement } from 'react'
import { ColorPicker } from '@genoffice/ui'
import { useI18n } from './i18n/locale'
import { isHexColor } from './color-presets'

interface ColorPickerPopoverProps {
  /** Current color as #rrggbb — highlights the matching swatch */
  value?: string
  /** Extra classes on the popover root (e.g. `rb-drop` when anchored to a ribbon trigger) */
  className?: string
  onPick: (hex: string) => void
  onClose: () => void
}

/** The single color picker used across the PDF app: the shared Word-style
    theme/standard palette plus a "More Colors…" native picker entry. Preset
    picks apply immediately and close; custom picks apply live while the native
    dialog is open. Callers own the open state and outside-click behavior. */
export function ColorPickerPopover({
  value,
  className,
  onPick,
  onClose,
}: ColorPickerPopoverProps): ReactElement {
  const { t } = useI18n()
  const rootRef = useRef<HTMLDivElement>(null)

  // Anchored/floating parents (ribbon anchor positioning, the text-edit
  // editor's page coordinates) land the popover on fractional pixels, which
  // smears every 1px hairline inside — snap it to whole device pixels
  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el) return
    el.style.transform = ''
    const dpr = window.devicePixelRatio || 1
    const r = el.getBoundingClientRect()
    const dx = Math.round(r.left * dpr) / dpr - r.left
    const dy = Math.round(r.top * dpr) / dpr - r.top
    if (dx || dy) el.style.transform = `translate(${dx}px, ${dy}px)`
  }, [])

  return (
    <div ref={rootRef} className={`pdf-color-popover${className ? ` ${className}` : ''}`}>
      <ColorPicker
        value={value && isHexColor(value) ? value : null}
        strings={{
          themeColors: t('themeColors'),
          standardColors: t('standardColors'),
          moreColors: t('moreColors'),
        }}
        onPick={(hex) => {
          if (!hex) return
          onPick(hex.toLowerCase())
          onClose()
        }}
        moreInputProps={{
          // Apply live without closing: the hidden native input lives inside
          // this popover, so unmounting it would tear down the open OS dialog
          onChange: (e) => onPick(e.currentTarget.value.toLowerCase()),
        }}
      />
    </div>
  )
}

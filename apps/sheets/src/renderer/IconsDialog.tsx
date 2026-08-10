import { createElement, useState } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import type { FluentIcon } from '@fluentui/react-icons'
import { ICON_CATALOG } from './icon-catalog'
import { useI18n } from './i18n/locale'

/// Excel's Insert → Icons: a categorized, searchable gallery. Icons are
/// vector components on screen; picking one rasterizes it to a transparent
/// PNG so it rides the same picture pipeline (journal + xl/media) as any
/// inserted image.

const COLORS = ['#333333', '#4472c4', '#ed7d31', '#70ad47', '#ffc000', '#c00000']
const RASTER_SIZE = 256
/// Anchor size the picture pipeline uses (~2 columns × 6 rows).
const INSERT_SIZE = 128

function rasterizeIcon(icon: FluentIcon, color: string): Promise<string> {
  const svg = renderToStaticMarkup(
    createElement(icon, {
      xmlns: 'http://www.w3.org/2000/svg',
      width: RASTER_SIZE,
      height: RASTER_SIZE,
      style: { color },
    } as never),
  )
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = RASTER_SIZE
      canvas.height = RASTER_SIZE
      const context = canvas.getContext('2d')
      if (!context) {
        reject(new Error('canvas 2d context unavailable'))
        return
      }
      context.drawImage(image, 0, 0, RASTER_SIZE, RASTER_SIZE)
      resolve(canvas.toDataURL('image/png'))
    }
    image.onerror = () => reject(new Error('icon rasterization failed'))
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  })
}

export function IconsDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (dataUrl: string, size: number, name: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [categoryIndex, setCategoryIndex] = useState(0)
  const [query, setQuery] = useState('')
  const [color, setColor] = useState(COLORS[0] ?? '#333333')
  const [busy, setBusy] = useState(false)

  const trimmed = query.trim().toLowerCase()
  const shown =
    trimmed === ''
      ? (ICON_CATALOG[categoryIndex]?.icons ?? [])
      : ICON_CATALOG.flatMap((category) =>
          category.icons.filter((entry) => entry.name.includes(trimmed)),
        )

  const pick = (icon: FluentIcon, name: string): void => {
    if (busy) return
    setBusy(true)
    rasterizeIcon(icon, color)
      .then((dataUrl) => {
        onInsert(dataUrl, INSERT_SIZE, name)
        onClose()
      })
      .catch(() => setBusy(false))
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog icons-dialog"
        role="dialog"
        aria-label={t('dlgIconsTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgIconsTitle')}</header>
        <div className="icons-toolbar">
          <input
            type="search"
            value={query}
            placeholder={t('dlgIconsSearch')}
            onChange={(event) => setQuery(event.target.value)}
          />
          <div className="icons-colors" role="radiogroup" aria-label={t('dlgIconsColor')}>
            {COLORS.map((candidate) => (
              <button
                key={candidate}
                type="button"
                role="radio"
                aria-checked={candidate === color}
                className={candidate === color ? 'active' : ''}
                style={{ background: candidate }}
                data-tip={candidate}
                aria-label={candidate}
                onClick={() => setColor(candidate)}
              />
            ))}
          </div>
        </div>
        {trimmed === '' && (
          <nav className="dialog-tabs">
            {ICON_CATALOG.map((category, index) => (
              <button
                key={category.labelKey}
                className={index === categoryIndex ? 'active' : ''}
                onClick={() => setCategoryIndex(index)}
              >
                {t(category.labelKey)}
              </button>
            ))}
          </nav>
        )}
        <section className="dialog-body">
          <div className="icons-grid" style={{ color }}>
            {shown.map((entry) => (
              <button
                key={entry.name}
                type="button"
                disabled={busy}
                data-tip={entry.name}
                aria-label={entry.name}
                onClick={() => pick(entry.icon, entry.name)}
              >
                <entry.icon />
              </button>
            ))}
          </div>
          {shown.length === 0 && <p className="dialog-note">{t('dlgIconsNoResults')}</p>}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}

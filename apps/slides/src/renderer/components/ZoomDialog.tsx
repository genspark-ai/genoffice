/**
 * Insert > Zoom picker (PowerPoint's Insert Summary / Section / Slide Zoom dialog): a grid of
 * thumbnails, click selects one, toggle-modifier click adds, Insert stays disabled until
 * something is picked. Section mode shows each section's first slide.
 */
import React, { useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { SectionInfo } from '../../shared/ipc'
import { useEscOverlay } from '../esc-overlay'
import { useI18n } from '../i18n/locale'
import { isToggleModifier } from '../platform-modifiers'
import { groupSections } from '../section-groups'
import { SlideThumb } from '../SlideThumb'
import type { ZoomMode } from '../zoom-actions'

const THUMB_W = 160

const TITLE_KEY = {
  summary: 'ribbonZoomSummary',
  section: 'ribbonZoomSection',
  slide: 'ribbonZoomSlide',
} as const

interface ZoomItem {
  /** Slide index (slide/summary) or groupSections() index (section) */
  key: number
  slide: RenderSlide
  caption: string
  disabled?: boolean
}

export function ZoomDialog({
  mode,
  slides,
  images,
  sections,
  currentSlide,
  onInsert,
  onClose,
}: {
  mode: ZoomMode
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  sections: SectionInfo[]
  currentSlide: number
  onInsert: (keys: number[]) => void
  onClose: () => void
}) {
  useEscOverlay(true, onClose)
  const { t } = useI18n()
  const groups = groupSections(sections, slides.length) ?? []
  const items: ZoomItem[] =
    mode === 'section'
      ? groups.map((g, k) => ({
          key: k,
          slide: slides[g.start]!,
          caption: t('ribbonZoomSectionItem', {
            n: g.start + 1,
            k: k + 1,
            name: g.id == null ? t('appSectionDefault') : g.name,
          }),
        }))
      : slides.map((slide, i) => ({
          key: i,
          slide,
          caption:
            t('ribbonSlideN', { n: i + 1 }) +
            (mode === 'slide' && i === currentSlide ? t('ribbonCurrentSlideSuffix') : ''),
          disabled: mode === 'slide' && i === currentSlide,
        }))
  const [selected, setSelected] = useState<number[]>(() =>
    mode === 'summary' ? groups.map((g) => g.start) : [],
  )

  const pick = (key: number, e: React.MouseEvent) => {
    if (isToggleModifier(e)) {
      setSelected((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]))
    } else setSelected([key])
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal zoom-dlg" onClick={(e) => e.stopPropagation()}>
        <h2>{t(TITLE_KEY[mode])}</h2>
        <div className="zoom-dlg-grid">
          {items.map((item) => (
            <button
              key={item.key}
              className={`zoom-dlg-item${selected.includes(item.key) ? ' selected' : ''}`}
              disabled={item.disabled}
              aria-pressed={selected.includes(item.key)}
              onClick={(e) => pick(item.key, e)}
            >
              <SlideThumb slide={item.slide} images={images} width={THUMB_W} />
              <span className="zoom-dlg-caption">{item.caption}</span>
            </button>
          ))}
        </div>
        <div className="zoom-dlg-footer">
          <span className="zoom-dlg-count">
            {t(mode === 'section' ? 'ribbonZoomSelectedSections' : 'ribbonZoomSelectedSlides', {
              n: selected.length,
            })}
          </span>
          <div className="modal-actions">
            <button onClick={onClose}>{t('ribbonCancel')}</button>
            <button
              className="primary"
              disabled={!selected.length}
              onClick={() => onInsert([...selected].sort((a, b) => a - b))}
            >
              {t('ribbonInsert')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

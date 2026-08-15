/**
 * The Insert tab's three dialogs: hyperlink / header & footer / equation.
 * Reuses SettingsModal's .modal-backdrop/.modal styles.
 */
import React, { useEffect, useState } from 'react'
import type { LinkTargetOp } from '../../shared/ipc'
import { EQUATION_GALLERY } from '../insert-presets'
import { useI18n } from '../i18n/locale'

// ── Hyperlink ─────────────────────────────────────────────────────────────

interface LinkDialogProps {
  /** Current link (for display; null = none) */
  initial: LinkTargetOp | null
  /** Document page count (page-jump dropdown) */
  slideCount: number
  currentSlide: number
  onApply: (target: LinkTargetOp | null) => void
  onClose: () => void
}

export function LinkDialog({
  initial,
  slideCount,
  currentSlide,
  onApply,
  onClose,
}: LinkDialogProps) {
  const { t } = useI18n()
  const [mode, setMode] = useState<'url' | 'slide'>(initial?.kind === 'slide' ? 'slide' : 'url')
  const [url, setUrl] = useState(initial?.kind === 'url' ? initial.url : 'https://')
  const [slideIndex, setSlideIndex] = useState(
    initial?.kind === 'slide' ? initial.slideIndex : (currentSlide + 1) % Math.max(slideCount, 1),
  )

  const apply = () => {
    if (mode === 'url') {
      const trimmed = url.trim()
      if (!trimmed || trimmed === 'https://') return
      onApply({ kind: 'url', url: trimmed })
    } else {
      onApply({ kind: 'slide', slideIndex })
    }
  }

  return (
    // data-keep-edit: opening over a text-edit session must not commit it — the run-level
    // link applies to the saved editor selection after the dialog closes
    <div className="modal-backdrop" data-keep-edit="" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgHyperlink')}</h2>
        <div className="dlg-radio-row">
          <label className="dlg-radio">
            <input type="radio" checked={mode === 'url'} onChange={() => setMode('url')} />
            {t('ribbonDlgWebAddress')}
          </label>
          <label className="dlg-radio">
            <input type="radio" checked={mode === 'slide'} onChange={() => setMode('slide')} />
            {t('ribbonDlgPlaceInDoc')}
          </label>
        </div>
        {mode === 'url' ? (
          <label>
            {t('ribbonDlgAddress')}
            <input
              type="text"
              value={url}
              autoFocus
              onChange={(e) => setUrl(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && apply()}
              placeholder="https://example.com"
            />
          </label>
        ) : (
          <label>
            {t('ribbonDlgGoTo')}
            <select value={slideIndex} onChange={(e) => setSlideIndex(Number(e.target.value))}>
              {Array.from({ length: slideCount }, (_, i) => (
                <option key={i} value={i}>
                  {t('ribbonSlideN', { n: i + 1 })}
                  {i === currentSlide ? t('ribbonCurrentSlideSuffix') : ''}
                </option>
              ))}
            </select>
          </label>
        )}
        <div className="modal-actions">
          {initial && (
            <button className="dlg-danger" onClick={() => onApply(null)}>
              {t('ribbonRemoveLink')}
            </button>
          )}
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="primary" onClick={apply}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

// ── Header & footer ───────────────────────────────────────────────────────

interface HeaderFooterDialogProps {
  onApply: (opts: {
    footer: string | null
    slideNum: boolean
    date: string | null
    dateAuto: boolean
  }) => void
  onClose: () => void
  /** Display: current page state */
  initial: { footer: string | null; slideNum: boolean; date: string | null }
}

export function HeaderFooterDialog({ initial, onApply, onClose }: HeaderFooterDialogProps) {
  const { t } = useI18n()
  const [dateOn, setDateOn] = useState(!!initial.date)
  const [dateAuto, setDateAuto] = useState(true)
  const [dateText, setDateText] = useState(initial.date ?? new Date().toLocaleDateString())
  const [numOn, setNumOn] = useState(initial.slideNum)
  const [footerOn, setFooterOn] = useState(!!initial.footer)
  const [footerText, setFooterText] = useState(initial.footer ?? '')

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgHeaderFooter')}</h2>
        <label className="dlg-check">
          <input type="checkbox" checked={dateOn} onChange={(e) => setDateOn(e.target.checked)} />
          {t('ribbonDlgDateTime')}
        </label>
        {dateOn && (
          <div className="dlg-indent">
            <label className="dlg-check">
              <input
                type="checkbox"
                checked={dateAuto}
                onChange={(e) => setDateAuto(e.target.checked)}
              />
              {t('ribbonDlgAutoUpdate')}
            </label>
            <input
              type="text"
              value={dateText}
              onChange={(e) => setDateText(e.target.value)}
              placeholder={t('ribbonDlgDateExample')}
            />
          </div>
        )}
        <label className="dlg-check">
          <input type="checkbox" checked={numOn} onChange={(e) => setNumOn(e.target.checked)} />
          {t('ribbonDlgSlideNum')}
        </label>
        <label className="dlg-check">
          <input
            type="checkbox"
            checked={footerOn}
            onChange={(e) => setFooterOn(e.target.checked)}
          />
          {t('ribbonDlgFooter')}
        </label>
        {footerOn && (
          <div className="dlg-indent">
            <input
              type="text"
              value={footerText}
              autoFocus
              onChange={(e) => setFooterText(e.target.value)}
              placeholder={t('ribbonDlgFooterPlaceholder')}
            />
          </div>
        )}
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button
            className="primary"
            onClick={() =>
              onApply({
                footer: footerOn && footerText.trim() ? footerText.trim() : null,
                slideNum: numOn,
                date: dateOn && dateText.trim() ? dateText.trim() : null,
                dateAuto,
              })
            }
          >
            {t('ribbonApplyToAllBtn')}
          </button>
        </div>
      </div>
    </div>
  )
}

// ── Equation ───────────────────────────────────────────────────────

interface EquationDialogProps {
  onInsert: (text: string) => void
  onClose: () => void
}

export function EquationDialog({ onInsert, onClose }: EquationDialogProps) {
  const { t } = useI18n()
  const [text, setText] = useState('')

  // Esc closes
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal modal-wide" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgInsertEquation')}</h2>
        <div className="eq-gallery">
          {EQUATION_GALLERY.map((eq) => (
            <button
              key={eq.text}
              className="eq-item"
              data-tip={eq.label}
              onClick={() => setText(eq.text)}
            >
              <span className="eq-preview">{eq.text}</span>
              <span className="eq-label">{eq.label}</span>
            </button>
          ))}
        </div>
        <label>
          {t('ribbonDlgEquationLabel')}
          <input
            type="text"
            value={text}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && text.trim() && onInsert(text.trim())}
            placeholder={t('ribbonDlgEquationPlaceholder')}
          />
        </label>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="primary" disabled={!text.trim()} onClick={() => onInsert(text.trim())}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

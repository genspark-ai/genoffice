import { useState } from 'react'
import type { ReactElement } from 'react'
import { DEFAULT_HEADER_FOOTER, DEFAULT_WATERMARK } from './stamps'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import type { TFunc } from './i18n/locale'
import { ColorPalette } from './ColorPalette'

const WM_COLORS = ['#d0342c', '#8a8a8a', '#2b66ff', '#217346']
const WM_COLOR_PRESETS = WM_COLORS.map((value) => ({ value }))

/** Watermark / header-footer config dialog; on confirm App generates stamps and marks unsaved changes */
export function StampDialog({
  t,
  onCancel,
  onApply,
}: {
  t: TFunc
  onCancel: () => void
  onApply: (wm: WatermarkConfig | null, hf: HeaderFooterConfig | null) => void
}): ReactElement {
  const [tab, setTab] = useState<'watermark' | 'hf'>('watermark')
  const [wm, setWm] = useState<WatermarkConfig>(DEFAULT_WATERMARK)
  const [hf, setHf] = useState<HeaderFooterConfig>(DEFAULT_HEADER_FOOTER)

  const hfUsed =
    hf.pageNumber ||
    [
      hf.headerLeft,
      hf.headerCenter,
      hf.headerRight,
      hf.footerLeft,
      hf.footerCenter,
      hf.footerRight,
    ].some((s) => s.trim())
  const canApply = wm.text.trim().length > 0 || hfUsed

  const field = (key: keyof HeaderFooterConfig, label: string): ReactElement => (
    <label className="pdf-field">
      <span>{label}</span>
      <input
        className="pdf-modal-input"
        value={String(hf[key])}
        onChange={(e) => setHf({ ...hf, [key]: e.target.value })}
      />
    </label>
  )

  return (
    <div className="pdf-modal-mask" onClick={onCancel}>
      <div className="pdf-modal pdf-modal-wide" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-modal-title">{t('stampTitle')}</div>
        <div className="pdf-sign-tabs">
          <button
            className={`pdf-sign-tab${tab === 'watermark' ? ' active' : ''}`}
            onClick={() => setTab('watermark')}
          >
            {t('watermark')}
          </button>
          <button
            className={`pdf-sign-tab${tab === 'hf' ? ' active' : ''}`}
            onClick={() => setTab('hf')}
          >
            {t('headerFooter')}
          </button>
        </div>

        {tab === 'watermark' ? (
          <>
            <label className="pdf-field">
              <span>{t('watermarkText')}</span>
              <input
                className="pdf-modal-input"
                value={wm.text}
                placeholder={t('watermarkPlaceholder')}
                autoFocus
                onChange={(e) => setWm({ ...wm, text: e.target.value })}
              />
            </label>
            <label className="pdf-field">
              <span>{t('watermarkAngle')}</span>
              <input
                type="range"
                min={-90}
                max={90}
                value={wm.angle}
                onChange={(e) => setWm({ ...wm, angle: Number(e.target.value) })}
              />
              <em>{wm.angle}°</em>
            </label>
            <label className="pdf-field">
              <span>{t('watermarkOpacity')}</span>
              <input
                type="range"
                min={5}
                max={100}
                value={Math.round(wm.opacity * 100)}
                onChange={(e) => setWm({ ...wm, opacity: Number(e.target.value) / 100 })}
              />
              <em>{Math.round(wm.opacity * 100)}%</em>
            </label>
            <label className="pdf-field">
              <span>{t('watermarkSize')}</span>
              <input
                type="range"
                min={4}
                max={22}
                value={Math.round(wm.sizeRatio * 100)}
                onChange={(e) => setWm({ ...wm, sizeRatio: Number(e.target.value) / 100 })}
              />
              <em>{Math.round(wm.sizeRatio * 100)}%</em>
            </label>
            <div className="pdf-field">
              <span>{t('drawColor')}</span>
              <ColorPalette
                value={wm.color}
                presets={WM_COLOR_PRESETS}
                moreColorsLabel={t('moreColors')}
                onChange={(value) => setWm({ ...wm, color: value })}
              />
            </div>
            <div
              className="pdf-wm-preview"
              style={{ color: wm.color, opacity: Math.max(wm.opacity, 0.25) }}
            >
              <span style={{ transform: `rotate(${-wm.angle}deg)` }}>
                {wm.text || t('watermarkPlaceholder')}
              </span>
            </div>
          </>
        ) : (
          <>
            <div className="pdf-field-grid">
              {field('headerLeft', t('hfHeaderLeft'))}
              {field('headerCenter', t('hfHeaderCenter'))}
              {field('headerRight', t('hfHeaderRight'))}
              {field('footerLeft', t('hfFooterLeft'))}
              {field('footerCenter', t('hfFooterCenter'))}
              {field('footerRight', t('hfFooterRight'))}
            </div>
            <label className="pdf-field">
              <input
                type="checkbox"
                checked={hf.pageNumber}
                onChange={(e) => setHf({ ...hf, pageNumber: e.target.checked })}
              />
              <span>{t('hfPageNumber')}</span>
            </label>
            {hf.pageNumber && (
              <label className="pdf-field">
                <span>{t('hfStartAt')}</span>
                <input
                  className="pdf-modal-input pdf-input-sm"
                  type="number"
                  min={1}
                  value={hf.startAt}
                  onChange={(e) =>
                    setHf({ ...hf, startAt: Math.max(1, Number(e.target.value) || 1) })
                  }
                />
              </label>
            )}
            <div className="pdf-sign-hint">{t('hfTokenHint')}</div>
          </>
        )}

        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button
            className="pdf-modal-btn primary"
            disabled={!canApply}
            onClick={() => onApply(wm.text.trim() ? wm : null, hfUsed ? hf : null)}
          >
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}

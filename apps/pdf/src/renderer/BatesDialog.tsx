import { useState } from 'react'
import type { TFunc } from './i18n/locale'
import {
  BATES_MAX_FONT,
  BATES_MIN_FONT,
  formatBates,
  validateBates,
  type BatesConfig,
  type BatesPosition,
} from './bates'

const POSITIONS: BatesPosition[] = [
  'top-left',
  'top-center',
  'top-right',
  'bottom-left',
  'bottom-center',
  'bottom-right',
]

const POSITION_LABEL: Record<
  BatesPosition,
  | 'batesPosTopLeft'
  | 'batesPosTopCenter'
  | 'batesPosTopRight'
  | 'batesPosBottomLeft'
  | 'batesPosBottomCenter'
  | 'batesPosBottomRight'
> = {
  'top-left': 'batesPosTopLeft',
  'top-center': 'batesPosTopCenter',
  'top-right': 'batesPosTopRight',
  'bottom-left': 'batesPosBottomLeft',
  'bottom-center': 'batesPosBottomCenter',
  'bottom-right': 'batesPosBottomRight',
}

const COLOR_CHOICES: [number, number, number][] = [
  [0.1, 0.1, 0.12],
  [0.86, 0.22, 0.18],
  [0.17, 0.4, 1],
  [0.1, 0.55, 0.28],
]

const css = (rgb: [number, number, number]): string =>
  `rgb(${rgb.map((v) => Math.round(v * 255)).join(' ')})`

/**
 * Review → Bates: sequential page identifiers through the text-insert pipeline.
 * The right column previews where the number lands — a top strip and a bottom
 * strip over a blank page outline, aligned as the chosen position aligns. The
 * dialog hands back the configuration only; the caller builds the inserts from
 * the real page geometry.
 */
export function BatesDialog({
  t,
  pageCount,
  onApply,
  onClose,
}: {
  t: TFunc
  pageCount: number
  onApply: (config: BatesConfig) => void
  onClose: () => void
}): React.JSX.Element {
  const [prefix, setPrefix] = useState('')
  const [start, setStart] = useState('1')
  const [digits, setDigits] = useState('6')
  const [fontSize, setFontSize] = useState('10')
  const [color, setColor] = useState<[number, number, number]>(COLOR_CHOICES[0])
  const [position, setPosition] = useState<BatesPosition>('bottom-right')
  const [firstPage, setFirstPage] = useState('1')
  const [lastPage, setLastPage] = useState('')
  const [error, setError] = useState<string | null>(null)

  const cfg: BatesConfig = {
    prefix,
    start: Number(start),
    digits: Number(digits),
    fontSize: Number(fontSize),
    color,
    position,
    firstPage: Number(firstPage),
    lastPage: lastPage.trim() === '' ? null : Number(lastPage),
  }
  const invalid = validateBates(cfg)
  const previewFont = Math.min(Math.max(Number(fontSize) || 10, 6), 20)
  const sample =
    formatBates(cfg.prefix, cfg.start, cfg.digits) +
    (pageCount > 1 ? ` … ${formatBates(cfg.prefix, cfg.start + pageCount - 1, cfg.digits)}` : '')
  const previewAlign = position.endsWith('center')
    ? 'center'
    : position.endsWith('right')
      ? 'flex-end'
      : 'flex-start'
  const previewTextAlign = position.endsWith('center')
    ? ('center' as const)
    : position.endsWith('right')
      ? ('right' as const)
      : ('left' as const)

  const apply = (): void => {
    if (invalid) {
      setError(t(invalid))
      return
    }
    onApply(cfg)
  }

  const numberInput = (
    value: string,
    onChange: (v: string) => void,
    props: { min?: number; max?: number; placeholder?: string } = {},
  ): React.JSX.Element => (
    <input
      type="number"
      className="pdf-modal-input"
      value={value}
      min={props.min}
      max={props.max}
      placeholder={props.placeholder}
      onChange={(e) => {
        onChange(e.target.value)
        setError(null)
      }}
    />
  )

  return (
    <div className="pdf-modal-mask" onClick={onClose}>
      <div
        className="pdf-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('batesTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('batesTitle')}</div>
        <div className="pdf-modal-hint">{t('batesHint')}</div>

        <div className="pdf-field">
          <label className="pdf-perm-row">
            <span>{t('batesPrefix')}</span>
            <input
              type="text"
              className="pdf-modal-input"
              value={prefix}
              placeholder="ACME-"
              onChange={(e) => {
                setPrefix(e.target.value)
                setError(null)
              }}
            />
          </label>
          <label className="pdf-perm-row">
            <span>{t('batesStart')}</span>
            {numberInput(start, setStart, { min: 0 })}
          </label>
          <label className="pdf-perm-row">
            <span>{t('batesDigits')}</span>
            {numberInput(digits, setDigits, { min: 1, max: 15 })}
          </label>
          <label className="pdf-perm-row">
            <span>{t('batesFontSize')}</span>
            {numberInput(fontSize, setFontSize, { min: BATES_MIN_FONT, max: BATES_MAX_FONT })}
          </label>
        </div>

        <div className="pdf-field">
          <div className="pdf-modal-hint">{t('batesColor')}</div>
          <div className="pdf-perm-row">
            {COLOR_CHOICES.map((choice) => (
              <button
                key={choice.join()}
                type="button"
                aria-label={t('batesColor')}
                onClick={() => setColor(choice)}
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 11,
                  background: css(choice),
                  border: color === choice ? '2px solid var(--accent)' : '1px solid var(--border)',
                }}
              />
            ))}
          </div>
        </div>

        <div className="pdf-field">
          <div className="pdf-modal-hint">{t('batesPosition')}</div>
          <div className="pdf-bates-positions">
            {POSITIONS.map((candidate) => (
              <label key={candidate} className="pdf-perm-row">
                <input
                  type="radio"
                  name="bates-position"
                  checked={position === candidate}
                  onChange={() => setPosition(candidate)}
                />
                <span>{t(POSITION_LABEL[candidate])}</span>
              </label>
            ))}
          </div>
        </div>

        <div className="pdf-field">
          <label className="pdf-perm-row">
            <span>{t('batesFirstPage')}</span>
            {numberInput(firstPage, setFirstPage, { min: 1 })}
          </label>
          <label className="pdf-perm-row">
            <span>{t('batesLastPage')}</span>
            {numberInput(lastPage, setLastPage, { placeholder: t('batesLastPageHint') })}
          </label>
        </div>

        <div className="pdf-field">
          <div className="pdf-modal-hint">{t('batesPreview')}</div>
          <div className="pdf-bates-preview">
            <div
              className="pdf-bates-preview-strip"
              style={{
                alignSelf: previewAlign,
                color: css(color),
                fontSize: previewFont,
                textAlign: previewTextAlign,
              }}
            >
              {sample}
            </div>
            <div className="pdf-bates-preview-page" />
            <div
              className="pdf-bates-preview-strip"
              style={{
                alignSelf: previewAlign,
                color: css(color),
                fontSize: previewFont,
                textAlign: previewTextAlign,
              }}
            >
              {sample}
            </div>
          </div>
        </div>

        {error && (
          <div className="pdf-field">
            <div className="pdf-cert-error" role="alert">
              {error}
            </div>
          </div>
        )}

        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" onClick={apply}>
            {t('batesApply')}
          </button>
        </div>
      </div>
    </div>
  )
}

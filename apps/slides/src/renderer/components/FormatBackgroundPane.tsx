/**
 * Format Background pane (PowerPoint's "Format Background"): solid / gradient / picture
 * background for the current slide, a "hide background graphics" toggle, plus
 * apply-to-all and reset-background. Shares the right dock with the other panes.
 *
 * Fill edits apply live to the current slide (like PowerPoint); "apply to all"
 * re-applies the current slide's background to every slide.
 */
import { useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import { useI18n } from '../i18n/locale'
import { ColorWell } from './ColorWell'
import { IconSidebarCollapse } from './icons'

/** Background op without the fields App fills in (slideIndex / sourceSlideIndex / fitWidthPx). */
export type BgPaneOp =
  | { kind: 'solid'; color: string }
  | { kind: 'gradient'; from: string; to: string; angleDeg?: number; radial?: boolean }
  | { kind: 'image'; mode: 'stretch' | 'tile'; pick?: boolean }
  | { kind: 'reset' }
  | { kind: 'hideGraphics'; hidden: boolean }

interface Props {
  slide: RenderSlide | undefined
  /** Dispatch a background edit; allSlides=true targets every slide */
  onApply: (op: BgPaneOp, allSlides?: boolean) => void
  onCollapse: () => void
}

type FillMode = 'solid' | 'gradient' | 'image'

const GRADIENT_DIRECTIONS = [
  ['→', 0, false],
  ['↓', 90, false],
  ['↘', 45, false],
  ['◎', 0, true],
] as const

export function FormatBackgroundPane({ slide, onApply, onCollapse }: Props) {
  const { t } = useI18n()
  const bg = slide?.background
  const currentMode: FillMode =
    bg?.kind === 'gradient' ? 'gradient' : bg?.kind === 'image' ? 'image' : 'solid'
  // The tab the user picked (may differ from the slide's fill until an edit applies)
  const [pickedMode, setPickedMode] = useState<FillMode | null>(null)
  const mode = pickedMode ?? currentMode

  const solidColor = bg?.kind === 'solid' ? toHex6(bg.color) : '#ffffff'
  const gradStops = bg?.kind === 'gradient' ? bg.stops : null
  const [gradFromPick, setGradFrom] = useState<string | null>(null)
  const [gradToPick, setGradTo] = useState<string | null>(null)
  const gradFrom = gradFromPick ?? (gradStops ? toHex6(gradStops[0]!.color) : '#4472c4')
  const gradTo = gradToPick ?? (gradStops ? toHex6(gradStops.at(-1)!.color) : '#ffffff')
  const gradAngle = bg?.kind === 'gradient' ? Math.round(bg.angleDeg) : 0
  const gradRadial = bg?.kind === 'gradient' && !!bg.radial
  const imageMode = bg?.kind === 'image' ? (bg.mode ?? 'stretch') : 'stretch'

  // The native color dialog fires change repeatedly while dragging; debounce before IPC
  const timer = useRef<number | null>(null)
  const debounced = (op: BgPaneOp) => {
    if (timer.current) window.clearTimeout(timer.current)
    timer.current = window.setTimeout(() => onApply(op), 200)
  }

  const gradientOp = (from: string, to: string, angleDeg: number, radial: boolean): BgPaneOp => ({
    kind: 'gradient',
    from,
    to,
    angleDeg,
    ...(radial ? { radial: true } : {}),
  })

  /** Re-apply the current slide's background fill (for "apply to all"). */
  const currentFillOp = (): BgPaneOp | null => {
    if (!bg) return null
    if (bg.kind === 'solid') return { kind: 'solid', color: toHex6(bg.color) }
    if (bg.kind === 'gradient') return gradientOp(gradFrom, gradTo, gradAngle, gradRadial)
    if (bg.kind === 'image') return { kind: 'image', mode: imageMode, pick: false }
    return null
  }

  const hasDoc = !!slide

  return (
    <aside className="format-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneBgTitle')}</span>
        <div className="ai-panel-header-actions">
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('paneFormatClose')}
            aria-label={t('paneFormatClose')}
          >
            <IconSidebarCollapse size={15} />
          </button>
        </div>
      </div>

      {!hasDoc ? (
        <div className="fp-empty">{t('paneBgEmpty')}</div>
      ) : (
        <div className="fp-body">
          <div className="fp-section">{t('paneFormatFill')}</div>
          <div className="fp-row">
            <button
              className={`fp-btn fp-btn-wrap ${mode === 'solid' ? 'active' : ''}`}
              onClick={() => {
                setPickedMode('solid')
                if (currentMode !== 'solid') onApply({ kind: 'solid', color: solidColor })
              }}
            >
              {t('paneBgSolid')}
            </button>
            <button
              className={`fp-btn fp-btn-wrap ${mode === 'gradient' ? 'active' : ''}`}
              onClick={() => {
                setPickedMode('gradient')
                if (currentMode !== 'gradient')
                  onApply(gradientOp(gradFrom, gradTo, gradAngle, gradRadial))
              }}
            >
              {t('paneBgGradient')}
            </button>
            <button
              className={`fp-btn fp-btn-wrap ${mode === 'image' ? 'active' : ''}`}
              onClick={() => {
                setPickedMode('image')
                if (currentMode !== 'image') onApply({ kind: 'image', mode: 'stretch', pick: true })
              }}
            >
              {t('paneBgImage')}
            </button>
          </div>

          {mode === 'solid' && (
            <div className="fp-row">
              <ColorWell
                value={solidColor}
                label={t('paneFormatSolidFill')}
                onPick={(hex) => debounced({ kind: 'solid', color: hex })}
              />
            </div>
          )}

          {mode === 'gradient' && (
            <div className="fp-row">
              <ColorWell
                value={gradFrom}
                label={t('paneFormatGradientFrom')}
                onPick={(hex) => {
                  setGradFrom(hex)
                  debounced(gradientOp(hex, gradTo, gradAngle, gradRadial))
                }}
              />
              <ColorWell
                value={gradTo}
                label={t('paneFormatGradientTo')}
                onPick={(hex) => {
                  setGradTo(hex)
                  debounced(gradientOp(gradFrom, hex, gradAngle, gradRadial))
                }}
              />
              {GRADIENT_DIRECTIONS.map(([label, angleDeg, radial]) => (
                <button
                  key={label}
                  className={`fp-btn ${
                    currentMode === 'gradient' &&
                    (radial ? gradRadial : !gradRadial && gradAngle === angleDeg)
                      ? 'active'
                      : ''
                  }`}
                  data-tip={radial ? t('paneFormatGradientRadial') : `${angleDeg}°`}
                  aria-label={radial ? t('paneFormatGradientRadial') : `${angleDeg}°`}
                  onClick={() => onApply(gradientOp(gradFrom, gradTo, angleDeg, radial))}
                >
                  {label}
                </button>
              ))}
            </div>
          )}

          {mode === 'image' && (
            <>
              {bg?.kind === 'image' && bg.dataUrl && (
                <div className="fp-row">
                  <img className="fp-bg-preview" src={bg.dataUrl} alt="" />
                </div>
              )}
              <div className="fp-row">
                <button
                  className="fp-btn"
                  onClick={() => onApply({ kind: 'image', mode: imageMode, pick: true })}
                >
                  {t('paneBgPickImage')}
                </button>
                {(
                  [
                    ['stretch', t('paneBgStretch')],
                    ['tile', t('paneBgTile')],
                  ] as const
                ).map(([m, label]) => (
                  <button
                    key={m}
                    className={`fp-btn ${bg?.kind === 'image' && imageMode === m ? 'active' : ''}`}
                    disabled={bg?.kind !== 'image'}
                    onClick={() => onApply({ kind: 'image', mode: m, pick: false })}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </>
          )}

          <div className="fp-row">
            <label className="fp-check">
              <input
                type="checkbox"
                checked={!!slide?.bgGraphicsHidden}
                onChange={(e) => onApply({ kind: 'hideGraphics', hidden: e.target.checked })}
              />
              <span>{t('paneBgHideGraphics')}</span>
            </label>
          </div>

          <div className="fp-section">{t('paneFormatActions')}</div>
          <div className="fp-row">
            <button
              className="fp-btn"
              disabled={!currentFillOp()}
              onClick={() => {
                const op = currentFillOp()
                if (op) onApply(op, true)
              }}
            >
              {t('ribbonApplyToAll')}
            </button>
            <button
              className="fp-btn"
              disabled={!slide?.bgOwn}
              data-tip={t('paneBgResetTip')}
              onClick={() => {
                setPickedMode(null)
                onApply({ kind: 'reset' })
              }}
            >
              {t('paneBgReset')}
            </button>
          </div>
        </div>
      )}
    </aside>
  )
}

function toHex6(c: string): string {
  const m = /^#?([0-9a-fA-F]{6})/.exec(c)
  return m ? `#${m[1]!.toLowerCase()}` : '#ffffff'
}

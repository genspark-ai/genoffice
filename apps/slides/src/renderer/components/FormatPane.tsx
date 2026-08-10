/**
 * Format pane (a trimmed-down PowerPoint Format Pane): position/size/rotation/fill of the
 * selected element. Shares the right dock area with the AI panel, mutually exclusive. Inputs
 * commit on blur/Enter; external changes (dragging etc.) sync default values by remounting
 * inputs via key.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { PictureRenderNode, RenderNode, ShapeRenderNode } from '@genoffice/pptx-render'
import type { GradientFillSpec, LinkTargetOp } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { armColorInput } from '../color-input'
import { IconSidebarCollapse } from './icons'

interface Props {
  node: RenderNode | null
  onTransform: (
    sourceId: string,
    box: { x: number; y: number; w: number; h: number; rotationDeg: number },
  ) => void
  onFill: (sourceId: string, fill: string | GradientFillSpec) => void
  /** Shape picture fill (main process opens the image picker dialog) */
  onImageFill?: (sourceId: string) => void
  /** Text box vertical alignment */
  onTextAnchor?: (sourceId: string, anchor: 'top' | 'middle' | 'bottom') => void
  onStroke: (
    sourceId: string,
    stroke: { color: string; widthPt: number; dash?: string } | null,
  ) => void
  onDelete: (sourceId: string) => void
  onCollapse: () => void
  /** Picture: enter crop mode */
  onPictureCrop?: () => void
  /** Picture: enter cutout (background removal) mode */
  onPictureCutout?: () => void
  /** Whether the selected picture supports background removal (audio/video poster frames etc. don't) */
  pictureCanCutout?: boolean
  /** Element hyperlink (null = none) */
  link?: LinkTargetOp | null
  /** Open the hyperlink dialog for the selected element */
  onOpenLink?: () => void
  /** Chart: current data + colors (per-point color editing) */
  chartData?: {
    kind: string
    categories: string[]
    series: Array<{ name: string; values: number[] }>
    seriesColors: Array<string | undefined>
    pointColors: Array<Array<string | undefined> | undefined>
  } | null
  onChartPointColor?: (seriesIdx: number, pointIdx: number, color: string) => void
}

const TRANSFORMABLE = new Set(['shape', 'text', 'picture', 'group', 'table', 'chart'])

const TRANSPARENCY_PRESETS = [0, 15, 30, 50, 65, 80, 95]

/** Office default series palette (mirrors pptx-render build-chart's PALETTE) */
const CHART_PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

/** OOXML prstDash presets with language-neutral glyph labels. */
const DASH_PRESETS: Array<[string, string]> = [
  ['solid', '───────'],
  ['sysDot', '·······'],
  ['dot', '• • • •'],
  ['dash', '– – – –'],
  ['lgDash', '— — —'],
  ['dashDot', '– · – ·'],
  ['lgDashDot', '— · — ·'],
  ['lgDashDotDot', '— · · — · ·'],
]

export function FormatPane({
  node,
  onTransform,
  onFill,
  onImageFill,
  onTextAnchor,
  onStroke,
  onDelete,
  onCollapse,
  onPictureCrop,
  onPictureCutout,
  pictureCanCutout,
  link,
  onOpenLink,
  chartData,
  onChartPointColor,
}: Props) {
  const { t } = useI18n()
  // The color picker fires change repeatedly while dragging; debounce before IPC
  const fillTimer = useRef<number | null>(null)
  const debouncedFill = (sourceId: string, value: string) => {
    if (fillTimer.current) window.clearTimeout(fillTimer.current)
    fillTimer.current = window.setTimeout(() => onFill(sourceId, value), 200)
  }

  const strokeTimer = useRef<number | null>(null)
  const pointColorTimer = useRef<number | null>(null)
  const debouncedPointColor = (si: number, pi: number, value: string) => {
    if (pointColorTimer.current) window.clearTimeout(pointColorTimer.current)
    pointColorTimer.current = window.setTimeout(() => onChartPointColor?.(si, pi, value), 200)
  }
  // The two gradient edit colors (stashed locally before applying)
  const [gradFrom, setGradFrom] = useState('#4472C4')
  const [gradTo, setGradTo] = useState('#FFFFFF')

  const box = node?.box
  const canTransform = !!node && TRANSFORMABLE.has(node.type)
  const shape =
    node && (node.type === 'shape' || node.type === 'text') ? (node as ShapeRenderNode) : null
  const pic = node && node.type === 'picture' ? (node as PictureRenderNode) : null
  const fillColor = shape?.fill.kind === 'solid' ? toHex6(shape.fill.color) : null
  const fillAlpha = shape?.fill.kind === 'solid' ? alphaOf(shape.fill.color) : 255
  // 0..100 transparency shown in the dropdown (0 = opaque)
  const fillTransparency = Math.round(((255 - fillAlpha) / 255) * 100)
  const stroke = (shape ?? pic)?.stroke
  const strokeWidthPt = stroke ? Math.max(0.5, Math.round(stroke.widthPt * 2) / 2) : 1
  const strokeColor = stroke ? toHex6(stroke.color) : '#000000'
  const strokeDash = stroke?.dashPreset ?? 'solid'

  /** Solid fill color + transparency merged into one #RRGGBB(AA) value. */
  const fillValue = (color: string, transparencyPct: number) => {
    const alpha = Math.round(((100 - transparencyPct) / 100) * 255)
    return alpha >= 255 && fillAlpha >= 255
      ? color
      : `${color}${Math.max(0, alpha).toString(16).padStart(2, '0')}`
  }

  // Latest intended stroke: each input contributes only its own dimension, so a debounced color
  // commit can't overwrite a width committed meanwhile (and vice versa)
  const strokeDraft = useRef<{ id: string; color: string; widthPt: number; dash: string } | null>(
    null,
  )
  useEffect(() => {
    if (!strokeTimer.current) strokeDraft.current = null
  }, [strokeColor, strokeWidthPt, strokeDash, node?.sourceId])
  const commitStroke = (
    sourceId: string,
    patch: Partial<{ color: string; widthPt: number; dash: string }>,
    immediate = false,
  ) => {
    if (strokeTimer.current) window.clearTimeout(strokeTimer.current)
    const prev =
      strokeDraft.current?.id === sourceId
        ? strokeDraft.current
        : { id: sourceId, color: strokeColor, widthPt: strokeWidthPt, dash: strokeDash }
    const draft = { ...prev, ...patch }
    strokeDraft.current = draft
    const fire = () => {
      strokeTimer.current = null
      onStroke(sourceId, { color: draft.color, widthPt: draft.widthPt, dash: draft.dash })
    }
    if (immediate) fire()
    else strokeTimer.current = window.setTimeout(fire, 200)
  }
  const clearStroke = (sourceId: string) => {
    if (strokeTimer.current) window.clearTimeout(strokeTimer.current)
    strokeTimer.current = null
    strokeDraft.current = null
    onStroke(sourceId, null)
  }

  const commit = (
    patch: Partial<{ x: number; y: number; w: number; h: number; rotationDeg: number }>,
  ) => {
    if (!node || !box) return
    onTransform(node.sourceId, {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      rotationDeg: box.rotationDeg,
      ...patch,
    })
  }

  const numField = (label: string, value: number, apply: (v: number) => void, min?: number) => (
    <label className="fp-field" key={label}>
      <span>{label}</span>
      <input
        key={`${node?.sourceId}:${value}`}
        type="number"
        defaultValue={Math.round(value)}
        min={min}
        onKeyDown={(e) => {
          if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
        }}
        onBlur={(e) => {
          const v = Number(e.target.value)
          if (!Number.isNaN(v) && Math.round(v) !== Math.round(value)) apply(v)
        }}
      />
    </label>
  )

  return (
    <aside className="format-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneFormatTitle')}</span>
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

      {!node ? (
        <div className="fp-empty">{t('paneFormatEmpty')}</div>
      ) : (
        <div className="fp-body">
          <div className="fp-section-title">
            {node.type === 'picture'
              ? t('paneFormatPicture')
              : node.type === 'group'
                ? t('paneFormatGroup')
                : node.type === 'text'
                  ? t('paneFormatTextBox')
                  : node.type === 'table'
                    ? t('ribbonGroupTable')
                    : node.type === 'chart'
                      ? t('ribbonChart')
                      : t('paneFormatShape')}
          </div>

          {canTransform && box && (
            <>
              <div className="fp-section">{t('paneFormatPosSize')}</div>
              <div className="fp-grid">
                {numField('X', box.x, (v) => commit({ x: v }))}
                {numField('Y', box.y, (v) => commit({ y: v }))}
                {numField(t('paneFormatW'), box.w, (v) => commit({ w: v }), 1)}
                {numField(t('paneFormatH'), box.h, (v) => commit({ h: v }), 1)}
              </div>
              <div className="fp-grid">
                {numField(t('paneFormatRotation'), box.rotationDeg, (v) =>
                  commit({ rotationDeg: v }),
                )}
              </div>
            </>
          )}

          {node.type === 'picture' && (
            <>
              <div className="fp-section">{t('paneFormatPicture')}</div>
              <div className="fp-row">
                <button className="fp-btn" onClick={() => onPictureCrop?.()}>
                  {t('paneFormatCrop')}
                </button>
                <button
                  className="fp-btn"
                  disabled={!pictureCanCutout}
                  data-tip={pictureCanCutout ? t('paneFormatCutoutTip') : t('paneFormatCutoutNA')}
                  onClick={() => onPictureCutout?.()}
                >
                  {t('paneCutoutTitle')}
                </button>
              </div>
            </>
          )}

          {shape && (
            <>
              <div className="fp-section">{t('paneFormatFill')}</div>
              <div className="fp-row">
                <input
                  key={`${node.sourceId}:${fillColor ?? 'none'}`}
                  type="color"
                  className="fp-color"
                  defaultValue={fillColor ?? '#ffffff'}
                  onPointerDown={(e) => armColorInput(e.currentTarget)}
                  onChange={(e) =>
                    debouncedFill(node.sourceId, fillValue(e.target.value, fillTransparency))
                  }
                  title={t('paneFormatSolidFill')}
                />
                <button
                  className={`fp-btn ${shape.fill.kind === 'none' ? 'active' : ''}`}
                  onClick={() => onFill(node.sourceId, 'none')}
                >
                  {t('paneFormatNoFill')}
                </button>
                {onImageFill && (
                  <button className="fp-btn" onClick={() => onImageFill(node.sourceId)}>
                    {t('paneFormatImageFill')}
                  </button>
                )}
              </div>
              <div className="fp-row">
                <label className="fp-field" style={{ flex: 1 }}>
                  <span>{t('ribbonTransparency')}</span>
                  <select
                    key={`${node.sourceId}:fa:${fillTransparency}`}
                    defaultValue={fillTransparency}
                    disabled={shape.fill.kind !== 'solid'}
                    onChange={(e) =>
                      onFill(
                        node.sourceId,
                        fillValue(fillColor ?? '#ffffff', Number(e.target.value)),
                      )
                    }
                  >
                    {!TRANSPARENCY_PRESETS.includes(fillTransparency) && (
                      <option value={fillTransparency}>{fillTransparency}%</option>
                    )}
                    {TRANSPARENCY_PRESETS.map((pct) => (
                      <option key={pct} value={pct}>
                        {pct}%
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              {shape.text && onTextAnchor && (
                <>
                  <div className="fp-section">{t('paneFormatTextAnchor')}</div>
                  <div className="fp-row">
                    {(
                      [
                        ['top', t('paneFormatAnchorTop')],
                        ['middle', t('paneFormatAnchorMiddle')],
                        ['bottom', t('paneFormatAnchorBottom')],
                      ] as const
                    ).map(([k, label]) => (
                      <button
                        key={k}
                        className={`fp-btn ${(shape.text?.anchor ?? 'top') === k ? 'active' : ''}`}
                        onClick={() => onTextAnchor(node.sourceId, k)}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </>
              )}
              <div className="fp-section">{t('paneFormatGradient')}</div>
              <div className="fp-row">
                <input
                  type="color"
                  className="fp-color"
                  value={gradFrom}
                  onChange={(e) => setGradFrom(e.target.value)}
                  data-tip={t('paneFormatGradientFrom')}
                />
                <input
                  type="color"
                  className="fp-color"
                  value={gradTo}
                  onChange={(e) => setGradTo(e.target.value)}
                  data-tip={t('paneFormatGradientTo')}
                />
                {(
                  [
                    ['→', 0, false],
                    ['↓', 90, false],
                    ['↘', 45, false],
                    ['◎', 0, true],
                  ] as const
                ).map(([label, angleDeg, radial]) => (
                  <button
                    key={label}
                    className="fp-btn"
                    data-tip={radial ? t('paneFormatGradientRadial') : `${angleDeg}°`}
                    aria-label={radial ? t('paneFormatGradientRadial') : `${angleDeg}°`}
                    onClick={() =>
                      onFill(node.sourceId, {
                        gradient: {
                          from: gradFrom,
                          to: gradTo,
                          angleDeg,
                          ...(radial ? { radial: true } : {}),
                        },
                      })
                    }
                  >
                    {label}
                  </button>
                ))}
              </div>
            </>
          )}

          {(shape || pic) && (
            <>
              <div className="fp-section">{t('paneFormatOutline')}</div>
              <div className="fp-row">
                <input
                  key={`${node.sourceId}:s:${strokeColor}`}
                  type="color"
                  className="fp-color"
                  defaultValue={strokeColor}
                  onPointerDown={(e) => armColorInput(e.currentTarget)}
                  onChange={(e) => commitStroke(node.sourceId, { color: e.target.value })}
                  data-tip={t('paneFormatOutlineColor')}
                />
                <label className="fp-field" style={{ flex: 1 }}>
                  <span>{t('paneFormatPt')}</span>
                  <input
                    key={`${node.sourceId}:sw:${strokeWidthPt}`}
                    type="number"
                    step={0.5}
                    min={0.5}
                    defaultValue={strokeWidthPt}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                    }}
                    onBlur={(e) => {
                      const v = Number(e.target.value)
                      if (!Number.isNaN(v) && v > 0)
                        commitStroke(node.sourceId, { widthPt: v }, true)
                    }}
                  />
                </label>
                <button
                  className={`fp-btn ${!stroke ? 'active' : ''}`}
                  onClick={() => clearStroke(node.sourceId)}
                >
                  {t('paneFormatNoOutline')}
                </button>
              </div>
              <div className="fp-row">
                <label className="fp-field" style={{ flex: 1 }}>
                  <span>{t('paneFormatDashStyle')}</span>
                  <select
                    key={`${node.sourceId}:sd:${strokeDash}`}
                    defaultValue={strokeDash}
                    onChange={(e) => commitStroke(node.sourceId, { dash: e.target.value }, true)}
                  >
                    {!DASH_PRESETS.some(([k]) => k === strokeDash) && (
                      <option value={strokeDash}>{strokeDash}</option>
                    )}
                    {DASH_PRESETS.map(([k, glyph]) => (
                      <option key={k} value={k}>
                        {glyph}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
            </>
          )}

          {node.type === 'chart' && chartData && onChartPointColor && (
            <>
              <div className="fp-section">{t('paneFormatChartPoints')}</div>
              {chartData.series.map((s, si) => (
                <React.Fragment key={si}>
                  {chartData.series.length > 1 && (
                    <div className="fp-chart-series">
                      {s.name || t('paneFormatChartSeriesN', { n: si + 1 })}
                    </div>
                  )}
                  {s.values.map((_, pi) => (
                    <div className="fp-row fp-chart-point" key={pi}>
                      <input
                        className="fp-color"
                        type="color"
                        value={toHex6(
                          chartData.pointColors[si]?.[pi] ??
                            (chartData.kind === 'pie'
                              ? CHART_PALETTE[pi % CHART_PALETTE.length]!
                              : (chartData.seriesColors[si] ??
                                CHART_PALETTE[si % CHART_PALETTE.length]!)),
                        )}
                        onPointerDown={(e) => armColorInput(e.currentTarget)}
                        onChange={(e) => debouncedPointColor(si, pi, e.target.value)}
                      />
                      <span className="fp-chart-point-label">
                        {chartData.categories[pi] || `#${pi + 1}`}
                      </span>
                    </div>
                  ))}
                </React.Fragment>
              ))}
            </>
          )}

          {onOpenLink && (
            <>
              <div className="fp-section">{t('paneFormatLink')}</div>
              <div className="fp-row">
                <span
                  className="fp-link-target"
                  data-tip={link?.kind === 'url' ? link.url : undefined}
                >
                  {link
                    ? link.kind === 'url'
                      ? link.url
                      : t('ribbonSlideN', { n: link.slideIndex + 1 })
                    : t('paneFormatLinkNone')}
                </span>
              </div>
              <button className="fp-btn" onClick={onOpenLink}>
                {t('paneFormatLinkSet')}
              </button>
            </>
          )}

          <div className="fp-section">{t('paneFormatActions')}</div>
          <button className="fp-btn fp-danger" onClick={() => onDelete(node.sourceId)}>
            {t('paneFormatDelete')}
          </button>
        </div>
      )}
    </aside>
  )
}

function toHex6(c: string): string {
  const m = /^#?([0-9a-fA-F]{6})/.exec(c)
  return m ? `#${m[1]!.toLowerCase()}` : '#ffffff'
}

/** Alpha byte of an #RRGGBBAA color (255 when absent). */
function alphaOf(c: string): number {
  const m = /^#?[0-9a-fA-F]{6}([0-9a-fA-F]{2})$/.exec(c)
  return m ? parseInt(m[1]!, 16) : 255
}

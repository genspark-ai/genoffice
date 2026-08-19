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
import { pathGradientCanvas } from '../konva-adapter'
import { ColorWell } from './ColorWell'
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

/** Path-gradient focus presets (fillToRect point), in PPT's gallery order. */
const FOCUS_POINTS: Array<[string, number, number]> = [
  ['↘', 1, 1],
  ['↙', 0, 1],
  ['◎', 0.5, 0.5],
  ['↗', 1, 0],
  ['↖', 0, 0],
]

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
  // PPT-style "Shape Options" / "Text Options" tabs; text tab only exists for text-bearing shapes
  const [paneTab, setPaneTab] = useState<'shape' | 'text'>('shape')
  // PPT-style second-level tabs: shape → fill&line / effects / size&props, text → fill&outline / effects / text box
  const [shapeSub, setShapeSub] = useState<'fill' | 'effects' | 'size'>('fill')
  const [textSub, setTextSub] = useState<'fill' | 'effects' | 'textbox'>('textbox')
  // PPT-style collapsible fill / line sections
  const [fillOpen, setFillOpen] = useState(true)
  const [lineOpen, setLineOpen] = useState(true)

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
  const fillKind = shape?.fill.kind ?? 'none'
  const gradFill = shape?.fill.kind === 'gradient' ? shape.fill : null

  // Editing an existing gradient starts from its actual stops
  useEffect(() => {
    if (gradFill?.stops.length) {
      setGradFrom(toHex6(gradFill.stops[0]!.color))
      setGradTo(toHex6(gradFill.stops[gradFill.stops.length - 1]!.color))
    }
  }, [node?.sourceId, gradFill])

  // Gradient stops being edited (model stops, or the two stashed colors before a gradient exists)
  const gradStops: Array<{ pos: number; color: string }> = gradFill
    ? gradFill.stops.map((s) => ({ pos: s.pos, color: s.color }))
    : [
        { pos: 0, color: gradFrom },
        { pos: 1, color: gradTo },
      ]
  const [stopIdx, setStopIdx] = useState(0)
  const selStopIdx = Math.min(stopIdx, gradStops.length - 1)
  const selStop = gradStops[selStopIdx]!
  const selStopTransparency = Math.round(((255 - alphaOf(selStop.color)) / 255) * 100)
  useEffect(() => setStopIdx(0), [node?.sourceId])

  // Current gradient shading type (linear, or the actual <a:path> kind: circle/rect/shape)
  const curPath: 'linear' | 'circle' | 'rect' | 'shape' = gradFill
    ? (gradFill.path ?? (gradFill.radial ? 'circle' : 'linear'))
    : 'linear'
  // Focus point of circle/rect path gradients (fillToRect; PPT's direction control for those types)
  const curCenter = gradFill?.center ?? { x: 0.5, y: 0.5 }

  /** Re-apply the gradient fill with one property changed (stops/type/direction commit immediately). */
  const applyGradient = (
    patch: Partial<{
      stops: Array<{ pos: number; color: string }>
      angleDeg: number
      path: 'linear' | 'circle' | 'rect' | 'shape'
      center: { x: number; y: number }
    }>,
  ) => {
    if (!node) return
    const stops = patch.stops ?? gradStops
    const path = patch.path ?? curPath
    onFill(node.sourceId, {
      gradient: {
        from: toHex6(stops[0]!.color),
        to: toHex6(stops[stops.length - 1]!.color),
        stops,
        angleDeg: patch.angleDeg ?? gradFill?.angleDeg ?? 90,
        ...(path !== 'linear' ? { path } : {}),
        // shape-path gradients have no focus in PPT; circle/rect keep or update theirs
        ...(path === 'circle' || path === 'rect' ? { center: patch.center ?? curCenter } : {}),
      },
    })
  }

  /** #RRGGBB + transparency% → #RRGGBB(AA) */
  const stopColor = (hex: string, transparencyPct: number) => {
    const a = Math.round(((100 - transparencyPct) / 100) * 255)
    return a >= 255 ? toHex6(hex) : `${toHex6(hex)}${a.toString(16).padStart(2, '0')}`
  }

  /** Patch one gradient stop; the list stays sorted by position and the moved stop stays selected. */
  const setStop = (i: number, patch: Partial<{ pos: number; color: string }>) => {
    const next = gradStops.map((s, j) => (j === i ? { ...s, ...patch } : s))
    const moved = next[i]!
    next.sort((a, b) => a.pos - b.pos)
    setStopIdx(next.indexOf(moved))
    applyGradient({ stops: next })
  }

  const addStop = () => {
    const a = gradStops[selStopIdx]!
    const b = gradStops[selStopIdx + 1] ?? gradStops[selStopIdx - 1] ?? a
    const added = { pos: (a.pos + b.pos) / 2, color: a.color }
    const next = [...gradStops, added].sort((x, y) => x.pos - y.pos)
    setStopIdx(next.indexOf(added))
    applyGradient({ stops: next })
  }

  const removeStop = () => {
    if (gradStops.length <= 2) return
    const next = gradStops.filter((_, j) => j !== selStopIdx)
    setStopIdx(Math.max(0, selStopIdx - 1))
    applyGradient({ stops: next })
  }

  /** The stop-transparency slider fires per pixel; debounce before IPC (same as fills) */
  const gradTimer = useRef<number | null>(null)
  const debouncedStopAlpha = (pct: number) => {
    if (gradTimer.current) window.clearTimeout(gradTimer.current)
    gradTimer.current = window.setTimeout(
      () => setStop(selStopIdx, { color: stopColor(selStop.color, pct) }),
      200,
    )
  }

  const curAngle = Math.round(((gradFill?.angleDeg ?? 90) % 360) * 10) / 10
  const normAngle = (v: number) => Math.round((((v % 360) + 360) % 360) * 10) / 10
  // Preview-tile colors: the gradient's own first/last stops
  const gradPrevFrom = toHex6(gradStops[0]!.color)
  const gradPrevTo = toHex6(gradStops[gradStops.length - 1]!.color)
  // Gradient-style tiles (PPT Mac-style): linear/radial previewed via CSS gradients; rect/shape via
  // the canvas renderer's own output so the tiles show the real metric (X seams / 45° inset)
  const styleTiles: Array<{ kind: 'linear' | 'circle' | 'rect' | 'shape'; css: string }> = (() => {
    const two = [
      { pos: 0, color: gradPrevFrom },
      { pos: 1, color: gradPrevTo },
    ]
    const mini = (kind: 'rect' | 'shape', mw: number, mh: number) =>
      pathGradientCanvas(kind, two, mw, mh, 0.5, 0.5)?.toDataURL()
    const rectUrl = mini('rect', 24, 24)
    const shapeUrl = mini('shape', 34, 22)
    return [
      { kind: 'linear', css: `linear-gradient(135deg, ${gradPrevFrom}, ${gradPrevTo})` },
      { kind: 'circle', css: `radial-gradient(circle, ${gradPrevFrom}, ${gradPrevTo})` },
      {
        kind: 'rect',
        css: rectUrl
          ? `url(${rectUrl}) center / 100% 100%`
          : `radial-gradient(circle, ${gradPrevFrom}, ${gradPrevTo})`,
      },
      {
        kind: 'shape',
        css: shapeUrl
          ? `url(${shapeUrl}) center / 100% 100%`
          : `radial-gradient(circle, ${gradPrevFrom}, ${gradPrevTo})`,
      },
    ]
  })()

  // Angle dial: the indicator ring follows the pointer imperatively (per-frame, no React
  // round-trip); model commits are throttled during the drag and finalized on release
  const dialTimer = useRef<number | null>(null)
  const onDialDown = (e: React.PointerEvent<HTMLButtonElement>) => {
    if (curPath !== 'linear') return
    const el = e.currentTarget
    const dot = el.querySelector<HTMLElement>('.fp-dial-dot')
    const r = el.getBoundingClientRect()
    const cxp = r.left + r.width / 2
    const cyp = r.top + r.height / 2
    el.setPointerCapture(e.pointerId)
    let lastCommit = 0
    let pendingAngle = curAngle
    const move = (ev: PointerEvent) => {
      const deg = (Math.atan2(ev.clientY - cyp, ev.clientX - cxp) * 180) / Math.PI
      pendingAngle = normAngle(deg)
      if (dot) dot.style.transform = `rotate(${pendingAngle}deg) translateX(6px)`
      const now = performance.now()
      if (dialTimer.current) window.clearTimeout(dialTimer.current)
      if (now - lastCommit > 100) {
        lastCommit = now
        applyGradient({ angleDeg: pendingAngle })
      } else {
        dialTimer.current = window.setTimeout(() => applyGradient({ angleDeg: pendingAngle }), 120)
      }
    }
    const up = () => {
      if (dialTimer.current) window.clearTimeout(dialTimer.current)
      applyGradient({ angleDeg: pendingAngle })
      el.removeEventListener('pointermove', move)
      el.removeEventListener('pointerup', up)
    }
    el.addEventListener('pointermove', move)
    el.addEventListener('pointerup', up)
  }

  // Focus popover, opened by clicking the radial/rect style tile (closes on outside pointerdown)
  const [dirOpenFor, setDirOpenFor] = useState<'circle' | 'rect' | null>(null)
  useEffect(() => {
    if (!dirOpenFor) return
    const close = () => setDirOpenFor(null)
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [dirOpenFor])

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

  /** PPT-style collapsible section header (chevron + label over a divider) */
  const secHeader = (label: string, open: boolean, toggle: () => void) => (
    <button type="button" className="fp-sec" aria-expanded={open} onClick={toggle}>
      <span className="fp-sec-caret">{open ? '▾' : '▸'}</span>
      {label}
    </button>
  )

  /** PPT-style fill/line mode radio row */
  const radioRow = (group: string, checked: boolean, label: string, pick: () => void) => (
    <label className="fp-radio" key={label}>
      <input type="radio" name={group} checked={checked} onChange={pick} />
      <span>{label}</span>
    </label>
  )

  const numField = (label: string, value: number, apply: (v: number) => void, min?: number) => (
    <label className="fp-prow" key={label}>
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

  const typeName = !node
    ? null
    : node.type === 'picture'
      ? t('paneFormatPicture')
      : node.type === 'group'
        ? t('paneFormatGroup')
        : node.type === 'text'
          ? t('paneFormatTextBox')
          : node.type === 'table'
            ? t('ribbonGroupTable')
            : node.type === 'chart'
              ? t('ribbonChart')
              : t('paneFormatShape')

  const hasTextTab = !!(shape?.text && onTextAnchor)
  const effTab = hasTextTab ? paneTab : 'shape'
  const sub = effTab === 'shape' ? shapeSub : textSub
  // The fill&line sub-tab has content only for fill/stroke/chart-color bearing nodes (not e.g. groups)
  const fillHasContent =
    !!shape || !!pic || !!(node?.type === 'chart' && chartData && onChartPointColor)

  return (
    <aside className="format-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">
          {typeName ? t('paneFormatTitleTyped', { type: typeName }) : t('paneFormatTitle')}
        </span>
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

      {node && hasTextTab && (
        <div className="fp-tabs">
          <button
            type="button"
            className={`fp-tab ${effTab === 'shape' ? 'active' : ''}`}
            onClick={() => setPaneTab('shape')}
          >
            {t('paneFormatTabShape')}
          </button>
          <button
            type="button"
            className={`fp-tab ${effTab === 'text' ? 'active' : ''}`}
            onClick={() => setPaneTab('text')}
          >
            {t('paneFormatTabText')}
          </button>
        </div>
      )}

      {node && (
        <div className="fp-subtabs">
          {(effTab === 'shape'
            ? ([
                ['fill', t('paneSubFillLine')],
                ['effects', t('paneSubEffects')],
                ['size', t('paneSubSizeProps')],
              ] as const)
            : ([
                ['fill', t('paneSubTextFill')],
                ['effects', t('paneSubEffects')],
                ['textbox', t('paneFormatTextBox')],
              ] as const)
          ).map(([k, label]) => (
            <button
              key={k}
              type="button"
              className={`fp-subtab ${sub === k ? 'active' : ''}`}
              onClick={() =>
                effTab === 'shape'
                  ? setShapeSub(k as 'fill' | 'effects' | 'size')
                  : setTextSub(k as 'fill' | 'effects' | 'textbox')
              }
            >
              {label}
            </button>
          ))}
        </div>
      )}

      {!node ? (
        <div className="fp-empty">{t('paneFormatEmpty')}</div>
      ) : (
        <div className="fp-body">
          {effTab === 'shape' && shapeSub === 'size' && canTransform && box && (
            <>
              <div className="fp-section">{t('paneFormatPosSize')}</div>
              {numField('X', box.x, (v) => commit({ x: v }))}
              {numField('Y', box.y, (v) => commit({ y: v }))}
              {numField(t('paneFormatW'), box.w, (v) => commit({ w: v }), 1)}
              {numField(t('paneFormatH'), box.h, (v) => commit({ h: v }), 1)}
              {numField(t('paneFormatRotation'), box.rotationDeg, (v) =>
                commit({ rotationDeg: v }),
              )}
            </>
          )}

          {effTab === 'shape' && shapeSub === 'size' && node.type === 'picture' && (
            <>
              <div className="fp-section">{t('paneFormatPicture')}</div>
              <div className="fp-prow fp-prow-end">
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

          {effTab === 'shape' && shapeSub === 'fill' && shape && (
            <>
              {secHeader(t('paneFormatFill'), fillOpen, () => setFillOpen((v) => !v))}
              {fillOpen && (
                <>
                  <div className="fp-radios">
                    {radioRow(
                      `fill-${node.sourceId}`,
                      fillKind === 'none',
                      t('paneFormatNoFill'),
                      () => onFill(node.sourceId, 'none'),
                    )}
                    {radioRow(
                      `fill-${node.sourceId}`,
                      fillKind === 'solid',
                      t('paneFormatSolidFill'),
                      () => onFill(node.sourceId, fillColor ?? gradFrom),
                    )}
                    {radioRow(
                      `fill-${node.sourceId}`,
                      fillKind === 'gradient',
                      t('paneFillGradient'),
                      () =>
                        applyGradient({
                          stops: [
                            { pos: 0, color: fillColor ?? gradFrom },
                            { pos: 1, color: gradTo },
                          ],
                        }),
                    )}
                    {onImageFill &&
                      radioRow(
                        `fill-${node.sourceId}`,
                        fillKind === 'image',
                        t('paneFillImage'),
                        () => onImageFill(node.sourceId),
                      )}
                  </div>
                  {fillKind === 'solid' && (
                    <>
                      <div className="fp-prow">
                        <span>{t('paneFormatSolidFill')}</span>
                        <ColorWell
                          value={fillColor ?? '#ffffff'}
                          label={t('paneFormatSolidFill')}
                          onPick={(hex) =>
                            debouncedFill(node.sourceId, fillValue(hex, fillTransparency))
                          }
                        />
                      </div>
                      <label className="fp-prow">
                        <span>{t('ribbonTransparency')}</span>
                        <select
                          key={`${node.sourceId}:fa:${fillTransparency}`}
                          defaultValue={fillTransparency}
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
                    </>
                  )}
                  {fillKind === 'gradient' && (
                    <>
                      <div className="fp-prow">
                        <span>{t('paneGradientStyle')}</span>
                        <div className="fp-dirwrap" onPointerDown={(e) => e.stopPropagation()}>
                          <div className="fp-btnrow">
                            {styleTiles.map(({ kind, css }) => (
                              <button
                                key={kind}
                                type="button"
                                className={`fp-gpreset ${curPath === kind ? 'sel' : ''}`}
                                aria-label={
                                  kind === 'linear'
                                    ? t('paneGradientLinear')
                                    : kind === 'circle'
                                      ? t('ribbonGradientDirRadial')
                                      : kind === 'rect'
                                        ? t('paneGradientRect')
                                        : t('paneGradientPath')
                                }
                                data-tip={
                                  kind === 'linear'
                                    ? t('paneGradientLinear')
                                    : kind === 'circle'
                                      ? t('ribbonGradientDirRadial')
                                      : kind === 'rect'
                                        ? t('paneGradientRect')
                                        : t('paneGradientPath')
                                }
                                style={{ background: css }}
                                onClick={() => {
                                  applyGradient({ path: kind })
                                  setDirOpenFor(kind === 'circle' || kind === 'rect' ? kind : null)
                                }}
                              />
                            ))}
                          </div>
                          {dirOpenFor && (
                            <div className="fp-dir-pop">
                              {FOCUS_POINTS.map(([glyph, x, y]) => (
                                <button
                                  key={glyph}
                                  type="button"
                                  className={`fp-gpreset ${curCenter.x === x && curCenter.y === y ? 'sel' : ''}`}
                                  aria-label={glyph}
                                  style={{
                                    background: `radial-gradient(circle at ${x * 100}% ${y * 100}%, ${gradPrevFrom}, ${gradPrevTo})`,
                                  }}
                                  onPointerDown={() => {
                                    applyGradient({ path: dirOpenFor, center: { x, y } })
                                    setDirOpenFor(null)
                                  }}
                                />
                              ))}
                            </div>
                          )}
                        </div>
                      </div>
                      {curPath === 'linear' && (
                        <div className="fp-prow">
                          <span>{t('paneGradientAngle')}</span>
                          <div className={`fp-angle ${curPath !== 'linear' ? 'off' : ''}`}>
                            <button
                              type="button"
                              className="fp-dial"
                              disabled={curPath !== 'linear'}
                              aria-label={t('paneGradientAngle')}
                              onPointerDown={onDialDown}
                            >
                              <span
                                className="fp-dial-dot"
                                style={{ transform: `rotate(${curAngle}deg) translateX(6px)` }}
                              />
                            </button>
                            <div className="fp-stepper">
                              <button
                                type="button"
                                disabled={curPath !== 'linear'}
                                aria-label="−10°"
                                onClick={() =>
                                  applyGradient({ angleDeg: normAngle(curAngle - 10) })
                                }
                              >
                                −
                              </button>
                              <input
                                key={`${node.sourceId}:gang:${curAngle}`}
                                type="number"
                                min={0}
                                max={359.9}
                                step={1}
                                defaultValue={curAngle}
                                // width hugs the text ('.' ≈ half a digit), so the value+° group truly centers
                                style={{
                                  width: `${
                                    String(curAngle).replace(/[^0-9]/g, '').length +
                                    (String(curAngle).includes('.') ? 0.5 : 0) +
                                    0.2
                                  }ch`,
                                }}
                                disabled={curPath !== 'linear'}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                                }}
                                onBlur={(e) => {
                                  const v = Number(e.target.value)
                                  if (!Number.isNaN(v)) applyGradient({ angleDeg: normAngle(v) })
                                }}
                              />
                              <span className="fp-stepper-deg">°</span>
                              <button
                                type="button"
                                disabled={curPath !== 'linear'}
                                aria-label="+10°"
                                onClick={() =>
                                  applyGradient({ angleDeg: normAngle(curAngle + 10) })
                                }
                              >
                                ＋
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                      <div className="fp-prow">
                        <span>{t('paneGradientStops')}</span>
                        <div className="fp-btnrow">
                          <button className="fp-btn" onClick={addStop} aria-label="+">
                            ＋
                          </button>
                          <button
                            className="fp-btn"
                            disabled={gradStops.length <= 2}
                            onClick={removeStop}
                            aria-label="−"
                          >
                            −
                          </button>
                        </div>
                      </div>
                      <div className="fp-gstops">
                        <div
                          className="fp-gstops-bar"
                          style={{
                            background: `linear-gradient(90deg, ${gradStops
                              .map((s) => `${s.color} ${Math.round(s.pos * 1000) / 10}%`)
                              .join(', ')})`,
                          }}
                        />
                        {gradStops.map((s, i) => (
                          <button
                            key={`${i}-${s.pos}`}
                            type="button"
                            className={`fp-gstop ${i === selStopIdx ? 'sel' : ''}`}
                            style={{ left: `${s.pos * 100}%`, background: toHex6(s.color) }}
                            aria-label={`${Math.round(s.pos * 100)}%`}
                            onClick={() => setStopIdx(i)}
                          />
                        ))}
                      </div>
                      <div className="fp-prow">
                        <span>{t('paneGradientColor')}</span>
                        <ColorWell
                          value={toHex6(selStop.color)}
                          label={t('paneGradientColor')}
                          onPick={(hex) =>
                            setStop(selStopIdx, { color: stopColor(hex, selStopTransparency) })
                          }
                        />
                      </div>
                      <label className="fp-prow">
                        <span>{t('paneGradientPos')}</span>
                        <input
                          key={`${node.sourceId}:gp:${selStopIdx}:${selStop.pos}`}
                          type="number"
                          min={0}
                          max={100}
                          defaultValue={Math.round(selStop.pos * 100)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                          }}
                          onBlur={(e) => {
                            const v = Number(e.target.value)
                            if (!Number.isNaN(v))
                              setStop(selStopIdx, { pos: Math.max(0, Math.min(100, v)) / 100 })
                          }}
                        />
                      </label>
                      <label className="fp-prow">
                        <span>{t('ribbonTransparency')}</span>
                        <div className="fp-slider">
                          <input
                            key={`${node.sourceId}:ga:${selStopIdx}:${selStopTransparency}`}
                            type="range"
                            min={0}
                            max={100}
                            defaultValue={selStopTransparency}
                            onChange={(e) => debouncedStopAlpha(Number(e.target.value))}
                          />
                          <span className="fp-slider-val">{selStopTransparency}%</span>
                        </div>
                      </label>
                    </>
                  )}
                  {fillKind === 'image' && onImageFill && (
                    <div className="fp-prow fp-prow-end">
                      <button className="fp-btn" onClick={() => onImageFill(node.sourceId)}>
                        {t('paneFormatImageFill')}
                      </button>
                    </div>
                  )}
                </>
              )}
            </>
          )}

          {effTab === 'text' && textSub === 'textbox' && shape?.text && onTextAnchor && (
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

          {effTab === 'shape' && shapeSub === 'fill' && (shape || pic) && (
            <>
              {secHeader(t('paneLineSection'), lineOpen, () => setLineOpen((v) => !v))}
              {lineOpen && (
                <>
                  <div className="fp-radios">
                    {radioRow(`line-${node.sourceId}`, !stroke, t('paneLineNone'), () =>
                      clearStroke(node.sourceId),
                    )}
                    {radioRow(`line-${node.sourceId}`, !!stroke, t('paneLineSolid'), () =>
                      commitStroke(node.sourceId, {}, true),
                    )}
                  </div>
                  {stroke && (
                    <>
                      <div className="fp-prow">
                        <span>{t('paneFormatOutlineColor')}</span>
                        <ColorWell
                          value={strokeColor}
                          label={t('paneFormatOutlineColor')}
                          onPick={(hex) => commitStroke(node.sourceId, { color: hex })}
                        />
                      </div>
                      <label className="fp-prow">
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
                      <label className="fp-prow">
                        <span>{t('paneFormatDashStyle')}</span>
                        <select
                          key={`${node.sourceId}:sd:${strokeDash}`}
                          defaultValue={strokeDash}
                          onChange={(e) =>
                            commitStroke(node.sourceId, { dash: e.target.value }, true)
                          }
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
                    </>
                  )}
                </>
              )}
            </>
          )}

          {effTab === 'shape' &&
            shapeSub === 'fill' &&
            node.type === 'chart' &&
            chartData &&
            onChartPointColor && (
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
                        <ColorWell
                          value={toHex6(
                            chartData.pointColors[si]?.[pi] ??
                              (chartData.kind === 'pie'
                                ? CHART_PALETTE[pi % CHART_PALETTE.length]!
                                : (chartData.seriesColors[si] ??
                                  CHART_PALETTE[si % CHART_PALETTE.length]!)),
                          )}
                          label={chartData.categories[pi] || `#${pi + 1}`}
                          onPick={(hex) => debouncedPointColor(si, pi, hex)}
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

          {effTab === 'shape' && shapeSub === 'size' && onOpenLink && (
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
              <div className="fp-prow fp-prow-end">
                <button className="fp-btn" onClick={onOpenLink}>
                  {t('paneFormatLinkSet')}
                </button>
              </div>
            </>
          )}

          {effTab === 'shape' && shapeSub === 'size' && (
            <>
              <div className="fp-section">{t('paneFormatActions')}</div>
              <div className="fp-prow fp-prow-end">
                <button className="fp-btn fp-danger" onClick={() => onDelete(node.sourceId)}>
                  {t('paneFormatDelete')}
                </button>
              </div>
            </>
          )}

          {/* Sub-tabs whose options aren't editable yet (effects, text fill; fill&line for e.g. groups) */}
          {((effTab === 'shape' && shapeSub === 'effects') ||
            (effTab === 'shape' && shapeSub === 'fill' && !fillHasContent) ||
            (effTab === 'text' && (textSub === 'fill' || textSub === 'effects'))) && (
            <div className="fp-empty">{t('paneSubNone')}</div>
          )}
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

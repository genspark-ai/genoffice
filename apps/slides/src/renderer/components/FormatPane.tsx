/**
 * Format pane (a trimmed-down PowerPoint Format Pane): position/size/rotation/fill of the
 * selected element. Shares the right dock area with the AI panel, mutually exclusive. Inputs
 * commit on blur/Enter; external changes (dragging etc.) sync default values by remounting
 * inputs via key.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { PictureRenderNode, RenderNode, ShapeRenderNode } from '@genoffice/pptx-render'
import type { GradientFillSpec, LinkTargetOp } from '../../shared/ipc'
import { Dropdown } from '@genoffice/ui'
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
    stroke: {
      color: string
      widthPt: number
      dash?: string
      cap?: 'flat' | 'rnd' | 'sq'
      join?: 'round' | 'bevel' | 'miter'
      compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri'
      gradient?: { stops: Array<{ pos: number; color: string }>; angleDeg: number }
    } | null,
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

/** Thin chevron for the width spinner (Lucide-style: currentColor, round caps). */
function SpinChevron({ up }: { up: boolean }) {
  return (
    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d={up ? 'M5.5 14.75 12 8.25l6.5 6.5' : 'M5.5 9.25 12 15.75l6.5-6.5'}
        stroke="currentColor"
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/**
 * PPT-style transparency control: slider + a text field showing "<value> %" —
 * the unit is part of the editable text, so selecting digits, digits+unit or
 * clearing all behaves like any text field. The slider is controlled by local
 * state so model round-trips never remount it mid-drag; commits are throttled
 * (~8/s leading + trailing) so the canvas previews live without flooding IPC.
 */
function PctControl({ value, onChange }: { value: number; onChange: (pct: number) => void }) {
  const [pct, setPct] = useState(value)
  const [text, setText] = useState(`${value} %`)
  const dragging = useRef(false)
  const lastSent = useRef(0)
  const trailing = useRef<number | null>(null)
  useEffect(() => {
    if (!dragging.current) {
      setPct(value)
      setText(`${value} %`)
    }
  }, [value])
  useEffect(
    () => () => {
      if (trailing.current) window.clearTimeout(trailing.current)
    },
    [],
  )
  const send = (v: number, force = false) => {
    if (trailing.current) window.clearTimeout(trailing.current)
    const now = performance.now()
    if (force || now - lastSent.current > 120) {
      lastSent.current = now
      onChange(v)
    } else {
      trailing.current = window.setTimeout(() => onChange(v), 130)
    }
  }
  const set = (v: number) => {
    const clamped = Math.max(0, Math.min(100, Math.round(v)))
    setPct(clamped)
    setText(`${clamped} %`)
    send(clamped, true)
  }
  return (
    <div className="fp-slider">
      <input
        type="range"
        min={0}
        max={100}
        value={pct}
        // elapsed-track fill: the CSS gradient reads the current position from this var
        style={{ '--fp-pct': `${pct}%` } as React.CSSProperties}
        onPointerDown={() => (dragging.current = true)}
        onPointerUp={() => {
          dragging.current = false
          send(pct, true)
        }}
        onChange={(e) => {
          const v = Number(e.target.value)
          setPct(v)
          setText(`${v} %`)
          send(v)
        }}
      />
      <div className="fp-unitstep fp-unitstep-pct" onMouseDown={focusSpinnerField}>
        <input
          type="text"
          inputMode="numeric"
          value={text}
          onChange={(e) => {
            const raw = e.target.value
            // free-form while editing, but entries beyond 100 are rejected as typed
            const num = parseFloat(raw)
            if (!Number.isNaN(num) && num > 100) return
            setText(raw)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
          }}
          onBlur={() => {
            const num = parseFloat(text)
            if (Number.isNaN(num)) {
              setText(`${pct} %`)
              return
            }
            set(num)
          }}
        />
        <span className="fp-unitstep-btns">
          <button type="button" aria-label="+1%" disabled={pct >= 100} onClick={() => set(pct + 1)}>
            <SpinChevron up />
          </button>
          <button type="button" aria-label="−1%" disabled={pct <= 0} onClick={() => set(pct - 1)}>
            <SpinChevron up={false} />
          </button>
        </span>
      </div>
    </div>
  )
}

/** Clicking anywhere in a spinner box except the buttons focuses the field and selects the value. */
function focusSpinnerField(e: React.MouseEvent<HTMLDivElement>) {
  const t = e.target as HTMLElement
  if (t.closest('button') || t.tagName === 'INPUT') return
  e.preventDefault()
  const inp = e.currentTarget.querySelector('input')
  if (inp) {
    inp.focus()
    inp.select()
  }
}

/** PowerPoint's line-width ceiling (pt); larger entries show a validation balloon. */
const MAX_LINE_PT = 1584

/** OOXML compound-line (cmpd) presets. */
type CompoundKind = 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri'

/** Stacked-line preview bands per preset: [y, height] pairs in a 0–16 box
 * (native <option> can't render graphics, so the dropdown draws PPT-style
 * previews itself). */
const COMPOUND_BANDS: Record<CompoundKind, Array<[number, number]>> = {
  sng: [[6.5, 3]],
  dbl: [
    [4, 2.5],
    [9.5, 2.5],
  ],
  thickThin: [
    [3.5, 4],
    [10, 1.5],
  ],
  thinThick: [
    [3.5, 1.5],
    [8.5, 4],
  ],
  tri: [
    [3, 1.5],
    [6.75, 2.5],
    [11.5, 1.5],
  ],
}

function CompoundPreview({ kind }: { readonly kind: CompoundKind }) {
  return (
    <svg
      viewBox="0 0 96 16"
      preserveAspectRatio="none"
      className="fp-line-preview"
      aria-hidden="true"
    >
      {COMPOUND_BANDS[kind].map(([y, h], i) => (
        <rect key={i} x="0" y={y} width="96" height={h} fill="currentColor" />
      ))}
    </svg>
  )
}

/** OOXML prstDash presets with dash-pattern previews (stroke-dasharray units). */
const DASH_PRESETS: Array<[string, string | undefined]> = [
  ['solid', undefined],
  ['sysDot', '1.5 2.5'],
  ['dot', '2 3.5'],
  ['dash', '6 4'],
  ['lgDash', '10 4'],
  ['dashDot', '6 4 2 4'],
  ['lgDashDot', '10 4 2 4'],
  ['lgDashDotDot', '10 4 2 4 2 4'],
]

function DashPreview({ dasharray }: { readonly dasharray?: string }) {
  return (
    <svg width="100%" height="16" className="fp-line-preview" aria-hidden="true">
      <line
        x1="2"
        y1="8"
        x2="98%"
        y2="8"
        stroke="currentColor"
        strokeWidth="2"
        strokeDasharray={dasharray}
      />
    </svg>
  )
}

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
  // Width-limit balloon (shown while a keystroke tries to exceed MAX_LINE_PT)
  const [widthLimitTip, setWidthLimitTip] = useState(false)
  const widthTipTimer = useRef<number | null>(null)
  const flashWidthLimitTip = () => {
    setWidthLimitTip(true)
    if (widthTipTimer.current) window.clearTimeout(widthTipTimer.current)
    widthTipTimer.current = window.setTimeout(() => setWidthLimitTip(false), 2500)
  }

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
  const strokeWidthPt = stroke ? Math.max(0.25, Math.round(stroke.widthPt * 4) / 4) : 1
  const strokeColor = stroke ? toHex6(stroke.color) : '#000000'
  const strokeDash = stroke?.dashPreset ?? 'solid'
  // 0..100 line transparency (alpha byte of the stroke color)
  const strokeTransparency = stroke ? Math.round(((255 - alphaOf(stroke.color)) / 255) * 100) : 0
  // UI works in the op's OOXML attribute values ('flat'/'rnd'/'sq'); render model carries canvas caps
  const capToOp = { butt: 'flat', round: 'rnd', square: 'sq' } as const
  const strokeCap: 'flat' | 'rnd' | 'sq' = stroke?.cap ? capToOp[stroke.cap] : 'flat'
  const strokeJoin: 'round' | 'bevel' | 'miter' = stroke?.join ?? 'round'
  const strokeCompound: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' =
    stroke?.compound ?? 'sng'
  const strokeGradient = stroke?.gradient ?? null
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
  interface StrokeDraft {
    id: string
    color: string
    transparencyPct: number
    widthPt: number
    dash: string
    cap: 'flat' | 'rnd' | 'sq'
    join: 'round' | 'bevel' | 'miter'
    compound: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri'
    /** null = solid line; stops keep any mid-stops an opened file had */
    gradient: { stops: Array<{ pos: number; color: string }>; angleDeg: number } | null
  }
  const strokeDraft = useRef<StrokeDraft | null>(null)
  useEffect(() => {
    if (!strokeTimer.current) strokeDraft.current = null
  }, [stroke, node?.sourceId])
  const commitStroke = (
    sourceId: string,
    patch: Partial<Omit<StrokeDraft, 'id'>>,
    immediate = false,
  ) => {
    if (strokeTimer.current) window.clearTimeout(strokeTimer.current)
    const prev: StrokeDraft =
      strokeDraft.current?.id === sourceId
        ? strokeDraft.current
        : {
            id: sourceId,
            color: strokeColor,
            transparencyPct: strokeTransparency,
            widthPt: strokeWidthPt,
            dash: strokeDash,
            cap: strokeCap,
            join: strokeJoin,
            compound: strokeCompound,
            gradient: strokeGradient
              ? { stops: strokeGradient.stops, angleDeg: strokeGradient.angleDeg }
              : null,
          }
    const draft = { ...prev, ...patch }
    strokeDraft.current = draft
    const fire = () => {
      strokeTimer.current = null
      const alpha = Math.round(((100 - draft.transparencyPct) / 100) * 255)
      const color =
        alpha >= 255
          ? draft.color
          : `${draft.color}${Math.max(0, alpha).toString(16).padStart(2, '0')}`
      onStroke(sourceId, {
        color,
        widthPt: draft.widthPt,
        dash: draft.dash,
        cap: draft.cap,
        join: draft.join,
        compound: draft.compound,
        ...(draft.gradient ? { gradient: draft.gradient } : {}),
      })
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
  /** Patch the gradient-line's first/last stop color, keeping any mid-stops. */
  const strokeGradEdge = (which: 'from' | 'to', hex: string) => {
    const cur = strokeDraft.current?.gradient ??
      strokeGradient ?? {
        stops: [
          { pos: 0, color: strokeColor },
          { pos: 1, color: '#FFFFFF' },
        ],
        angleDeg: 90,
      }
    const stops = cur.stops.map((s, i) =>
      (which === 'from' ? i === 0 : i === cur.stops.length - 1) ? { ...s, color: hex } : s,
    )
    return { ...cur, stops }
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
    // Deliberately a div, not a label: only the radio circle itself is the hit
    // area — text and row whitespace must not toggle the option
    <div className="fp-radio" key={label}>
      <input type="radio" name={group} checked={checked} onChange={pick} aria-label={label} />
      <span>{label}</span>
    </div>
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
                        <PctControl
                          value={fillTransparency}
                          onChange={(pct) =>
                            onFill(node.sourceId, fillValue(fillColor ?? '#ffffff', pct))
                          }
                        />
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
                        <PctControl
                          value={selStopTransparency}
                          onChange={(pct) =>
                            setStop(selStopIdx, { color: stopColor(selStop.color, pct) })
                          }
                        />
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
                    {radioRow(
                      `line-${node.sourceId}`,
                      !!stroke && !strokeGradient,
                      t('paneLineSolid'),
                      () => commitStroke(node.sourceId, { gradient: null }, true),
                    )}
                    {radioRow(
                      `line-${node.sourceId}`,
                      !!stroke && !!strokeGradient,
                      t('paneLineGradient'),
                      () =>
                        commitStroke(
                          node.sourceId,
                          { gradient: strokeGradEdge('from', strokeColor) },
                          true,
                        ),
                    )}
                  </div>
                  {stroke && !strokeGradient && (
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
                        <span>{t('ribbonTransparency')}</span>
                        <PctControl
                          value={strokeTransparency}
                          onChange={(pct) =>
                            commitStroke(node.sourceId, { transparencyPct: pct }, true)
                          }
                        />
                      </label>
                    </>
                  )}
                  {stroke && strokeGradient && (
                    <>
                      <div className="fp-prow">
                        <span>{t('paneGradientColor')} 1</span>
                        <ColorWell
                          value={toHex6(strokeGradient.stops[0]?.color ?? strokeColor)}
                          label={`${t('paneGradientColor')} 1`}
                          onPick={(hex) =>
                            commitStroke(node.sourceId, { gradient: strokeGradEdge('from', hex) })
                          }
                        />
                      </div>
                      <div className="fp-prow">
                        <span>{t('paneGradientColor')} 2</span>
                        <ColorWell
                          value={toHex6(
                            strokeGradient.stops[strokeGradient.stops.length - 1]?.color ??
                              '#ffffff',
                          )}
                          label={`${t('paneGradientColor')} 2`}
                          onPick={(hex) =>
                            commitStroke(node.sourceId, { gradient: strokeGradEdge('to', hex) })
                          }
                        />
                      </div>
                      <label className="fp-prow">
                        <span>{t('paneGradientAngle')}</span>
                        <input
                          key={`${node.sourceId}:sga:${strokeGradient.angleDeg}`}
                          type="number"
                          min={0}
                          max={359.9}
                          defaultValue={Math.round(strokeGradient.angleDeg)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                          }}
                          onBlur={(e) => {
                            const v = Number(e.target.value)
                            if (!Number.isNaN(v))
                              commitStroke(
                                node.sourceId,
                                {
                                  gradient: {
                                    ...strokeGradEdge(
                                      'from',
                                      toHex6(strokeGradient.stops[0]?.color ?? strokeColor),
                                    ),
                                    angleDeg: ((v % 360) + 360) % 360,
                                  },
                                },
                                true,
                              )
                          }}
                        />
                      </label>
                    </>
                  )}
                  {stroke && (
                    <>
                      <label className="fp-prow">
                        <span>{t('paneLineWidth')}</span>
                        <div className="fp-unitstep" onMouseDown={focusSpinnerField}>
                          {widthLimitTip && (
                            <div className="fp-limit-tip" role="alert">
                              {t('paneLineWidthMax', { max: MAX_LINE_PT.toLocaleString() })}
                            </div>
                          )}
                          <input
                            key={`${node.sourceId}:sw:${strokeWidthPt}`}
                            type="text"
                            inputMode="decimal"
                            defaultValue={`${strokeWidthPt} ${t('paneFormatPt')}`}
                            onChange={(e) => {
                              // PPT blocks entries over 1584pt as they are typed: the
                              // offending keystroke is reverted and a balloon flashes
                              const raw = e.target.value
                              const v = parseFloat(raw)
                              if (!Number.isNaN(v) && v > MAX_LINE_PT) {
                                e.target.value =
                                  e.target.dataset['prev'] ??
                                  `${strokeWidthPt} ${t('paneFormatPt')}`
                                flashWidthLimitTip()
                              } else {
                                e.target.dataset['prev'] = raw
                                setWidthLimitTip(false)
                              }
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                            }}
                            onBlur={(e) => {
                              const v = parseFloat(e.target.value)
                              if (!Number.isNaN(v) && v > 0 && v <= MAX_LINE_PT)
                                commitStroke(node.sourceId, { widthPt: v }, true)
                              else e.target.value = `${strokeWidthPt} ${t('paneFormatPt')}`
                            }}
                          />
                          <span className="fp-unitstep-btns">
                            <button
                              type="button"
                              aria-label={`+0.25 ${t('paneFormatPt')}`}
                              disabled={
                                (strokeDraft.current?.widthPt ?? strokeWidthPt) >= MAX_LINE_PT
                              }
                              onClick={() =>
                                commitStroke(
                                  node.sourceId,
                                  {
                                    widthPt: Math.min(
                                      MAX_LINE_PT,
                                      (strokeDraft.current?.widthPt ?? strokeWidthPt) + 0.25,
                                    ),
                                  },
                                  true,
                                )
                              }
                            >
                              <SpinChevron up />
                            </button>
                            <button
                              type="button"
                              aria-label={`−0.25 ${t('paneFormatPt')}`}
                              disabled={(strokeDraft.current?.widthPt ?? strokeWidthPt) <= 0.25}
                              onClick={() =>
                                commitStroke(
                                  node.sourceId,
                                  {
                                    widthPt: Math.max(
                                      0.25,
                                      (strokeDraft.current?.widthPt ?? strokeWidthPt) - 0.25,
                                    ),
                                  },
                                  true,
                                )
                              }
                            >
                              <SpinChevron up={false} />
                            </button>
                          </span>
                        </div>
                      </label>
                      <div className="fp-prow">
                        <span>{t('paneLineCompound')}</span>
                        <Dropdown
                          value={strokeCompound}
                          ariaLabel={t('paneLineCompound')}
                          options={(
                            [
                              ['sng', t('paneLineCompoundSng')],
                              ['dbl', t('paneLineCompoundDbl')],
                              ['thickThin', t('paneLineCompoundThickThin')],
                              ['thinThick', t('paneLineCompoundThinThick')],
                              ['tri', t('paneLineCompoundTri')],
                            ] as const
                          ).map(([k, label]) => ({
                            value: k,
                            label,
                            render: <CompoundPreview kind={k} />,
                          }))}
                          onPick={(kind) => commitStroke(node.sourceId, { compound: kind }, true)}
                        />
                      </div>
                      <div className="fp-prow">
                        <span>{t('paneFormatDashStyle')}</span>
                        <Dropdown
                          value={strokeDash}
                          ariaLabel={t('paneFormatDashStyle')}
                          options={[
                            // an off-preset dash from the file keeps a text entry (parity
                            // with the old select's fallback <option>)
                            ...(DASH_PRESETS.some(([k]) => k === strokeDash)
                              ? []
                              : [{ value: strokeDash, label: strokeDash }]),
                            ...DASH_PRESETS.map(([k, dasharray]) => ({
                              value: k,
                              label: k,
                              render: <DashPreview dasharray={dasharray} />,
                            })),
                          ]}
                          onPick={(dash) => commitStroke(node.sourceId, { dash }, true)}
                        />
                      </div>
                      <div className="fp-prow">
                        <span>{t('paneLineCap')}</span>
                        <Dropdown
                          value={strokeCap}
                          ariaLabel={t('paneLineCap')}
                          options={(
                            [
                              ['flat', t('paneLineCapFlat')],
                              ['rnd', t('paneLineCapRound')],
                              ['sq', t('paneLineCapSquare')],
                            ] as const
                          ).map(([k, label]) => ({ value: k, label }))}
                          onPick={(cap) => commitStroke(node.sourceId, { cap }, true)}
                        />
                      </div>
                      <div className="fp-prow">
                        <span>{t('paneLineJoin')}</span>
                        <Dropdown
                          value={strokeJoin}
                          ariaLabel={t('paneLineJoin')}
                          options={(
                            [
                              ['round', t('paneLineJoinRound')],
                              ['bevel', t('paneLineJoinBevel')],
                              ['miter', t('paneLineJoinMiter')],
                            ] as const
                          ).map(([k, label]) => ({ value: k, label }))}
                          onPick={(join) => commitStroke(node.sourceId, { join }, true)}
                        />
                      </div>
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

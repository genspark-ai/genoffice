/**
 * Format pane (a trimmed-down PowerPoint Format Pane): position/size/rotation/fill of the
 * selected element. Shares the right dock area with the AI panel, mutually exclusive. Inputs
 * commit on blur/Enter; external changes (dragging etc.) sync default values by remounting
 * inputs via key.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { PictureRenderNode, RenderNode, ShapeRenderNode } from '@genoffice/pptx-render'
import type { GradientFillSpec, SetEffectsPatch } from '../../shared/ipc'
import { Dropdown, useDismissablePopover } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import { pathGradientCanvas } from '../konva-adapter'
import { ColorWell } from './ColorWell'
import { IconSidebarCollapse } from './icons'

/** PPT "Offset: Bottom Right" style defaults for a freshly enabled shadow (60% transparency, 4pt blur, 3pt dist). */
const DEFAULT_SHADOW = { color: '#00000066', blurRad: 50800, dist: 38100, dirDeg: 45 }
/** Fresh glow default: 5pt gold (PPT presets use theme accents; a fixed one keeps document colors theme-independent). */
const DEFAULT_GLOW = { color: '#FFC000', radius: 63500 }
/** Shadow offset presets: direction key → outerShdw dir angle (clockwise, 0 = right). */
const SHADOW_DIRS = [
  ['r', 0],
  ['br', 45],
  ['b', 90],
  ['bl', 135],
  ['l', 180],
  ['tl', 225],
  ['t', 270],
  ['tr', 315],
] as const
const SHADOW_DIR_LABELS = {
  r: 'paneEffDirR',
  br: 'paneEffDirBR',
  b: 'paneEffDirB',
  bl: 'paneEffDirBL',
  l: 'paneEffDirL',
  tl: 'paneEffDirTL',
  t: 'paneEffDirT',
  tr: 'paneEffDirTR',
} as const
const GLOW_PT_PRESETS = [5, 8, 11, 18]
const SOFT_EDGE_PT_PRESETS = [1, 2.5, 5, 10, 25]

/** Gallery tile shadow offsets (px) per preset key — a miniature of the effect itself. */
/** One gallery preset: the full shadow the tile applies (PPT's offset/inside/perspective values). */
interface ShadowPresetDef {
  inner?: boolean
  dirDeg: number
  dist: number
  blurRad: number
  sx?: number
  sy?: number
  kxDeg?: number
  algn?: string
  /** Preset color (the alpha byte is the preset's transparency; RGB keeps the user's pick) */
  color: string
}

const OUTER = (dirDeg: number): ShadowPresetDef => ({
  dirDeg,
  dist: 38100,
  blurRad: 50800,
  color: '#00000066',
})
const INNER = (dirDeg: number): ShadowPresetDef => ({
  inner: true,
  dirDeg,
  dist: 50800,
  blurRad: 63500,
  color: '#00000080',
})
/** Gallery presets. Inner dir points AWAY from the named edge (the outside mass casts
 * inward from there — "Inside: Top Left" is dir 45°, matching PowerPoint's XML). */
const SHADOW_PRESETS: Record<string, ShadowPresetDef> = {
  r: OUTER(0),
  br: OUTER(45),
  b: OUTER(90),
  bl: OUTER(135),
  l: OUTER(180),
  tl: OUTER(225),
  t: OUTER(270),
  tr: OUTER(315),
  center: { dirDeg: 0, dist: 0, blurRad: 63500, color: '#00000066' },
  'in-tl': INNER(45),
  'in-t': INNER(90),
  'in-tr': INNER(135),
  'in-l': INNER(0),
  'in-center': { inner: true, dirDeg: 0, dist: 0, blurRad: 114300, color: '#00000080' },
  'in-r': INNER(180),
  'in-bl': INNER(315),
  'in-b': INNER(270),
  'in-br': INNER(225),
  // Perspective: ground shadows built from PowerPoint's real preset numbers ("Upper
  // Left" reads back in PPT as angle 225° / distance 0pt / blur 6pt / transparency 80%):
  // a flipped 23%-squashed silhouette hugging the bottom edge (dist 0), leaned by kx —
  // positive kxDeg leans the slab's far edge LEFT. Row 2 drops farther with more lean.
  'p-ul': {
    dirDeg: 225,
    dist: 0,
    blurRad: 76200,
    sx: 1,
    sy: -0.23,
    kxDeg: 13.34,
    color: '#00000033',
  },
  'p-ur': {
    dirDeg: 315,
    dist: 0,
    blurRad: 76200,
    sx: 1,
    sy: -0.23,
    kxDeg: -13.34,
    color: '#00000033',
  },
  'p-b': {
    dirDeg: 90,
    dist: 76200,
    blurRad: 76200,
    sx: 1,
    sy: -0.23,
    color: '#00000033',
  },
  'p-ll': {
    dirDeg: 90,
    dist: 114300,
    blurRad: 76200,
    sx: 1,
    sy: -0.3,
    kxDeg: 20,
    color: '#00000040',
  },
  'p-lr': {
    dirDeg: 90,
    dist: 114300,
    blurRad: 76200,
    sx: 1,
    sy: -0.3,
    kxDeg: -20,
    color: '#00000040',
  },
}

/** PPT gallery orders (3×3 grids; perspective is a 3+2 row). */
const SHADOW_GAL_OUTER = ['br', 'b', 'bl', 'r', 'center', 'l', 'tr', 't', 'tl'] as const
const SHADOW_GAL_INNER = [
  'in-tl',
  'in-t',
  'in-tr',
  'in-l',
  'in-center',
  'in-r',
  'in-bl',
  'in-b',
  'in-br',
] as const
const SHADOW_GAL_PERSP = ['p-ul', 'p-ur', 'p-b', 'p-ll', 'p-lr'] as const

/** Perspective tiles: ground-shadow slabs matching PPT's gallery (bottom-hugging
 * lean-left / lean-right, a detached flat bar below, then the two farther drops). */
const PERSP_TILE_SHADOW: Record<string, string> = {
  'p-ul': '-5px 4px 3px -1px var(--fp-gal-shadow)',
  'p-ur': '5px 4px 3px -1px var(--fp-gal-shadow)',
  'p-b': '0 8px 3px -4px var(--fp-gal-shadow)',
  'p-ll': '-6px 9px 4px -2px var(--fp-gal-shadow)',
  'p-lr': '6px 9px 4px -2px var(--fp-gal-shadow)',
}

function galTileShadow(key: string): React.CSSProperties | undefined {
  if (key === 'none') return undefined
  const persp = PERSP_TILE_SHADOW[key]
  if (persp) return { boxShadow: persp }
  const p = SHADOW_PRESETS[key]
  if (!p) return undefined
  const rad = (p.dirDeg * Math.PI) / 180
  if (p.inner) {
    if (!p.dist) return { boxShadow: 'inset 0 0 5px var(--fp-gal-shadow)' }
    return {
      boxShadow: `inset ${(Math.cos(rad) * 3).toFixed(1)}px ${(Math.sin(rad) * 3).toFixed(1)}px 4px var(--fp-gal-shadow)`,
    }
  }
  if (!p.dist) return { boxShadow: '0 0 4px 1px var(--fp-gal-shadow)' }
  return {
    boxShadow: `${(Math.cos(rad) * 2.8).toFixed(1)}px ${(Math.sin(rad) * 2.8).toFixed(1)}px 3px var(--fp-gal-shadow)`,
  }
}

/** Nearest direction key by angular distance. */
function nearestDirKey(deg: number, dirs: ReadonlyArray<readonly [string, number]>): string {
  const dist = (x: number) => {
    const n = ((x % 360) + 360) % 360
    return Math.min(n, 360 - n)
  }
  let best = dirs[0]![0]
  let bestD = Infinity
  for (const [k, a] of dirs) {
    const d = dist(deg - a)
    if (d < bestD) {
      bestD = d
      best = k
    }
  }
  return best
}

/** Inner presets keyed by the edge the shadow hugs (dir points away from it). */
const INNER_DIRS = [
  ['in-tl', 45],
  ['in-t', 90],
  ['in-tr', 135],
  ['in-l', 0],
  ['in-r', 180],
  ['in-bl', 315],
  ['in-b', 270],
  ['in-br', 225],
] as const

/** PPT-style shadow preset gallery: thumbnail trigger + grouped grid of live shadow tiles. */
function ShadowPresetPicker({
  current,
  onPick,
}: {
  readonly current: string
  readonly onPick: (key: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLSpanElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  const dirLabel = (dir: string) =>
    dir === 'center'
      ? t('paneEffCenter')
      : t(SHADOW_DIR_LABELS[dir as keyof typeof SHADOW_DIR_LABELS])
  const PERSP_DIR: Record<string, string> = {
    'p-ul': 'tl',
    'p-ur': 'tr',
    'p-b': 'b',
    'p-ll': 'bl',
    'p-lr': 'br',
  }
  const label = (key: string) =>
    key === 'none'
      ? t('paneEffNoShadow')
      : key.startsWith('p-')
        ? `${t('paneEffPerspective')}: ${dirLabel(PERSP_DIR[key]!)}`
        : key.startsWith('in-')
          ? `${t('paneEffInner')}: ${dirLabel(key.slice(3))}`
          : dirLabel(key)
  const tile = (key: string) => (
    <button
      key={key}
      type="button"
      role="option"
      aria-selected={current === key}
      aria-label={label(key)}
      title={label(key)}
      className={`fp-galopt${current === key ? ' selected' : ''}`}
      onClick={() => {
        setOpen(false)
        onPick(key)
      }}
    >
      <span className="fp-galtile" style={galTileShadow(key)} />
    </button>
  )
  return (
    <span ref={wrapRef} className="fp-shadowgal">
      <button
        type="button"
        className="fp-shadowgal-btn"
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={t('paneEffPreset')}
        title={label(current)}
        onClick={() => setOpen((v) => !v)}
      >
        {/* Fixed glyph: the trigger doesn't preview the current pick (matches PPT) */}
        <span className="fp-galtile fp-galtile-cur" />
        <span className="gs-dd-caret" aria-hidden="true">
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none">
            <path
              d="M5.5 9.25 12 15.75l6.5-6.5"
              stroke="currentColor"
              strokeWidth="2.6"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </span>
      </button>
      {open && (
        <div className="fp-shadowgal-pop" role="listbox" aria-label={t('paneEffPreset')}>
          <div className="fp-shadowgal-group">{t('paneEffNoShadow')}</div>
          {tile('none')}
          <div className="fp-shadowgal-group">{t('paneEffOuter')}</div>
          <div className="fp-shadowgal-grid">{SHADOW_GAL_OUTER.map((k) => tile(k))}</div>
          <div className="fp-shadowgal-group">{t('paneEffInner')}</div>
          <div className="fp-shadowgal-grid">{SHADOW_GAL_INNER.map((k) => tile(k))}</div>
          <div className="fp-shadowgal-group">{t('paneEffPerspective')}</div>
          <div className="fp-shadowgal-grid">{SHADOW_GAL_PERSP.map((k) => tile(k))}</div>
        </div>
      )}
    </span>
  )
}

/** Reflection gallery (PPT's 3×3: tight/half/full × touching/4pt/8pt offset). */
const REFL_KINDS = [
  ['tight', 0.35],
  ['half', 0.55],
  ['full', 0.9],
] as const
const REFL_DISTS = [0, 50800, 101600] as const
const DEFAULT_REFLECTION = { blurRad: 6350, startA: 0.52, endPos: 0.55, dist: 0 }
const REFL_KIND_LABELS = {
  tight: 'paneEffReflTight',
  half: 'paneEffReflHalf',
  full: 'paneEffReflFull',
} as const

/** PPT-style reflection preset gallery: tiles show a face with a fading mirrored bar. */
function ReflectionPresetPicker({
  current,
  onPick,
}: {
  readonly current: string
  readonly onPick: (key: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLSpanElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  const label = (key: string) => {
    if (key === 'none') return t('paneEffNoReflection')
    const [kind, di] = key.split('-') as [keyof typeof REFL_KIND_LABELS, string]
    const dist = REFL_DISTS[Number(di)] ?? 0
    return `${t(REFL_KIND_LABELS[kind])}: ${
      dist ? `${dist / 12700} ${t('paneFormatPt')}` : t('paneEffReflTouch')
    }`
  }
  const tile = (key: string) => {
    const [kind, di] = key.split('-')
    const endPos = REFL_KINDS.find(([k]) => k === kind)?.[1] ?? 0
    return (
      <button
        key={key}
        type="button"
        role="option"
        aria-selected={current === key}
        aria-label={label(key)}
        title={label(key)}
        className={`fp-galopt${current === key ? ' selected' : ''}`}
        onClick={() => {
          setOpen(false)
          onPick(key)
        }}
      >
        <span className="fp-refltile">
          <span className="fp-galtile fp-refltile-face" />
          {key !== 'none' && (
            <span
              className="fp-refltile-fade"
              style={{ height: 4 + endPos * 12, marginTop: 1 + Number(di) * 2 }}
            />
          )}
        </span>
      </button>
    )
  }
  return (
    <span ref={wrapRef} className="fp-shadowgal">
      <button
        type="button"
        className="fp-shadowgal-btn"
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={t('paneEffPreset')}
        title={label(current)}
        onClick={() => setOpen((v) => !v)}
      >
        {/* Fixed glyph (matches the shadow picker): a tile with a mirrored fade */}
        <span className="fp-refltile fp-refltile-cur">
          <span className="fp-galtile fp-refltile-face" />
          <span className="fp-refltile-fade" style={{ height: 6, marginTop: 1 }} />
        </span>
        <span className="gs-dd-caret" aria-hidden="true">
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none">
            <path
              d="M5.5 9.25 12 15.75l6.5-6.5"
              stroke="currentColor"
              strokeWidth="2.6"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </span>
      </button>
      {open && (
        <div className="fp-shadowgal-pop" role="listbox" aria-label={t('paneEffPreset')}>
          <div className="fp-shadowgal-group">{t('paneEffNoReflection')}</div>
          {tile('none')}
          <div className="fp-shadowgal-group">{t('paneEffPreset')}</div>
          <div className="fp-shadowgal-grid">
            {REFL_DISTS.map((_, di) => REFL_KINDS.map(([kind]) => tile(`${kind}-${di}`)))}
          </div>
        </div>
      )}
    </span>
  )
}

type TextVert = 'horz' | 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'

/** PPT-style text-direction glyph: a CJK sample word + ABC laid out the way the
 * option lays out text. Icon art (like WordArt previews), so the glyph text is fixed. */
function TextDirIcon({ kind }: { readonly kind: TextVert }): React.JSX.Element {
  const rows = (
    <>
      <text x="10" y="8.5" textAnchor="middle" fontSize="8">
        文字
      </text>
      <text x="10" y="16.5" textAnchor="middle" fontSize="6.5">
        ABC
      </text>
    </>
  )
  let body: React.ReactNode
  if (kind === 'horz') body = rows
  else if (kind === 'vert') body = <g transform="rotate(90 10 10)">{rows}</g>
  else if (kind === 'vert270') body = <g transform="rotate(-90 10 10)">{rows}</g>
  else if (kind === 'eaVert')
    body = (
      <>
        <text x="14.5" y="9" textAnchor="middle" fontSize="8">
          文
        </text>
        <text x="14.5" y="18" textAnchor="middle" fontSize="8">
          字
        </text>
        <text x="5.5" y="12.5" textAnchor="middle" fontSize="6.5" transform="rotate(90 5.5 10)">
          ABC
        </text>
      </>
    )
  else
    body = (
      <>
        <text x="6" y="9" textAnchor="middle" fontSize="8">
          文
        </text>
        <text x="6" y="18" textAnchor="middle" fontSize="8">
          字
        </text>
        <text x="14.5" y="7" textAnchor="middle" fontSize="6">
          A
        </text>
        <text x="14.5" y="13" textAnchor="middle" fontSize="6">
          B
        </text>
        <text x="14.5" y="19" textAnchor="middle" fontSize="6">
          C
        </text>
      </>
    )
  return (
    <svg
      className="fp-dirico"
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="currentColor"
      aria-hidden="true"
    >
      {body}
    </svg>
  )
}

interface Props {
  node: RenderNode | null
  /** Viewport scale of the render tree (fitWidthPx / slide baseline px width) — needed to show sizes in cm */
  viewScale?: number
  /** Slide dimensions in viewport px — needed for the position "From: Center" reference */
  slideSizePx?: { w: number; h: number }
  onTransform: (
    sourceId: string,
    box: { x: number; y: number; w: number; h: number; rotationDeg: number },
  ) => void
  onFill: (sourceId: string, fill: string | GradientFillSpec) => void
  /** Shape picture fill (main process opens the image picker dialog) */
  onImageFill?: (sourceId: string) => void
  /** Text box vertical alignment */
  onTextAnchor?: (sourceId: string, anchor: 'top' | 'middle' | 'bottom') => void
  /** Text box body properties (direction / autofit / internal margins / wrap) */
  onTextBodyProps?: (
    sourceId: string,
    props: {
      vert?: 'horz' | 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'
      autofit?: 'none' | 'shrink' | 'resize'
      insets?: Partial<{ l: number; t: number; r: number; b: number }>
      wrap?: boolean
    },
  ) => void
  /** Shape/picture effects (shadow / glow / soft edge); null clears an effect */
  onEffects?: (sourceId: string, effects: SetEffectsPatch) => void
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
  onCollapse: () => void
  /** Picture: enter crop mode */
  onPictureCrop?: () => void
  /** Picture: enter cutout (background removal) mode */
  onPictureCutout?: () => void
  /** Whether the selected picture supports background removal (audio/video poster frames etc. don't) */
  pictureCanCutout?: boolean
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
/** Slider + numeric box (PPT effect-pane row). Values are display units; text entry
 * keeps one decimal, the slider moves in integer steps, sends are ~120ms throttled. */
function SliderControl({
  value,
  min,
  max,
  unit,
  onChange,
}: {
  value: number
  min: number
  max: number
  /** Display suffix, including any leading space (' %', ' ' + localized pt, '°') */
  unit: string
  onChange: (v: number) => void
}) {
  const fmt = (v: number) => `${Math.round(v * 10) / 10}${unit}`
  const [cur, setCur] = useState(value)
  const [text, setText] = useState(fmt(value))
  const dragging = useRef(false)
  const lastSent = useRef(0)
  const trailing = useRef<number | null>(null)
  useEffect(() => {
    if (!dragging.current) {
      setCur(value)
      setText(fmt(value))
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- fmt is stable per unit
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
    const clamped = Math.max(min, Math.min(max, Math.round(v * 10) / 10))
    setCur(clamped)
    setText(fmt(clamped))
    send(clamped, true)
  }
  return (
    <div className="fp-slider">
      <input
        type="range"
        min={min}
        max={max}
        value={Math.round(cur)}
        // elapsed-track fill: the CSS gradient reads the current position from this var
        style={
          {
            '--fp-pct': `${((cur - min) / Math.max(max - min, 1)) * 100}%`,
          } as React.CSSProperties
        }
        onPointerDown={() => (dragging.current = true)}
        onPointerUp={() => {
          dragging.current = false
          send(cur, true)
        }}
        onChange={(e) => {
          const v = Number(e.target.value)
          setCur(v)
          setText(fmt(v))
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
            // free-form while editing, but entries beyond the range are rejected as typed
            const num = parseFloat(raw)
            if (!Number.isNaN(num) && num > max) return
            setText(raw)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
          }}
          onBlur={() => {
            const num = parseFloat(text)
            if (Number.isNaN(num)) {
              setText(fmt(cur))
              return
            }
            set(num)
          }}
        />
        <span className="fp-unitstep-btns">
          <button
            type="button"
            aria-label={`+1${unit}`}
            disabled={cur >= max}
            onClick={() => set(cur + 1)}
          >
            <SpinChevron up />
          </button>
          <button
            type="button"
            aria-label={`−1${unit}`}
            disabled={cur <= min}
            onClick={() => set(cur - 1)}
          >
            <SpinChevron up={false} />
          </button>
        </span>
      </div>
    </div>
  )
}

function PctControl({ value, onChange }: { value: number; onChange: (pct: number) => void }) {
  return (
    <SliderControl
      value={value}
      min={0}
      max={100}
      unit=" %"
      onChange={(v) => onChange(Math.round(v))}
    />
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
  viewScale,
  slideSizePx,
  onTransform,
  onFill,
  onImageFill,
  onTextAnchor,
  onTextBodyProps,
  onEffects,
  onStroke,
  onCollapse,
  onPictureCrop,
  onPictureCutout,
  pictureCanCutout,
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
  const [sizeOpen, setSizeOpen] = useState(true)
  const [posOpen, setPosOpen] = useState(true)
  // Position "From:" reference per axis (PPT: top-left corner / center of the slide)
  const [posFromH, setPosFromH] = useState<'tl' | 'center'>('tl')
  const [posFromV, setPosFromV] = useState<'tl' | 'center'>('tl')
  const [picOpen, setPicOpen] = useState(true)
  const [txtOpen, setTxtOpen] = useState(true)
  const [effShadowOpen, setEffShadowOpen] = useState(true)
  const [effReflOpen, setEffReflOpen] = useState(true)
  const [effGlowOpen, setEffGlowOpen] = useState(true)
  const [effSoftOpen, setEffSoftOpen] = useState(true)
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
  // PPT Size-section state: aspect-ratio lock + "relative to original picture size".
  // The natural size is decoded from the picture dataUrl (the render node doesn't carry it).
  const [lockRatio, setLockRatio] = useState(false)
  const [relOriginal, setRelOriginal] = useState(true)
  const [naturalSize, setNaturalSize] = useState<{ url: string; w: number; h: number } | null>(null)
  useEffect(() => {
    setLockRatio(node?.type === 'picture')
    setRelOriginal(true)
  }, [node?.sourceId, node?.type])
  useEffect(() => {
    const url = pic?.dataUrl
    if (!url) return
    let alive = true
    const img = new Image()
    img.onload = () => {
      if (alive) setNaturalSize({ url, w: img.naturalWidth, h: img.naturalHeight })
    }
    img.src = url
    return () => {
      alive = false
    }
  }, [pic?.dataUrl])
  // Derived at render time and keyed by dataUrl: a selection change can paint before the
  // decode effect runs, and the previous picture's dimensions must never leak into scaleBasis
  const naturalPx = pic?.dataUrl && naturalSize?.url === pic.dataUrl ? naturalSize : null
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

  // Suppresses the click that follows a stop drag (it would re-select by the stale index)
  const stopDragged = useRef(false)
  /** Drag a gradient stop along the bar: live-preview locally (committing mid-drag would
   * remount the button and break pointer capture), commit once on release. */
  const onStopDown = (e: React.PointerEvent<HTMLButtonElement>, i: number) => {
    setStopIdx(i)
    const el = e.currentTarget
    const wrap = el.parentElement
    if (!wrap) return
    const bar = wrap.querySelector<HTMLElement>('.fp-gstops-bar')
    const rect = wrap.getBoundingClientRect()
    el.setPointerCapture(e.pointerId)
    let pos = gradStops[i]!.pos
    let moved = false
    const move = (ev: PointerEvent) => {
      moved = true
      pos = Math.max(0, Math.min(1, (ev.clientX - rect.left) / rect.width))
      el.style.left = `${pos * 100}%`
      if (bar) {
        const preview = gradStops
          .map((s, j) => (j === i ? { ...s, pos } : s))
          .sort((a, b) => a.pos - b.pos)
        bar.style.background = `linear-gradient(90deg, ${preview
          .map((s) => `${s.color} ${Math.round(s.pos * 1000) / 10}%`)
          .join(', ')})`
      }
    }
    const up = () => {
      el.removeEventListener('pointermove', move)
      el.removeEventListener('pointerup', up)
      if (moved) {
        stopDragged.current = true
        setStop(i, { pos: Math.round(pos * 1000) / 1000 })
      }
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

  // ---- PPT Size/Position section helpers: values are shown in cm / ° / % ----
  const vscale = viewScale ?? 1
  const pxPerCm = (96 / 2.54) * vscale
  const MIN_SIZE_PX = 1
  const fmtNum = (v: number) => String(Math.round(v * 100) / 100)
  const fmtCm = (px: number) => `${fmtNum(px / pxPerCm)} ${t('paneSizeCm')}`

  // ---- Effects section: values shown in pt/°/%, the op speaks EMU ----
  const pxToEmu = (px: number) => Math.round((px / vscale) * 9525)
  const ptToEmu = (pt: number) => Math.round(pt * 12700)
  const fmtEmuPt = (emu: number) => `${fmtNum(emu / 12700)} ${t('paneFormatPt')}`
  const effTarget = shape ?? pic
  const shadowVals = effTarget?.shadow
    ? {
        color: effTarget.shadow.color,
        blurRad: pxToEmu(effTarget.shadow.blurPx),
        // Prefer the source values: the offset vector loses the direction at distance 0
        dist: pxToEmu(
          effTarget.shadow.distPx ?? Math.hypot(effTarget.shadow.offsetX, effTarget.shadow.offsetY),
        ),
        dirDeg:
          effTarget.shadow.dirDeg != null
            ? Math.round(((effTarget.shadow.dirDeg % 360) + 360) % 360)
            : Math.round(
                (Math.atan2(effTarget.shadow.offsetY, effTarget.shadow.offsetX) * 180) / Math.PI +
                  360,
              ) % 360,
        // Kind/perspective fields ride along so blur/distance/angle/color edits keep
        // an inner or perspective shadow what it is instead of degrading it to outer
        ...(effTarget.shadow.inner ? { inner: true as const } : {}),
        ...(effTarget.shadow.scaleX != null ? { sx: effTarget.shadow.scaleX } : {}),
        ...(effTarget.shadow.scaleY != null ? { sy: effTarget.shadow.scaleY } : {}),
        ...(effTarget.shadow.skewXDeg ? { kxDeg: effTarget.shadow.skewXDeg } : {}),
        ...(effTarget.shadow.skewYDeg ? { kyDeg: effTarget.shadow.skewYDeg } : {}),
        ...(effTarget.shadow.algn ? { algn: effTarget.shadow.algn } : {}),
      }
    : null
  const glowVals = effTarget?.glow
    ? { color: effTarget.glow.color, radius: pxToEmu(effTarget.glow.blurPx) }
    : null
  const reflVals = effTarget?.reflection
    ? {
        blurRad: pxToEmu(effTarget.reflection.blurPx),
        startA: effTarget.reflection.startAlpha,
        endPos: effTarget.reflection.endPos,
        dist: pxToEmu(effTarget.reflection.distPx),
      }
    : null
  const softEdgePt = pic?.softEdgePx != null ? pxToEmu(pic.softEdgePx) / 12700 : null
  // Optimistic commit base: while an IPC round trip is in flight, the render node still
  // carries pre-flight values — rebuilding the full effect from it would freeze the other
  // fields and let a queued request overwrite the change that just landed. Each commit
  // builds on the values last SENT for this element; the cache invalidates as soon as the
  // render tree catches up (or an external edit / selection change replaces the node).
  const sentEffects = useRef<{
    key: string
    shadow?: NonNullable<typeof shadowVals>
    glow?: NonNullable<typeof glowVals>
    reflection?: NonNullable<typeof reflVals>
    lastCommitAt: number
  }>({ key: '', lastCommitAt: 0 })
  // Invalidate the optimistic base only after commits have gone QUIET: while edits are
  // streaming, an in-flight IPC result is usually behind the values we already sent, and
  // wiping the base on its arrival would let the next commit rebuild from stale render
  // values. A node update with no commit in the last 600ms can only be our final
  // response or an external edit (undo / AI) — both mean the node is the truth again.
  useEffect(() => {
    const clearBase = () => {
      sentEffects.current.shadow = undefined
      sentEffects.current.glow = undefined
      sentEffects.current.reflection = undefined
    }
    const remaining = 600 - (performance.now() - sentEffects.current.lastCommitAt)
    if (remaining <= 0) {
      clearBase()
      return
    }
    // A node update inside the quiet window may itself be the final state (undo landing
    // right after a drag): re-check once the window expires so the base cannot stick
    // until some unrelated future update. Fresh commits push lastCommitAt forward, and
    // their responses re-run this effect, so the re-check always converges.
    const timer = window.setTimeout(() => {
      if (performance.now() - sentEffects.current.lastCommitAt >= 600) clearBase()
    }, remaining + 20)
    return () => window.clearTimeout(timer)
  }, [effTarget?.shadow, effTarget?.glow, effTarget?.reflection])
  const sentFor = (id: string) => {
    if (sentEffects.current.key !== id) sentEffects.current = { key: id, lastCommitAt: 0 }
    return sentEffects.current
  }
  // Effective current values: everything that BUILDS a patch (relative scaling, color /
  // transparency recombination) must read these, not the render-tree values that lag
  // behind while a round trip is in flight
  const curShadow = (): NonNullable<typeof shadowVals> =>
    (node ? sentFor(node.sourceId).shadow : undefined) ?? shadowVals ?? { ...DEFAULT_SHADOW }
  const curGlow = (): NonNullable<typeof glowVals> =>
    (node ? sentFor(node.sourceId).glow : undefined) ?? glowVals ?? { ...DEFAULT_GLOW }
  const commitShadow = (patch: Partial<NonNullable<typeof shadowVals>>) => {
    if (!node) return
    const sent = sentFor(node.sourceId)
    const next = { ...curShadow(), ...patch }
    sent.shadow = next
    sent.lastCommitAt = performance.now()
    onEffects?.(node.sourceId, { shadow: next })
  }
  const commitGlow = (patch: Partial<NonNullable<typeof glowVals>>) => {
    if (!node) return
    const sent = sentFor(node.sourceId)
    const next = { ...curGlow(), ...patch }
    sent.glow = next
    sent.lastCommitAt = performance.now()
    onEffects?.(node.sourceId, { glow: next })
  }
  const commitReflection = (patch: Partial<NonNullable<typeof reflVals>>) => {
    if (!node) return
    const sent = sentFor(node.sourceId)
    const next = { ...(sent.reflection ?? reflVals ?? DEFAULT_REFLECTION), ...patch }
    sent.reflection = next
    sent.lastCommitAt = performance.now()
    onEffects?.(node.sourceId, { reflection: next })
  }
  // Color pickers / sliders fire repeatedly while dragging; debounce per CHANNEL before
  // IPC — a shared timer would let a glow drag silently cancel a pending shadow edit
  const effectTimers = useRef(new Map<string, number>())
  const debouncedEffects = (channel: string, fire: () => void) => {
    const prev = effectTimers.current.get(channel)
    if (prev) window.clearTimeout(prev)
    effectTimers.current.set(channel, window.setTimeout(fire, 200))
  }
  /** hex6 + transparency% → #RRGGBB(AA) */
  const withAlpha = (hex6: string, transparencyPct: number) => {
    const a = Math.round(((100 - transparencyPct) / 100) * 255)
    return a >= 255 ? hex6 : `${hex6}${a.toString(16).padStart(2, '0')}`
  }
  const transparencyOf = (color: string) => Math.round(((255 - alphaOf(color)) / 255) * 100)

  /** Size commit honoring the aspect-ratio lock (editing one dim scales the other). */
  const sizeCommit = (patch: { w?: number; h?: number }) => {
    if (!box) return
    let w = Math.max(MIN_SIZE_PX, patch.w ?? box.w)
    let h = Math.max(MIN_SIZE_PX, patch.h ?? box.h)
    if (lockRatio) {
      if (patch.h != null && patch.w == null) w = Math.max(MIN_SIZE_PX, box.w * (h / box.h))
      if (patch.w != null && patch.h == null) h = Math.max(MIN_SIZE_PX, box.h * (w / box.w))
    }
    commit({ w, h })
  }

  // Scale basis the % fields are measured against: for pictures with "relative to
  // original picture size" the natural image size (96dpi baseline → viewport px);
  // otherwise the element's size when it was selected (PPT session semantics — the
  // fields track edits made while the selection lasts, and reset on reselect).
  const scaleBaseRef = useRef<{ id: string; w: number; h: number } | null>(null)
  if (node && box && scaleBaseRef.current?.id !== node.sourceId) {
    scaleBaseRef.current = {
      id: node.sourceId,
      w: Math.max(MIN_SIZE_PX, box.w),
      h: Math.max(MIN_SIZE_PX, box.h),
    }
  }
  const scaleBasis =
    node?.type === 'picture' && relOriginal && naturalPx
      ? { w: naturalPx.w * vscale, h: naturalPx.h * vscale }
      : node && scaleBaseRef.current?.id === node.sourceId
        ? { w: scaleBaseRef.current.w, h: scaleBaseRef.current.h }
        : null
  const commitScale = (raw: string, dim: 'w' | 'h') => {
    const v = parseFloat(raw)
    if (Number.isNaN(v) || v <= 0 || !box) return
    sizeCommit({ [dim]: ((scaleBasis ? scaleBasis[dim] : box[dim]) * v) / 100 })
  }
  const stepScale = (dir: 1 | -1, dim: 'w' | 'h') => {
    if (!box) return
    const basis = scaleBasis ? scaleBasis[dim] : box[dim]
    const curPct = Math.round((box[dim] / basis) * 100)
    sizeCommit({ [dim]: (basis * Math.max(1, curPct + dir)) / 100 })
  }

  /** PPT-style value row: label left, "<value> <unit>" spinner box right. */
  const spinRow = (
    label: string,
    display: string,
    commitText: (raw: string) => void,
    onStep: (dir: 1 | -1) => void,
  ) => (
    <label className="fp-prow" key={label}>
      <span>{label}</span>
      <div className="fp-unitstep" onMouseDown={focusSpinnerField}>
        <input
          key={`${node?.sourceId}:${label}:${display}`}
          type="text"
          inputMode="decimal"
          defaultValue={display}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
          }}
          onBlur={(e) => {
            const raw = e.target.value.trim()
            if (raw !== '' && raw !== display) commitText(raw)
            e.target.value = display
          }}
        />
        <span className="fp-unitstep-btns">
          <button type="button" aria-label={`${label} +`} onClick={() => onStep(1)}>
            <SpinChevron up />
          </button>
          <button type="button" aria-label={`${label} −`} onClick={() => onStep(-1)}>
            <SpinChevron up={false} />
          </button>
        </span>
      </div>
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
              {secHeader(t('paneSizeSection'), sizeOpen, () => setSizeOpen((v) => !v))}
              {sizeOpen && (
                <>
                  {spinRow(
                    t('paneSizeHeight'),
                    fmtCm(box.h),
                    (raw) => {
                      const v = parseFloat(raw)
                      if (!Number.isNaN(v) && v > 0) sizeCommit({ h: v * pxPerCm })
                    },
                    (dir) => sizeCommit({ h: box.h + dir * 0.1 * pxPerCm }),
                  )}
                  {spinRow(
                    t('paneSizeWidth'),
                    fmtCm(box.w),
                    (raw) => {
                      const v = parseFloat(raw)
                      if (!Number.isNaN(v) && v > 0) sizeCommit({ w: v * pxPerCm })
                    },
                    (dir) => sizeCommit({ w: box.w + dir * 0.1 * pxPerCm }),
                  )}
                  {spinRow(
                    t('paneSizeRotation'),
                    `${Math.round(box.rotationDeg)}°`,
                    (raw) => {
                      const v = parseFloat(raw)
                      if (!Number.isNaN(v)) commit({ rotationDeg: ((v % 360) + 360) % 360 })
                    },
                    (dir) => commit({ rotationDeg: (((box.rotationDeg + dir) % 360) + 360) % 360 }),
                  )}
                  {spinRow(
                    t('paneSizeScaleH'),
                    `${Math.round((box.h / (scaleBasis?.h ?? box.h)) * 100)}%`,
                    (raw) => commitScale(raw, 'h'),
                    (dir) => stepScale(dir, 'h'),
                  )}
                  {spinRow(
                    t('paneSizeScaleW'),
                    `${Math.round((box.w / (scaleBasis?.w ?? box.w)) * 100)}%`,
                    (raw) => commitScale(raw, 'w'),
                    (dir) => stepScale(dir, 'w'),
                  )}
                  <label className="fp-checkrow">
                    <input
                      type="checkbox"
                      checked={lockRatio}
                      onChange={(e) => setLockRatio(e.target.checked)}
                    />
                    <span>{t('paneSizeLockRatio')}</span>
                  </label>
                  <label className="fp-checkrow">
                    <input
                      type="checkbox"
                      checked={node.type === 'picture' && relOriginal && !!naturalPx}
                      disabled={node.type !== 'picture' || !naturalPx}
                      onChange={(e) => setRelOriginal(e.target.checked)}
                    />
                    <span>{t('paneSizeRelOriginal')}</span>
                  </label>
                </>
              )}

              {secHeader(t('panePosSection'), posOpen, () => setPosOpen((v) => !v))}
              {posOpen && (
                <>
                  {spinRow(
                    t('panePosH'),
                    fmtCm(
                      posFromH === 'center' && slideSizePx
                        ? box.x + box.w / 2 - slideSizePx.w / 2
                        : box.x,
                    ),
                    (raw) => {
                      const v = parseFloat(raw)
                      if (Number.isNaN(v)) return
                      commit({
                        x:
                          posFromH === 'center' && slideSizePx
                            ? slideSizePx.w / 2 + v * pxPerCm - box.w / 2
                            : v * pxPerCm,
                      })
                    },
                    (dir) => commit({ x: box.x + dir * 0.1 * pxPerCm }),
                  )}
                  <div className="fp-prow">
                    <span>{t('panePosFrom')}</span>
                    <Dropdown
                      className="fp-dd-narrow"
                      value={posFromH}
                      ariaLabel={t('panePosFrom')}
                      options={(
                        [
                          ['tl', t('panePosFromTL')],
                          ['center', t('panePosFromCenter')],
                        ] as const
                      ).map(([k, label]) => ({ value: k, label }))}
                      onPick={setPosFromH}
                    />
                  </div>
                  {spinRow(
                    t('panePosV'),
                    fmtCm(
                      posFromV === 'center' && slideSizePx
                        ? box.y + box.h / 2 - slideSizePx.h / 2
                        : box.y,
                    ),
                    (raw) => {
                      const v = parseFloat(raw)
                      if (Number.isNaN(v)) return
                      commit({
                        y:
                          posFromV === 'center' && slideSizePx
                            ? slideSizePx.h / 2 + v * pxPerCm - box.h / 2
                            : v * pxPerCm,
                      })
                    },
                    (dir) => commit({ y: box.y + dir * 0.1 * pxPerCm }),
                  )}
                  <div className="fp-prow">
                    <span>{t('panePosFrom')}</span>
                    <Dropdown
                      className="fp-dd-narrow"
                      value={posFromV}
                      ariaLabel={t('panePosFrom')}
                      options={(
                        [
                          ['tl', t('panePosFromTL')],
                          ['center', t('panePosFromCenter')],
                        ] as const
                      ).map(([k, label]) => ({ value: k, label }))}
                      onPick={setPosFromV}
                    />
                  </div>
                </>
              )}
            </>
          )}

          {effTab === 'shape' && shapeSub === 'size' && node.type === 'picture' && (
            <>
              {secHeader(t('paneFormatPicture'), picOpen, () => setPicOpen((v) => !v))}
              {picOpen && (
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
              )}
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
                            onPointerDown={(e) => onStopDown(e, i)}
                            onClick={() => {
                              // after a drag setStop already re-selected by position; the
                              // trailing click would re-select the pre-sort index
                              if (stopDragged.current) {
                                stopDragged.current = false
                                return
                              }
                              setStopIdx(i)
                            }}
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
                      {spinRow(
                        t('paneGradientPos'),
                        `${Math.round(selStop.pos * 100)}%`,
                        (raw) => {
                          const v = parseFloat(raw)
                          if (!Number.isNaN(v))
                            setStop(selStopIdx, { pos: Math.max(0, Math.min(100, v)) / 100 })
                        },
                        (dir) => {
                          const cur = Math.round(selStop.pos * 100)
                          setStop(selStopIdx, {
                            pos: Math.max(0, Math.min(100, cur + dir)) / 100,
                          })
                        },
                      )}
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
                      {spinRow(
                        t('paneGradientAngle'),
                        `${Math.round(strokeGradient.angleDeg)}°`,
                        (raw) => {
                          const v = parseFloat(raw)
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
                        },
                        (dir) =>
                          commitStroke(
                            node.sourceId,
                            {
                              gradient: {
                                ...strokeGradEdge(
                                  'from',
                                  toHex6(strokeGradient.stops[0]?.color ?? strokeColor),
                                ),
                                angleDeg:
                                  (((Math.round(strokeGradient.angleDeg) + dir) % 360) + 360) % 360,
                              },
                            },
                            true,
                          ),
                      )}
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

          {effTab === 'shape' && shapeSub === 'size' && shape?.text && onTextBodyProps && (
            <>
              {secHeader(t('paneFormatTextBox'), txtOpen, () => setTxtOpen((v) => !v))}
              {txtOpen && (
                <>
                  <div className="fp-prow">
                    <span>{t('paneTextboxVAlign')}</span>
                    <Dropdown
                      value={shape.text.anchor}
                      disabled={!onTextAnchor}
                      ariaLabel={t('paneTextboxVAlign')}
                      options={(
                        [
                          ['top', t('paneVAlignTop')],
                          ['middle', t('paneVAlignMiddle')],
                          ['bottom', t('paneVAlignBottom')],
                        ] as const
                      ).map(([k, label]) => ({ value: k, label }))}
                      onPick={(anchor) => onTextAnchor?.(node.sourceId, anchor)}
                    />
                  </div>
                  <div className="fp-prow">
                    <span>{t('paneTextboxDirection')}</span>
                    <Dropdown
                      className="fp-dd-textdir"
                      value={shape.text.vert ?? 'horz'}
                      ariaLabel={t('paneTextboxDirection')}
                      options={(
                        [
                          ['horz', t('paneTextDirH')],
                          ['eaVert', t('paneTextDirV')],
                          ['vert', t('paneTextDirRot90')],
                          ['vert270', t('paneTextDirRot270')],
                          ['wordArtVert', t('paneTextDirStacked')],
                        ] as const
                      ).map(([k, label]) => ({
                        value: k,
                        label,
                        render: (
                          <>
                            <TextDirIcon kind={k} />
                            <span className="fp-dirico-label">{label}</span>
                          </>
                        ),
                      }))}
                      onPick={(vert) => onTextBodyProps(node.sourceId, { vert })}
                    />
                  </div>
                  <div className="fp-radios">
                    {radioRow(
                      `autofit-${node.sourceId}`,
                      (shape.text.autofit ?? 'none') === 'none',
                      t('paneAutofitNone'),
                      () => onTextBodyProps(node.sourceId, { autofit: 'none' }),
                    )}
                    {radioRow(
                      `autofit-${node.sourceId}`,
                      shape.text.autofit === 'shrink',
                      t('paneAutofitShrink'),
                      () => onTextBodyProps(node.sourceId, { autofit: 'shrink' }),
                    )}
                    {radioRow(
                      `autofit-${node.sourceId}`,
                      shape.text.autofit === 'resize',
                      t('paneAutofitResize'),
                      () => onTextBodyProps(node.sourceId, { autofit: 'resize' }),
                    )}
                  </div>
                  {(
                    [
                      ['l', 'paneInsetL'],
                      ['r', 'paneInsetR'],
                      ['t', 'paneInsetT'],
                      ['b', 'paneInsetB'],
                    ] as const
                  ).map(([side, key]) =>
                    spinRow(
                      t(key),
                      fmtCm(shape.text!.insets[side]),
                      (raw) => {
                        const v = parseFloat(raw)
                        if (!Number.isNaN(v) && v >= 0)
                          onTextBodyProps(node.sourceId, {
                            insets: { [side]: Math.round(v * 360000) },
                          })
                      },
                      (dir) => {
                        const curCm = shape.text!.insets[side] / pxPerCm
                        const next = Math.max(0, Math.round((curCm + dir * 0.1) * 100) / 100)
                        onTextBodyProps(node.sourceId, {
                          insets: { [side]: Math.round(next * 360000) },
                        })
                      },
                    ),
                  )}
                  <label className="fp-checkrow">
                    <input
                      type="checkbox"
                      checked={shape.text.wrap}
                      onChange={(e) => onTextBodyProps(node.sourceId, { wrap: e.target.checked })}
                    />
                    <span>{t('paneTextWrap')}</span>
                  </label>
                </>
              )}
            </>
          )}

          {/* PPT-style Effects tab: shadow / glow / soft edge (reflection & 3-D are not rendered yet) */}
          {effTab === 'shape' && shapeSub === 'effects' && effTarget && onEffects && node && (
            <>
              {secHeader(t('paneEffShadow'), effShadowOpen, () => setEffShadowOpen((v) => !v))}
              {effShadowOpen && (
                <>
                  <div className="fp-prow">
                    <span>{t('paneEffPreset')}</span>
                    <ShadowPresetPicker
                      current={
                        !shadowVals
                          ? 'none'
                          : shadowVals.inner
                            ? shadowVals.dist < 2000
                              ? 'in-center'
                              : nearestDirKey(shadowVals.dirDeg, INNER_DIRS)
                            : (shadowVals.sy ?? 1) < 0 || shadowVals.kxDeg
                              ? Math.abs(shadowVals.kxDeg ?? 0) < 6
                                ? 'p-b'
                                : shadowVals.dist >= 50000
                                  ? (shadowVals.kxDeg ?? 0) > 0
                                    ? 'p-ll'
                                    : 'p-lr'
                                  : (shadowVals.kxDeg ?? 0) > 0
                                    ? 'p-ul'
                                    : 'p-ur'
                              : shadowVals.dist < 2000
                                ? 'center'
                                : nearestDirKey(shadowVals.dirDeg, SHADOW_DIRS)
                      }
                      onPick={(v) => {
                        if (v === 'none') {
                          sentFor(node.sourceId).shadow = undefined
                          return onEffects(node.sourceId, { shadow: null })
                        }
                        // Preset semantics: the tile applies the full PPT preset (blur /
                        // distance / direction / perspective transform / transparency),
                        // keeping only the user's color hue. Applied whole (not through
                        // commitShadow) so stale kind/perspective fields never survive,
                        // but still recorded as the optimistic base for follow-up edits.
                        const p = SHADOW_PRESETS[v]!
                        const a = alphaOf(p.color)
                        const rgb = toHex6(
                          sentFor(node.sourceId).shadow?.color ?? shadowVals?.color ?? p.color,
                        )
                        const next = {
                          color: a >= 255 ? rgb : `${rgb}${a.toString(16).padStart(2, '0')}`,
                          blurRad: p.blurRad,
                          dist: p.dist,
                          dirDeg: p.dirDeg,
                          ...(p.inner ? { inner: true as const } : {}),
                          ...(p.sx != null ? { sx: p.sx } : {}),
                          ...(p.sy != null ? { sy: p.sy } : {}),
                          ...(p.kxDeg ? { kxDeg: p.kxDeg } : {}),
                          ...(p.algn ? { algn: p.algn } : {}),
                        }
                        sentFor(node.sourceId).shadow = next
                        onEffects(node.sourceId, { shadow: next })
                      }}
                    />
                  </div>
                  <div className="fp-prow">
                    <span>{t('paneGradientColor')}</span>
                    <ColorWell
                      value={toHex6(shadowVals?.color ?? DEFAULT_SHADOW.color)}
                      label={t('paneEffShadow')}
                      onPick={(hex) =>
                        debouncedEffects('shadow', () =>
                          commitShadow({
                            color: withAlpha(hex, transparencyOf(curShadow().color)),
                          }),
                        )
                      }
                    />
                  </div>
                  <label className="fp-prow">
                    <span>{t('ribbonTransparency')}</span>
                    <PctControl
                      value={transparencyOf(shadowVals?.color ?? DEFAULT_SHADOW.color)}
                      onChange={(pct) =>
                        debouncedEffects('shadow', () =>
                          commitShadow({ color: withAlpha(toHex6(curShadow().color), pct) }),
                        )
                      }
                    />
                  </label>
                  {/* Size = outerShdw sx/sy silhouette scale; <a:innerShdw> has no scale attrs */}
                  {!shadowVals?.inner && (
                    <label className="fp-prow">
                      <span>{t('paneEffSize')}</span>
                      <SliderControl
                        value={Math.round(Math.abs(shadowVals?.sx ?? 1) * 100)}
                        min={1}
                        max={200}
                        unit="%"
                        onChange={(v) => {
                          // Uniform silhouette scale. Perspective shadows keep their
                          // squash/flip: both axes scale by the same factor.
                          const base = curShadow()
                          const k = v / 100 / Math.max(Math.abs(base.sx ?? 1), 0.01)
                          commitShadow({ sx: (base.sx ?? 1) * k, sy: (base.sy ?? 1) * k })
                        }}
                      />
                    </label>
                  )}
                  <label className="fp-prow">
                    <span>{t('paneEffBlur')}</span>
                    <SliderControl
                      value={(shadowVals?.blurRad ?? 0) / 12700}
                      min={0}
                      max={100}
                      unit={` ${t('paneFormatPt')}`}
                      onChange={(v) => commitShadow({ blurRad: ptToEmu(v) })}
                    />
                  </label>
                  <label className="fp-prow">
                    <span>{t('paneEffDistance')}</span>
                    <SliderControl
                      value={(shadowVals?.dist ?? 0) / 12700}
                      min={0}
                      max={200}
                      unit={` ${t('paneFormatPt')}`}
                      onChange={(v) => commitShadow({ dist: ptToEmu(v) })}
                    />
                  </label>
                  <label className="fp-prow">
                    <span>{t('paneEffAngle')}</span>
                    <SliderControl
                      value={shadowVals?.dirDeg ?? DEFAULT_SHADOW.dirDeg}
                      min={0}
                      max={359}
                      unit="°"
                      onChange={(v) => commitShadow({ dirDeg: ((v % 360) + 360) % 360 })}
                    />
                  </label>
                </>
              )}

              {secHeader(t('paneEffReflection'), effReflOpen, () => setEffReflOpen((v) => !v))}
              {effReflOpen && (
                <>
                  <div className="fp-prow">
                    <span>{t('paneEffPreset')}</span>
                    <ReflectionPresetPicker
                      current={
                        !reflVals
                          ? 'none'
                          : `${REFL_KINDS.reduce(
                              (best, [k, e]) =>
                                Math.abs(reflVals.endPos - e) <
                                Math.abs(
                                  reflVals.endPos -
                                    (REFL_KINDS.find(([bk]) => bk === best)?.[1] ?? 0),
                                )
                                  ? k
                                  : best,
                              'half' as string,
                            )}-${REFL_DISTS.reduce<number>(
                              (bi, d, i) =>
                                Math.abs(reflVals.dist - d) <
                                Math.abs(reflVals.dist - REFL_DISTS[bi]!)
                                  ? i
                                  : bi,
                              0,
                            )}`
                      }
                      onPick={(v) => {
                        if (v === 'none') {
                          sentFor(node.sourceId).reflection = undefined
                          return onEffects(node.sourceId, { reflection: null })
                        }
                        const [kind, di] = v.split('-')
                        commitReflection({
                          endPos: REFL_KINDS.find(([k]) => k === kind)?.[1] ?? 0.55,
                          dist: REFL_DISTS[Number(di)] ?? 0,
                        })
                      }}
                    />
                  </div>
                  <label className="fp-prow">
                    <span>{t('paneEffSize')}</span>
                    <SliderControl
                      value={Math.round((reflVals?.endPos ?? DEFAULT_REFLECTION.endPos) * 100)}
                      min={1}
                      max={100}
                      unit="%"
                      onChange={(v) => commitReflection({ endPos: v / 100 })}
                    />
                  </label>
                  <label className="fp-prow">
                    <span>{t('ribbonTransparency')}</span>
                    <SliderControl
                      value={Math.round(
                        (1 - (reflVals?.startA ?? DEFAULT_REFLECTION.startA)) * 100,
                      )}
                      min={0}
                      max={100}
                      unit=" %"
                      onChange={(v) => commitReflection({ startA: 1 - v / 100 })}
                    />
                  </label>
                  <label className="fp-prow">
                    <span>{t('paneEffBlur')}</span>
                    <SliderControl
                      value={(reflVals?.blurRad ?? DEFAULT_REFLECTION.blurRad) / 12700}
                      min={0}
                      max={100}
                      unit={` ${t('paneFormatPt')}`}
                      onChange={(v) => commitReflection({ blurRad: ptToEmu(v) })}
                    />
                  </label>
                  <label className="fp-prow">
                    <span>{t('paneEffDistance')}</span>
                    <SliderControl
                      value={(reflVals?.dist ?? 0) / 12700}
                      min={0}
                      max={100}
                      unit={` ${t('paneFormatPt')}`}
                      onChange={(v) => commitReflection({ dist: ptToEmu(v) })}
                    />
                  </label>
                </>
              )}

              {secHeader(t('paneEffGlow'), effGlowOpen, () => setEffGlowOpen((v) => !v))}
              {effGlowOpen && (
                <>
                  <div className="fp-prow">
                    <span>{t('paneEffPreset')}</span>
                    <Dropdown
                      value={
                        !glowVals
                          ? 'none'
                          : String(
                              GLOW_PT_PRESETS.reduce((best, pt) =>
                                Math.abs(glowVals.radius / 12700 - pt) <
                                Math.abs(glowVals.radius / 12700 - best)
                                  ? pt
                                  : best,
                              ),
                            )
                      }
                      ariaLabel={t('paneEffPreset')}
                      options={[
                        { value: 'none', label: t('paneEffNone') },
                        ...GLOW_PT_PRESETS.map((pt) => ({
                          value: String(pt),
                          label: `${pt} ${t('paneFormatPt')}`,
                        })),
                      ]}
                      onPick={(v) => {
                        if (v === 'none') {
                          sentFor(node.sourceId).glow = undefined
                          onEffects(node.sourceId, { glow: null })
                        } else commitGlow({ radius: ptToEmu(parseFloat(v)) })
                      }}
                    />
                  </div>
                  <div className="fp-prow">
                    <span>{t('paneGradientColor')}</span>
                    <ColorWell
                      value={toHex6(glowVals?.color ?? DEFAULT_GLOW.color)}
                      label={t('paneEffGlow')}
                      onPick={(hex) =>
                        debouncedEffects('glow', () =>
                          commitGlow({
                            color: withAlpha(hex, transparencyOf(curGlow().color)),
                          }),
                        )
                      }
                    />
                  </div>
                  <label className="fp-prow">
                    <span>{t('ribbonTransparency')}</span>
                    <PctControl
                      value={transparencyOf(glowVals?.color ?? DEFAULT_GLOW.color)}
                      onChange={(pct) =>
                        debouncedEffects('glow', () =>
                          commitGlow({ color: withAlpha(toHex6(curGlow().color), pct) }),
                        )
                      }
                    />
                  </label>
                  {spinRow(
                    t('paneEffSize'),
                    fmtEmuPt(glowVals?.radius ?? 0),
                    (raw) => {
                      const v = parseFloat(raw)
                      if (!Number.isNaN(v)) commitGlow({ radius: ptToEmu(Math.max(0, v)) })
                    },
                    (dir) =>
                      commitGlow({ radius: Math.max(0, (glowVals?.radius ?? 0) + dir * 12700) }),
                  )}
                </>
              )}

              {pic && (
                <>
                  {secHeader(t('paneEffSoftEdge'), effSoftOpen, () => setEffSoftOpen((v) => !v))}
                  {effSoftOpen && (
                    <div className="fp-prow">
                      <span>{t('paneEffPreset')}</span>
                      <Dropdown
                        value={
                          softEdgePt == null
                            ? 'none'
                            : String(
                                SOFT_EDGE_PT_PRESETS.reduce((best, pt) =>
                                  Math.abs(softEdgePt - pt) < Math.abs(softEdgePt - best)
                                    ? pt
                                    : best,
                                ),
                              )
                        }
                        ariaLabel={t('paneEffPreset')}
                        options={[
                          { value: 'none', label: t('paneEffNone') },
                          ...SOFT_EDGE_PT_PRESETS.map((pt) => ({
                            value: String(pt),
                            label: `${pt} ${t('paneFormatPt')}`,
                          })),
                        ]}
                        onPick={(v) =>
                          onEffects(node.sourceId, {
                            softEdge: v === 'none' ? null : ptToEmu(parseFloat(v)),
                          })
                        }
                      />
                    </div>
                  )}
                </>
              )}
            </>
          )}

          {/* Sub-tabs whose options aren't editable yet (text fill; fill&line for e.g. groups) */}
          {((effTab === 'shape' && shapeSub === 'effects' && !(effTarget && onEffects)) ||
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

/**
 * Change chart type dialog: grid of SVG thumbnail cards (modeled on PowerPoint Change Chart Type).
 * Single click selects, double click applies directly; OK calls back onConfirm.
 */
import { useState, type ReactNode } from 'react'
import type { ChartPresetDef } from '../insert-presets'
import { useI18n, type StringKey } from '../i18n/locale'

type ChartKind = ChartPresetDef['kind']

const KINDS: Array<{ kind: ChartKind; label: StringKey }> = [
  { kind: 'bar', label: 'ribbonChartKindBar' },
  { kind: 'barStacked', label: 'ribbonChartKindBarStacked' },
  { kind: 'barPercentStacked', label: 'ribbonChartKindBarPercentStacked' },
  { kind: 'barH', label: 'ribbonChartKindBarH' },
  { kind: 'bar3D', label: 'ribbonChartKindBar3D' },
  { kind: 'line', label: 'ribbonChartKindLine' },
  { kind: 'area', label: 'ribbonChartKindArea' },
  { kind: 'pie', label: 'ribbonChartKindPie' },
  { kind: 'pie3D', label: 'ribbonChartKindPie3D' },
  { kind: 'doughnut', label: 'ribbonChartKindDoughnut' },
  { kind: 'scatter', label: 'ribbonChartKindScatter' },
  { kind: 'radar', label: 'ribbonChartKindRadar' },
  { kind: 'comboBarLine', label: 'ribbonChartKindCombo' },
]

const C1 = '#4472C4'
const C2 = '#ED7D31'
const AXIS = '#B9BEC6'
const GRID = '#E4E7EB'

/** SVG chart-kind thumbnail, shared by this dialog and the Insert-tab chart gallery. */
export function ChartKindThumb({ kind, width = 96 }: { kind: ChartKind; width?: number }) {
  const W = 96
  const H = 60
  const plot = { x: 9, y: 7, w: W - 18, h: H - 15 }
  const bottom = plot.y + plot.h
  const els: ReactNode[] = []
  const cartesian = kind !== 'pie' && kind !== 'pie3D' && kind !== 'doughnut' && kind !== 'radar'

  if (cartesian) {
    for (let i = 1; i <= 2; i++) {
      const y = plot.y + (plot.h * i) / 3
      els.push(
        <line
          key={`g${i}`}
          x1={plot.x}
          y1={y}
          x2={plot.x + plot.w}
          y2={y}
          stroke={GRID}
          strokeWidth={1}
        />,
      )
    }
    els.push(
      <path
        key="axis"
        d={`M ${plot.x} ${plot.y} V ${bottom} H ${plot.x + plot.w}`}
        fill="none"
        stroke={AXIS}
        strokeWidth={1.2}
      />,
    )
  }

  const px = (t: number) => plot.x + plot.w * t
  const py = (v: number) => bottom - plot.h * v

  if (kind === 'bar' || kind === 'comboBarLine') {
    const vals = kind === 'bar' ? [0.55, 0.85, 0.45, 0.7] : [0.5, 0.75, 0.4, 0.62]
    const slot = plot.w / vals.length
    const barW = 10
    vals.forEach((v, i) => {
      const x = plot.x + i * slot + (slot - (kind === 'bar' ? barW * 2 : barW)) / 2
      els.push(<rect key={`b${i}`} x={x} y={py(v)} width={barW} height={plot.h * v} fill={C1} />)
      if (kind === 'bar')
        els.push(
          <rect
            key={`b${i}2`}
            x={x + barW}
            y={py(v * 0.6)}
            width={barW}
            height={plot.h * v * 0.6}
            fill={C2}
          />,
        )
    })
    if (kind === 'comboBarLine') {
      const pts = vals.map((v, i) => `${plot.x + (i + 0.5) * slot},${py(Math.min(v + 0.22, 0.95))}`)
      els.push(<polyline key="cl" points={pts.join(' ')} fill="none" stroke={C2} strokeWidth={2} />)
      vals.forEach((v, i) =>
        els.push(
          <circle
            key={`cd${i}`}
            cx={plot.x + (i + 0.5) * slot}
            cy={py(Math.min(v + 0.22, 0.95))}
            r={2}
            fill={C2}
          />,
        ),
      )
    }
  } else if (kind === 'bar3D') {
    const vals = [0.55, 0.85, 0.45, 0.7]
    const slot = plot.w / vals.length
    const barW = 12
    const d = 4
    vals.forEach((v, i) => {
      const x = plot.x + i * slot + (slot - barW) / 2
      const y = py(v)
      const h = plot.h * v
      els.push(
        <path key={`t${i}`} d={`M ${x} ${y} h ${barW} l ${d} ${-d} h ${-barW} Z`} fill="#6B8FD4" />,
      )
      els.push(
        <path
          key={`f${i}`}
          d={`M ${x + barW} ${y} l ${d} ${-d} v ${h} l ${-d} ${d} Z`}
          fill="#2F5496"
        />,
      )
      els.push(<rect key={`b${i}`} x={x} y={y} width={barW} height={h} fill={C1} />)
    })
  } else if (kind === 'barStacked' || kind === 'barPercentStacked') {
    const vals: Array<[number, number]> =
      kind === 'barStacked'
        ? [
            [0.4, 0.3],
            [0.55, 0.35],
            [0.3, 0.4],
            [0.45, 0.25],
          ]
        : [
            [0.4, 0.6],
            [0.6, 0.4],
            [0.3, 0.7],
            [0.5, 0.5],
          ]
    const slot = plot.w / vals.length
    const barW = 13
    vals.forEach(([a, b], i) => {
      const x = plot.x + i * slot + (slot - barW) / 2
      els.push(<rect key={`s${i}a`} x={x} y={py(a)} width={barW} height={plot.h * a} fill={C1} />)
      els.push(
        <rect key={`s${i}b`} x={x} y={py(a + b)} width={barW} height={plot.h * b} fill={C2} />,
      )
    })
  } else if (kind === 'barH') {
    const vals = [0.85, 0.55, 0.7, 0.4]
    const slot = plot.h / vals.length
    const barW = 8
    vals.forEach((v, i) => {
      const y = plot.y + i * slot + (slot - barW) / 2
      els.push(
        <rect
          key={`h${i}`}
          x={plot.x}
          y={y}
          width={plot.w * v}
          height={barW}
          fill={i % 2 ? C2 : C1}
        />,
      )
    })
  } else if (kind === 'line') {
    const mk = (vals: number[], color: string, key: string) => {
      const pts = vals.map((v, i) => [px((i + 0.5) / vals.length), py(v)] as const)
      els.push(
        <polyline
          key={key}
          points={pts.map((p) => p.join(',')).join(' ')}
          fill="none"
          stroke={color}
          strokeWidth={2}
        />,
      )
      pts.forEach((p, i) =>
        els.push(<circle key={`${key}d${i}`} cx={p[0]} cy={p[1]} r={2} fill={color} />),
      )
    }
    mk([0.25, 0.6, 0.45, 0.85], C1, 'l1')
    mk([0.12, 0.32, 0.6, 0.5], C2, 'l2')
  } else if (kind === 'area') {
    const mk = (vals: number[], color: string, key: string) => {
      const pts = vals.map((v, i) => `${px(i / (vals.length - 1))},${py(v)}`)
      els.push(
        <path
          key={key}
          d={`M ${plot.x} ${bottom} L ${pts.join(' L ')} L ${plot.x + plot.w} ${bottom} Z`}
          fill={color}
          opacity={0.75}
        />,
      )
    }
    mk([0.4, 0.7, 0.5, 0.85, 0.6], C1, 'a1')
    mk([0.2, 0.35, 0.28, 0.5, 0.32], C2, 'a2')
  } else if (kind === 'scatter') {
    const d1: Array<[number, number]> = [
      [0.12, 0.25],
      [0.28, 0.45],
      [0.45, 0.4],
      [0.62, 0.65],
      [0.82, 0.85],
    ]
    const d2: Array<[number, number]> = [
      [0.2, 0.1],
      [0.4, 0.22],
      [0.58, 0.35],
      [0.78, 0.5],
    ]
    d1.forEach(([x, y], i) =>
      els.push(<circle key={`s1${i}`} cx={px(x)} cy={py(y)} r={2.6} fill={C1} />),
    )
    d2.forEach(([x, y], i) =>
      els.push(<circle key={`s2${i}`} cx={px(x)} cy={py(y)} r={2.6} fill={C2} />),
    )
  } else if (kind === 'radar') {
    const cx = W / 2
    const cy = H / 2 + 1
    const R = H / 2 - 6
    const spoke = (i: number, r: number) => {
      const a = (-90 + i * 60) * (Math.PI / 180)
      return [cx + Math.cos(a) * r, cy + Math.sin(a) * r] as const
    }
    for (const r of [R, R * 0.55]) {
      const pts = Array.from({ length: 6 }, (_, i) => spoke(i, r).join(','))
      els.push(
        <polygon key={`w${r}`} points={pts.join(' ')} fill="none" stroke={GRID} strokeWidth={1} />,
      )
    }
    for (let i = 0; i < 6; i++) {
      const [x, y] = spoke(i, R)
      els.push(<line key={`sp${i}`} x1={cx} y1={cy} x2={x} y2={y} stroke={GRID} strokeWidth={1} />)
    }
    const data = [0.9, 0.6, 0.8, 0.55, 0.85, 0.65].map((v, i) => spoke(i, R * v).join(','))
    els.push(
      <polygon
        key="rd"
        points={data.join(' ')}
        fill={C1}
        fillOpacity={0.25}
        stroke={C1}
        strokeWidth={1.8}
      />,
    )
  } else if (kind === 'pie3D') {
    const cx = W / 2
    const cy = H / 2 - 3
    const rx = H / 2 - 8
    const ry = rx * 0.55
    const d = 7
    els.push(
      <path
        key="rim"
        d={`M ${cx - rx} ${cy} A ${rx} ${ry} 0 0 0 ${cx + rx} ${cy} v ${d} A ${rx} ${ry} 0 0 1 ${cx - rx} ${cy + d} Z`}
        fill="#2F5496"
      />,
    )
    els.push(<ellipse key="p1" cx={cx} cy={cy} rx={rx} ry={ry} fill={C1} />)
    const a = ((-90 + 120) * Math.PI) / 180
    els.push(
      <path
        key="p2"
        d={`M ${cx} ${cy} L ${cx} ${cy - ry} A ${rx} ${ry} 0 0 1 ${cx + Math.cos(a) * rx} ${cy + Math.sin(a) * ry} Z`}
        fill={C2}
      />,
    )
  } else {
    const cx = W / 2
    const cy = H / 2
    const r = H / 2 - 6
    els.push(<circle key="p1" cx={cx} cy={cy} r={r} fill={C1} />)
    const a = ((-90 + 120) * Math.PI) / 180
    els.push(
      <path
        key="p2"
        d={`M ${cx} ${cy} L ${cx} ${cy - r} A ${r} ${r} 0 0 1 ${cx + Math.cos(a) * r} ${cy + Math.sin(a) * r} Z`}
        fill={C2}
      />,
    )
    if (kind === 'doughnut')
      els.push(<circle key="ph" cx={cx} cy={cy} r={r * 0.5} fill="var(--surface, #fff)" />)
  }

  return (
    <svg width={width} height={(width * H) / W} viewBox={`0 0 ${W} ${H}`} aria-hidden>
      {els}
    </svg>
  )
}

interface Props {
  current?: ChartKind
  onConfirm: (kind: ChartKind) => void
  onClose: () => void
}

export function ChartTypeDialog({ current, onConfirm, onClose }: Props) {
  const { t } = useI18n()
  const [selected, setSelected] = useState<ChartKind>(current ?? 'bar')

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal chart-type-modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonChangeChartType')}</h2>
        <div className="chart-type-grid">
          {KINDS.map((k) => (
            <button
              key={k.kind}
              className={`chart-type-card ${selected === k.kind ? 'active' : ''}`}
              onClick={() => setSelected(k.kind)}
              onDoubleClick={() => onConfirm(k.kind)}
            >
              <ChartKindThumb kind={k.kind} />
              <span className="chart-type-label">{t(k.label)}</span>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('paneCancel')}
          </button>
          <button className="btn-primary" onClick={() => onConfirm(selected)}>
            {t('paneOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

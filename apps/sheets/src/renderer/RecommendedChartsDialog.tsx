import type { ChartRecommendations, RecommendedKind } from '../domain/chart-recommend'
import { useI18n, type StringKey } from './i18n/locale'

/// Excel's Insert → Recommended Charts: ranked suggestions with a mini
/// preview sketched from the actual selection, one click to insert.

const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47']
const W = 180
const H = 104
const PAD = 8

const KIND_LABEL: Record<RecommendedKind, StringKey> = {
  column: 'appChartColumn',
  bar: 'appChartBar',
  line: 'appChartLine',
  area: 'appChartTypeArea',
  pie: 'appChartPie',
  doughnut: 'appChartDoughnut',
  scatter: 'appChartScatter',
  radar: 'appChartRadar',
  combo: 'appChartCombo',
}

const REASON_LABEL: Record<ChartRecommendations['items'][number]['reason'], StringKey> = {
  time: 'dlgRecoReasonTime',
  proportion: 'dlgRecoReasonProportion',
  correlation: 'dlgRecoReasonCorrelation',
  comparison: 'dlgRecoReasonComparison',
  manyPoints: 'dlgRecoReasonManyPoints',
  longLabels: 'dlgRecoReasonLongLabels',
  mixedScales: 'dlgRecoReasonMixedScales',
}

interface PreviewSeries {
  color: string
  points: number[]
}

/// Values normalized to 0..1 with a shared scale, capped for legibility.
function previewSeries(parsed: ChartRecommendations['parsed'], maxPoints: number): PreviewSeries[] {
  const series = parsed.series.slice(0, 3)
  const sliced = series.map((entry) => entry.values.slice(0, maxPoints))
  const flat = sliced.flat()
  const top = Math.max(...flat, 0)
  const bottom = Math.min(...flat, 0)
  const span = top - bottom || 1
  return sliced.map((values, index) => ({
    color: PALETTE[index % PALETTE.length] ?? '#4472c4',
    points: values.map((value) => (value - bottom) / span),
  }))
}

function linePath(points: number[], plotWidth: number): string {
  const step = points.length > 1 ? plotWidth / (points.length - 1) : 0
  return points
    .map(
      (value, index) =>
        `${index === 0 ? 'M' : 'L'}${(PAD + index * step).toFixed(1)},${(H - PAD - value * (H - PAD * 2)).toFixed(1)}`,
    )
    .join(' ')
}

function pieSlices(values: number[]): { d: string; color: string }[] {
  const total = values.reduce((sum, value) => sum + Math.max(value, 0), 0) || 1
  const cx = W / 2
  const cy = H / 2
  const r = H / 2 - PAD
  let angle = -Math.PI / 2
  return values.slice(0, 8).map((value, index) => {
    const sweep = (Math.max(value, 0) / total) * Math.PI * 2
    const from = angle
    angle += sweep
    const x1 = cx + r * Math.cos(from)
    const y1 = cy + r * Math.sin(from)
    const x2 = cx + r * Math.cos(angle)
    const y2 = cy + r * Math.sin(angle)
    return {
      d: `M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${x2.toFixed(1)},${y2.toFixed(1)} Z`,
      color: PALETTE[index % PALETTE.length] ?? '#4472c4',
    }
  })
}

function ChartPreview({
  kind,
  parsed,
}: {
  readonly kind: RecommendedKind
  readonly parsed: ChartRecommendations['parsed']
}): React.JSX.Element {
  const plotWidth = W - PAD * 2
  const plotHeight = H - PAD * 2
  const content = (): React.JSX.Element => {
    switch (kind) {
      case 'pie':
      case 'doughnut': {
        const first = parsed.series[0]
        return (
          <>
            {pieSlices(first?.values ?? []).map((slice, index) => (
              <path key={index} d={slice.d} fill={slice.color} />
            ))}
            {kind === 'doughnut' && (
              <circle cx={W / 2} cy={H / 2} r={(H / 2 - PAD) * 0.55} fill="var(--surface, #fff)" />
            )}
          </>
        )
      }
      case 'bar': {
        const series = previewSeries(parsed, 6)
        const groups = series[0]?.points.length ?? 0
        const band = plotHeight / Math.max(groups, 1)
        const barHeight = Math.min(10, (band * 0.7) / series.length)
        return (
          <>
            {series.map((entry, seriesIndex) =>
              entry.points.map((value, index) => (
                <rect
                  key={`${seriesIndex}-${index}`}
                  x={PAD}
                  y={PAD + index * band + seriesIndex * barHeight + band * 0.15}
                  width={Math.max(2, value * plotWidth)}
                  height={barHeight}
                  fill={entry.color}
                />
              )),
            )}
          </>
        )
      }
      case 'line':
      case 'area': {
        const series = previewSeries(parsed, 12)
        return (
          <>
            {series.map((entry, index) => (
              <g key={index}>
                {kind === 'area' && (
                  <path
                    d={`${linePath(entry.points, plotWidth)} L${W - PAD},${H - PAD} L${PAD},${H - PAD} Z`}
                    fill={entry.color}
                    opacity={0.35}
                  />
                )}
                <path
                  d={linePath(entry.points, plotWidth)}
                  fill="none"
                  stroke={entry.color}
                  strokeWidth={2}
                />
              </g>
            ))}
          </>
        )
      }
      case 'scatter': {
        const [xs, ys] = previewSeries(parsed, 24)
        const points = ys ?? xs
        const xPoints = ys ? (xs?.points ?? []) : []
        return (
          <>
            {(points?.points ?? []).map((value, index) => (
              <circle
                key={index}
                cx={
                  PAD +
                  (xPoints[index] ?? index / Math.max((points?.points.length ?? 1) - 1, 1)) *
                    plotWidth
                }
                cy={H - PAD - value * plotHeight}
                r={2.5}
                fill={PALETTE[0]}
              />
            ))}
          </>
        )
      }
      case 'radar': {
        const series = previewSeries(parsed, 8)
        const count = series[0]?.points.length ?? 3
        const cx = W / 2
        const cy = H / 2
        const radius = H / 2 - PAD
        const vertex = (value: number, index: number): string => {
          const angle = -Math.PI / 2 + (index / count) * Math.PI * 2
          return `${(cx + radius * value * Math.cos(angle)).toFixed(1)},${(cy + radius * value * Math.sin(angle)).toFixed(1)}`
        }
        return (
          <>
            <polygon
              points={Array.from({ length: count }, (_, index) => vertex(1, index)).join(' ')}
              fill="none"
              stroke="var(--border, #ccc)"
            />
            {series.map((entry, seriesIndex) => (
              <polygon
                key={seriesIndex}
                points={entry.points.map((value, index) => vertex(value, index)).join(' ')}
                fill={entry.color}
                opacity={0.4}
                stroke={entry.color}
              />
            ))}
          </>
        )
      }
      case 'combo': {
        const series = previewSeries(parsed, 8)
        const bars = series[0]
        const line = series[1] ?? series[0]
        const groups = bars?.points.length ?? 0
        const band = plotWidth / Math.max(groups, 1)
        return (
          <>
            {(bars?.points ?? []).map((value, index) => (
              <rect
                key={index}
                x={PAD + index * band + band * 0.2}
                y={H - PAD - value * plotHeight}
                width={band * 0.6}
                height={Math.max(2, value * plotHeight)}
                fill={PALETTE[0]}
              />
            ))}
            {line && (
              <path
                d={linePath(line.points, plotWidth)}
                fill="none"
                stroke={PALETTE[1]}
                strokeWidth={2}
              />
            )}
          </>
        )
      }
      default: {
        const series = previewSeries(parsed, 8)
        const groups = series[0]?.points.length ?? 0
        const band = plotWidth / Math.max(groups, 1)
        const barWidth = Math.min(14, (band * 0.7) / series.length)
        return (
          <>
            {series.map((entry, seriesIndex) =>
              entry.points.map((value, index) => (
                <rect
                  key={`${seriesIndex}-${index}`}
                  x={PAD + index * band + seriesIndex * barWidth + band * 0.15}
                  y={H - PAD - value * plotHeight}
                  width={barWidth}
                  height={Math.max(2, value * plotHeight)}
                  fill={entry.color}
                />
              )),
            )}
          </>
        )
      }
    }
  }
  return (
    <svg viewBox={`0 0 ${W} ${H}`} aria-hidden="true">
      <line x1={PAD} y1={H - PAD} x2={W - PAD} y2={H - PAD} stroke="var(--border, #ccc)" />
      {content()}
    </svg>
  )
}

export function RecommendedChartsDialog({
  recommendations,
  onPick,
  onClose,
}: {
  readonly recommendations: ChartRecommendations
  readonly onPick: (kind: RecommendedKind) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog recommended-charts-dialog"
        role="dialog"
        aria-label={t('dlgRecoTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgRecoTitle')}</header>
        <section className="dialog-body">
          <p className="dialog-note">{t('dlgRecoNote')}</p>
          <div className="recommended-charts-grid">
            {recommendations.items.map((item) => (
              <button
                key={item.kind}
                type="button"
                onClick={() => {
                  onPick(item.kind)
                  onClose()
                }}
              >
                <ChartPreview kind={item.kind} parsed={recommendations.parsed} />
                <strong>{t(KIND_LABEL[item.kind])}</strong>
                <span>{t(REASON_LABEL[item.reason])}</span>
              </button>
            ))}
          </div>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}

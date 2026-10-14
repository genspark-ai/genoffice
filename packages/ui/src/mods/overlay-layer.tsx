import { slotsAt, type FeatureRecord, type OverlayShape } from '@genoffice/agent-core'
import './mods.css'

/**
 * The overlay layer (K4): what features draw over the content area.
 *
 * It is inert by design — no pointer events, no focus, aria-hidden — so it can
 * only guide the eye while every press still reaches the app underneath. Colours
 * come from the tone classes in mods.css, which point at the app's tokens, so
 * light and dark both work without the feature knowing anything about them.
 */

function percent(fraction: number): string {
  return (fraction * 100).toFixed(3) + '%'
}

function toneClass(tone: string | undefined): string {
  return 'slot-overlay-tone-' + (tone ?? 'accent')
}

function Shape({ shape }: { shape: OverlayShape }): React.JSX.Element | null {
  const className = toneClass(shape.tone)
  switch (shape.kind) {
    case 'frame':
      return (
        <g className={className}>
          <rect
            className={'slot-overlay-frame' + (shape.dashed ? ' slot-overlay-frame--dashed' : '')}
            x={percent(shape.x)}
            y={percent(shape.y)}
            width={percent(shape.width)}
            height={percent(shape.height)}
            rx="6"
          />
          {shape.label ? (
            <text className="slot-overlay-label" x={percent(shape.x)} y={percent(shape.y)} dy="-4">
              {shape.label}
            </text>
          ) : null}
        </g>
      )
    case 'line':
      return (
        <line
          className={
            className + ' slot-overlay-line' + (shape.dashed ? ' slot-overlay-line--dashed' : '')
          }
          x1={percent(shape.x1)}
          y1={percent(shape.y1)}
          x2={percent(shape.x2)}
          y2={percent(shape.y2)}
        />
      )
    case 'dot':
      return (
        <g className={className}>
          <circle className="slot-overlay-dot" cx={percent(shape.x)} cy={percent(shape.y)} r="4" />
          {shape.label ? (
            <text className="slot-overlay-label" x={percent(shape.x)} y={percent(shape.y)} dx="10">
              {shape.label}
            </text>
          ) : null}
        </g>
      )
    case 'label':
      return (
        <text
          className={className + ' slot-overlay-label'}
          x={percent(shape.x)}
          y={percent(shape.y)}
        >
          {shape.value}
        </text>
      )
    default:
      return null
  }
}

/** Every overlay registered, oldest first. Renders nothing when there is none. */
export function OverlayLayer({
  features,
}: {
  features: FeatureRecord[]
}): React.JSX.Element | null {
  const overlays = slotsAt(features, 'home.overlay')
  if (!overlays.length) return null
  return (
    <div className="slot-overlay" aria-hidden="true">
      {overlays.map((feature) => {
        const element = feature.slot?.element
        if (!element || element.type !== 'overlay') return null
        return (
          <svg key={feature.id} className="slot-overlay-svg" width="100%" height="100%">
            {element.shapes.map((shape, index) => (
              <Shape key={index} shape={shape} />
            ))}
          </svg>
        )
      })}
    </div>
  )
}

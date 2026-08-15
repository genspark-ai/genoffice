/**
 * 2.2 Fill / stroke resolution — converts pptx-engine Fill/Stroke (final values after
 * inheritance) into RenderFill/RenderStroke the render layer can consume directly
 * (px line width, angle in degrees, image dataUrl).
 */
import type { Fill, Stroke, ShadowEffect } from '@genoffice/pptx-engine'
import type { RenderFill, RenderStroke, RenderShadow } from './render-tree'
import { emuToPx, EMU_PER_PT, type Viewport } from './coords'

/** Lookup function for image mediaRef → dataUrl (injected by the caller, may lazy-load). */
export type MediaResolver = (mediaRef: string) => string | undefined

export function resolveFill(
  fill: Fill | undefined,
  vp: Viewport,
  media?: MediaResolver,
): RenderFill {
  if (!fill) return { kind: 'none' }
  switch (fill.type) {
    case 'none':
      return { kind: 'none' }
    case 'solid':
      return { kind: 'solid', color: fill.color }
    case 'gradient':
      return {
        kind: 'gradient',
        stops: fill.stops.map((s) => ({ pos: s.pos, color: s.color })),
        angleDeg: fill.angle != null ? fill.angle / 60000 : 0,
        ...(fill.path ? { radial: true } : {}),
        ...(fill.path && fill.fillTo
          ? {
              center: {
                x: (fill.fillTo.l + (1 - fill.fillTo.r)) / 2,
                y: (fill.fillTo.t + (1 - fill.fillTo.b)) / 2,
              },
            }
          : {}),
      }
    case 'image':
      return {
        kind: 'image',
        dataUrl: media?.(fill.mediaRef),
        mode: fill.mode ?? 'stretch',
      }
    case 'pattern':
      // Pattern fills are approximated with the foreground color for now (Phase 2 skips 8x8 tile replication)
      return { kind: 'solid', color: fill.fg }
    default:
      return { kind: 'none' }
  }
}

export function resolveStroke(stroke: Stroke | undefined, vp: Viewport): RenderStroke | undefined {
  if (!stroke) return undefined
  const rf = stroke.fill
  let color = '#000000'
  if (rf.type === 'solid') color = rf.color
  else if (rf.type === 'none') return undefined
  const widthPx = Math.max(emuToPx(stroke.width || 12700, vp.scale), 0.5)
  const widthPt = (stroke.width || 12700) / EMU_PER_PT
  const dash = dashPreset(stroke.dash, widthPx)
  return {
    color,
    widthPx,
    widthPt,
    ...(dash ? { dash } : {}),
    ...(stroke.dash && stroke.dash !== 'solid' ? { dashPreset: stroke.dash } : {}),
  }
}

/** Outer shadow EMU/angle → px offsets (OOXML dir is clockwise, y-down, matching canvas). */
export function resolveGlow(
  glow: import('@genoffice/pptx-engine').GlowEffect | undefined,
  vp: Viewport,
): import('./render-tree').RenderGlow | undefined {
  if (!glow) return undefined
  return { color: glow.color, blurPx: emuToPx(glow.radius, vp.scale) }
}

export function resolveShadow(
  shadow: ShadowEffect | undefined,
  vp: Viewport,
): RenderShadow | undefined {
  if (!shadow) return undefined
  const distPx = emuToPx(shadow.dist, vp.scale)
  const rad = (shadow.dirDeg * Math.PI) / 180
  return {
    color: shadow.color,
    blurPx: emuToPx(shadow.blurRad, vp.scale),
    offsetX: Math.cos(rad) * distPx,
    offsetY: Math.sin(rad) * distPx,
  }
}

/** OOXML preset dash name → canvas dash array (relative to line width). */
function dashPreset(name: string | undefined, w: number): number[] | undefined {
  if (!name || name === 'solid') return undefined
  const u = w
  switch (name) {
    case 'dot':
    case 'sysDot':
      return [u, u]
    case 'dash':
    case 'sysDash':
      return [4 * u, 3 * u]
    case 'lgDash':
      return [8 * u, 3 * u]
    case 'dashDot':
    case 'sysDashDot':
      return [4 * u, 3 * u, u, 3 * u]
    case 'lgDashDot':
      return [8 * u, 3 * u, u, 3 * u]
    case 'lgDashDotDot':
      return [8 * u, 3 * u, u, 3 * u, u, 3 * u]
    case 'dashDotDot':
    case 'sysDashDotDot':
      return [4 * u, 3 * u, u, 3 * u, u, 3 * u]
    default:
      return undefined
  }
}

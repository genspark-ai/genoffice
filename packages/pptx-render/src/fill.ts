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
        ...(fill.path ? { radial: true, path: fill.path } : {}),
        ...(fill.path && fill.fillTo
          ? {
              center: {
                x: (fill.fillTo.l + (1 - fill.fillTo.r)) / 2,
                y: (fill.fillTo.t + (1 - fill.fillTo.b)) / 2,
              },
            }
          : {}),
      }
    case 'image': {
      // Tile natural size: PowerPoint lays dpi-less bitmaps out at 144dpi (measured on
      // page_transparent_bitmap: 94px tile = 0.653in), i.e. 2/3 of a 96dpi unit
      const pxPerImagePx = vp.scale * (96 / 144)
      return {
        kind: 'image',
        dataUrl: media?.(fill.mediaRef),
        mode: fill.mode ?? 'stretch',
        ...(fill.alpha != null ? { alpha: fill.alpha } : {}),
        ...(fill.fillRect ? { fillRect: fill.fillRect } : {}),
        ...(fill.duotone ? { duotone: fill.duotone } : {}),
        ...(fill.lum ? { lum: fill.lum } : {}),
        ...(fill.clrChange ? { clrChange: fill.clrChange } : {}),
        ...(fill.tile
          ? {
              tile: {
                scaleX: pxPerImagePx * fill.tile.sx,
                scaleY: pxPerImagePx * fill.tile.sy,
                txPx: emuToPx(fill.tile.tx, vp.scale),
                tyPx: emuToPx(fill.tile.ty, vp.scale),
                algn: fill.tile.algn,
              },
            }
          : {}),
      }
    }
    case 'pattern':
      // 8x8 preset mask, one mask pixel per 96dpi pixel (PowerPoint measured on n820786)
      return {
        kind: 'pattern',
        preset: fill.preset,
        fg: fill.fg,
        bg: fill.bg,
        cellPx: 8 * vp.scale,
      }
    default:
      return { kind: 'none' }
  }
}

export function resolveStroke(stroke: Stroke | undefined, vp: Viewport): RenderStroke | undefined {
  if (!stroke) return undefined
  const rf = stroke.fill
  let color = '#000000'
  let gradient: RenderStroke['gradient']
  if (rf.type === 'solid') color = rf.color
  else if (rf.type === 'gradient' && rf.stops.length) {
    gradient = {
      stops: rf.stops.map((s) => ({ pos: s.pos, color: s.color })),
      angleDeg: rf.angle != null ? rf.angle / 60000 : 0,
    }
    color = rf.stops[0]!.color
  } else if (rf.type === 'none') return undefined
  const widthPx = Math.max(emuToPx(stroke.width || 12700, vp.scale), 0.5)
  const widthPt = (stroke.width || 12700) / EMU_PER_PT
  const dash = dashPreset(stroke.dash, widthPx)
  const capMap = { flat: 'butt', round: 'round', square: 'square' } as const
  return {
    color,
    widthPx,
    widthPt,
    ...(dash ? { dash } : {}),
    ...(stroke.dash && stroke.dash !== 'solid' ? { dashPreset: stroke.dash } : {}),
    ...(stroke.cap ? { cap: capMap[stroke.cap] } : {}),
    ...(stroke.join ? { join: stroke.join } : {}),
    ...(stroke.compound ? { compound: stroke.compound } : {}),
    ...(gradient ? { gradient } : {}),
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

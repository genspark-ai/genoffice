/**
 * 2.1 Coordinate system — EMU → px conversion + slide viewport mapping.
 *
 * OOXML uses EMU (English Metric Unit): 1 inch = 914400 EMU, 1 px@96dpi = 9525 EMU.
 * Rendering maps EMU geometry to target canvas pixels; the slide size plus the
 * target viewport size determine one uniform scale factor, keeping proportions
 * without stretching.
 *
 * Angles: OOXML rot unit = 1/60000 degree; clockwise positive (matching Canvas/CSS).
 */
import type { EmuRect, Transform, SlideSize } from '@genoffice/pptx-engine'

export const EMU_PER_INCH = 914400
export const EMU_PER_PX_96 = 9525 // 914400 / 96
export const EMU_PER_PT = 12700 // 914400 / 72

/** EMU → px (under the given scale). With scale=1, converts at the 96dpi baseline. */
export function emuToPx(emu: number, scale = 1): number {
  return (emu / EMU_PER_PX_96) * scale
}

/** pt → px (for font sizes), 96dpi: 1pt = 96/72 px. */
export function ptToPx(pt: number, scale = 1): number {
  return ((pt * 96) / 72) * scale
}

/** OOXML rot (1/60000 degree) → degrees. */
export function rotToDeg(rot: number): number {
  return rot / 60000
}

/** OOXML rot (1/60000 degree) → radians. */
export function rotToRad(rot: number): number {
  return (rotToDeg(rot) * Math.PI) / 180
}

export interface Viewport {
  /** Target canvas width (px) */
  widthPx: number
  /** Target canvas height (px) */
  heightPx: number
  /** Uniform EMU→px scale factor (96dpi baseline included) */
  scale: number
}

/**
 * Computes a proportional viewport from the slide size + target width.
 * Pass fitWidthPx: the canvas fills that width proportionally, height follows the slide ratio.
 * Degenerate inputs (non-finite or <= 0) fall back to the default slide size
 * 9144000x6858000 EMU per axis so the result stays finite.
 */
export function makeViewport(size: SlideSize, fitWidthPx: number): Viewport {
  // Default 10in x 7.5in slide in EMU.
  const DEFAULT_CX_EMU = 9144000
  const DEFAULT_CY_EMU = 6858000
  const isPositiveFinite = (v: number): boolean => Number.isFinite(v) && v > 0
  const safeCx = isPositiveFinite(size.cx) ? size.cx : DEFAULT_CX_EMU
  const safeCy = isPositiveFinite(size.cy) ? size.cy : DEFAULT_CY_EMU
  // Fall back to the natural width (scale 1) when the target width is degenerate.
  const safeFitWidthPx = isPositiveFinite(fitWidthPx) ? fitWidthPx : safeCx / EMU_PER_PX_96
  const baseWidthPx = safeCx / EMU_PER_PX_96 // the slide's "natural" pixel width at 96dpi
  const scale = safeFitWidthPx / baseWidthPx
  return {
    widthPx: safeFitWidthPx,
    heightPx: (safeCy / EMU_PER_PX_96) * scale,
    scale,
  }
}

/** A converted pixel rectangle. */
export interface PxRect {
  x: number
  y: number
  w: number
  h: number
}

/** EMU rect → px rect. */
export function rectToPx(r: EmuRect, vp: Viewport): PxRect {
  return {
    x: emuToPx(r.x, vp.scale),
    y: emuToPx(r.y, vp.scale),
    w: emuToPx(r.cx, vp.scale),
    h: emuToPx(r.cy, vp.scale),
  }
}

/** Converted placement info including rotation/flip. Rotation pivots on the rect center. */
export interface PlacedBox extends PxRect {
  /** Rotation angle (degrees, clockwise positive) */
  rotationDeg: number
  flipH: boolean
  flipV: boolean
  /** Rotation center (px, absolute coordinates) */
  centerX: number
  centerY: number
}

/** Parent coordinate mapping for nested groups: converted top-left offset + ext/chExt scaling. */
export interface ParentPlacement {
  x: number
  y: number
  /** Group ext/chExt scaling (default 1): baked into child geometry at placement; font sizes don't scale */
  scaleX?: number
  scaleY?: number
}

/**
 * Transform (EMU geometry + rotation/flip) → pixel placement box.
 * parent: for nested groups, the parent's converted top-left offset + group scaling (default none).
 */
export function placeTransform(
  t: Transform,
  vp: Viewport,
  parent: ParentPlacement = { x: 0, y: 0 },
): PlacedBox {
  const r = rectToPx(t.offset, vp)
  const sx = parent.scaleX ?? 1
  const sy = parent.scaleY ?? 1
  const x = r.x * sx + parent.x
  const y = r.y * sy + parent.y
  const w = r.w * sx
  const h = r.h * sy
  return {
    x,
    y,
    w,
    h,
    rotationDeg: rotToDeg(t.rot),
    flipH: t.flipH,
    flipV: t.flipV,
    centerX: x + w / 2,
    centerY: y + h / 2,
  }
}

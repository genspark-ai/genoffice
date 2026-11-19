import type { ImageDpi } from '@genoffice/pptx-render'
import type { Size } from './video-size'

/** PowerPoint for Mac lays untagged bitmaps out at 2 px per point; Windows at 96 dpi. */
export const DEFAULT_PICTURE_DPI = process.platform === 'darwin' ? 144 : 96

const EMU_PER_PT = 12700
/** Measured cap: every oversize landscape picture came in exactly 8.5 in wide. */
const OVERSIZE_FIT_WIDTH_EMU = 612 * EMU_PER_PT

export interface Frame {
  x: number
  y: number
  cx: number
  cy: number
}

export interface PictureFrameOptions {
  /** Anchor to center on instead of the slide center (drop point). */
  center?: { x: number; y: number }
  /** Smaller side is grown towards this, aspect preserved, never past the fit box. */
  minEmu?: number
}

/**
 * PowerPoint placement for an inserted picture: physical size from the pixel grid
 * and the declared dpi, shrunk (never enlarged) to fit min(8.5 in, slide width) by
 * slide height. PowerPoint's oversize placement follows no rule we could measure,
 * so we center on the slide (or the drop point) either way.
 */
export function pictureFrame(
  slide: Size,
  naturalPx: Size | null,
  dpi: ImageDpi | undefined,
  fallbackDpi = DEFAULT_PICTURE_DPI,
  { center, minEmu = 0 }: PictureFrameOptions = {},
): Frame {
  const boxW = Math.min(slide.width, OVERSIZE_FIT_WIDTH_EMU)
  const known = naturalPx !== null && naturalPx.width > 0 && naturalPx.height > 0
  const w = known ? (naturalPx.width / (dpi?.x ?? fallbackDpi)) * 72 * EMU_PER_PT : boxW
  const h = known ? (naturalPx.height / (dpi?.y ?? fallbackDpi)) * 72 * EMU_PER_PT : (boxW * 3) / 4
  const fit = Math.min(boxW / w, slide.height / h)
  let scale = Math.min(1, fit)
  const smaller = Math.min(w, h) * scale
  if (smaller < minEmu) scale = Math.min(fit, (scale * minEmu) / smaller)
  const cx = Math.round(w * scale)
  const cy = Math.round(h * scale)
  const c = center ?? { x: slide.width / 2, y: slide.height / 2 }
  return { x: Math.round(c.x - cx / 2), y: Math.round(c.y - cy / 2), cx, cy }
}

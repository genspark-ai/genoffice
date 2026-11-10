/**
 * Insert-menu default frames measured on PowerPoint for Mac (960 x 540 and 720 x 540 pt decks):
 *  - Equation: a 1 in square text box centered on the slide, wrap="none" + spAutoFit so it
 *    follows the typed math, Cambria Math italic at the 18 pt body default.
 *  - Chart / SmartArt: a graphic frame 2/3 of the slide width in a 3:2 aspect, centered
 *    (640 x 426.7 pt on 16:9, 480 x 320 pt on 4:3 — width-relative, not height-relative).
 *  - Slide / Section Zoom: thumbnails at 1/4 of the slide size, centered; several tiles
 *    cascade 12 pt apart in both axes around the center, later tiles on top.
 *  - Summary Zoom: a grid of slide-aspect tiles inside the body placeholder of the new
 *    "Title and Content" slide, 90% fill, gap = tile height / 15, last row left-aligned.
 */
import { PX_PER_INCH } from './app-constants'

const PX_PER_PT = PX_PER_INCH / 72

export interface InsertFrame {
  x: number
  y: number
  w: number
  h: number
}

export const EQUATION_BOX_PX = Math.round(72 * PX_PER_PT)
export const EQUATION_FONT_PT = 18
export const EQUATION_FONT_FAMILY = 'Cambria Math'
export const EQUATION_BODY_PR = { wrap: 'none', autoFit: 'resize' } as const

/** widthPx is always the FIT_WIDTH viewport; scale = viewport px per 96 dpi px, so absolute pt sizes multiply by it */
type SlideSize = { widthPx: number; heightPx: number; scale: number }

function centered(slide: SlideSize, w: number, h: number): InsertFrame {
  return {
    x: Math.round((slide.widthPx - w) / 2),
    y: Math.round((slide.heightPx - h) / 2),
    w,
    h,
  }
}

export function equationInsertFrame(slide: SlideSize): InsertFrame {
  const box = Math.round(EQUATION_BOX_PX * slide.scale)
  return centered(slide, box, box)
}

/** Chart and SmartArt share one frame: 2/3 of the slide width, 3:2 aspect, centered. */
export function graphicFrameInsertFrame(slide: SlideSize): InsertFrame {
  const w = Math.round((slide.widthPx * 2) / 3)
  return centered(slide, w, Math.round((w * 2) / 3))
}

export function zoomInsertFrame(slide: SlideSize): InsertFrame {
  return centered(slide, Math.round(slide.widthPx / 4), Math.round(slide.heightPx / 4))
}

export const ZOOM_CASCADE_STEP_PT = 12

export function zoomCascadeFrames(slide: SlideSize, n: number): InsertFrame[] {
  const base = zoomInsertFrame(slide)
  const step = ZOOM_CASCADE_STEP_PT * PX_PER_PT * slide.scale
  return Array.from({ length: n }, (_, i) => {
    const d = (i - (n - 1) / 2) * step
    return { x: Math.round(base.x + d), y: Math.round(base.y + d), w: base.w, h: base.h }
  })
}

/** Tiles fill rows left to right; the grid is centered in `box`, so a short last row stays left-aligned. */
export function summaryZoomLayout(n: number, box: InsertFrame, aspect: number): InsertFrame[] {
  if (n <= 0) return []
  let best = { cols: 1, rows: n, tileW: 0 }
  for (let cols = 1; cols <= n; cols++) {
    const rows = Math.ceil(n / cols)
    const tileW = Math.min((0.9 * box.w) / cols, ((0.9 * box.h) / rows) * aspect)
    if (tileW > best.tileW) best = { cols, rows, tileW }
  }
  const { cols, rows, tileW } = best
  const tileH = tileW / aspect
  const gap = tileH / 15
  const x0 = box.x + (box.w - (cols * tileW + (cols - 1) * gap)) / 2
  const y0 = box.y + (box.h - (rows * tileH + (rows - 1) * gap)) / 2
  return Array.from({ length: n }, (_, i) => ({
    x: Math.round(x0 + (i % cols) * (tileW + gap)),
    y: Math.round(y0 + Math.floor(i / cols) * (tileH + gap)),
    w: Math.round(tileW),
    h: Math.round(tileH),
  }))
}

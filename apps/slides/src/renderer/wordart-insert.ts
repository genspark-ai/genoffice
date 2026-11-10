/**
 * Insert > WordArt (PowerPoint parity): PowerPoint for Mac drops a 354.6 x 72.7 pt box centered on
 * the slide with 54 pt text, wrap="none" + spAutoFit so the box follows the text, and opens it for
 * editing with the whole placeholder selected so typing replaces it.
 */
import { PX_PER_INCH } from './app-constants'

const PX_PER_PT = PX_PER_INCH / 72
export const WORDART_W_PX = Math.round(354.6 * PX_PER_PT)
export const WORDART_H_PX = Math.round(72.7 * PX_PER_PT)
export const WORDART_FONT_PT = 54
export const WORDART_BODY_PR = { wrap: 'none', autoFit: 'resize' } as const

export function wordArtInsertSpec(slide: { widthPx: number; heightPx: number }) {
  return {
    x: Math.round((slide.widthPx - WORDART_W_PX) / 2),
    y: Math.round((slide.heightPx - WORDART_H_PX) / 2),
    w: WORDART_W_PX,
    h: WORDART_H_PX,
    bodyPr: WORDART_BODY_PR,
  }
}

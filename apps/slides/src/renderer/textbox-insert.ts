/**
 * Click-to-type text box (PowerPoint parity): a click drops an EMPTY one-line box that grows in
 * width as the user types (wrap="none" + spAutoFit); a drag fixes the width and grows the height
 * (wrap="square" + spAutoFit). PowerPoint for Mac measures the click box at 14.5 x 29 pt.
 */
import type { EditParagraph } from '../shared/ipc'
import { PX_PER_INCH } from './app-constants'
import type { DrawRect } from './draw-shape'

const PX_PER_PT = PX_PER_INCH / 72
export const TEXTBOX_CLICK_W_PX = Math.round(14.5 * PX_PER_PT)
export const TEXTBOX_CLICK_H_PX = Math.round(29 * PX_PER_PT)

export interface TextBoxInsertSpec {
  x: number
  y: number
  w: number
  h: number
  bodyPr: { wrap: 'none' | 'square'; autoFit: 'resize' }
}

/** Draw gesture → box to insert: click = grow-in-width box at the click point, drag = fixed width, one line tall. */
export function textBoxInsertSpec(rect: DrawRect): TextBoxInsertSpec {
  const click = rect.click === true
  return {
    x: Math.round(rect.x),
    y: Math.round(rect.y),
    w: click ? TEXTBOX_CLICK_W_PX : Math.max(Math.round(rect.w), TEXTBOX_CLICK_W_PX),
    h: TEXTBOX_CLICK_H_PX,
    bodyPr: { wrap: click ? 'none' : 'square', autoFit: 'resize' },
  }
}

/** No characters at all (PowerPoint discards a click-to-type box left empty). */
export function paragraphsBlank(paragraphs: EditParagraph[]): boolean {
  return paragraphs.every((p) => p.runs.every((r) => r.text === ''))
}

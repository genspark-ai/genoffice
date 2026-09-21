import { describe, it, expect } from 'vitest'
import {
  TEXTBOX_CLICK_H_PX,
  TEXTBOX_CLICK_W_PX,
  paragraphsBlank,
  textBoxInsertSpec,
} from '../src/renderer/textbox-insert'

const PX_PER_PT = 96 / 72

describe('click-to-type text box geometry (PowerPoint for Mac measurements)', () => {
  it('a single click drops a 14.5 x 29 pt empty box (default insets around one 18 pt line)', () => {
    expect(TEXTBOX_CLICK_W_PX).toBe(Math.round(14.5 * PX_PER_PT))
    expect(TEXTBOX_CLICK_H_PX).toBe(Math.round(29 * PX_PER_PT))
    expect(TEXTBOX_CLICK_W_PX).toBe(19)
    expect(TEXTBOX_CLICK_H_PX).toBe(39)
  })

  it('click: top-left at the pointer, wrap="none" so the box grows in width while typing', () => {
    const spec = textBoxInsertSpec({ x: 100.4, y: 80.6, w: 96, h: 96, click: true })
    expect(spec).toEqual({
      x: 100,
      y: 81,
      w: TEXTBOX_CLICK_W_PX,
      h: TEXTBOX_CLICK_H_PX,
      bodyPr: { wrap: 'none', autoFit: 'resize' },
    })
  })

  it('drag: keeps the drawn width, one line tall, wrap="square" so the height grows instead', () => {
    const spec = textBoxInsertSpec({ x: 40, y: 50, w: 300, h: 200 })
    expect(spec).toEqual({
      x: 40,
      y: 50,
      w: 300,
      h: TEXTBOX_CLICK_H_PX,
      bodyPr: { wrap: 'square', autoFit: 'resize' },
    })
  })

  it('a drag narrower than the click box is widened to it', () => {
    expect(textBoxInsertSpec({ x: 0, y: 0, w: 4, h: 4 }).w).toBe(TEXTBOX_CLICK_W_PX)
  })
})

describe('paragraphsBlank (empty click-to-type boxes are discarded)', () => {
  it('no paragraphs or only empty runs count as blank', () => {
    expect(paragraphsBlank([])).toBe(true)
    expect(paragraphsBlank([{ runs: [{ text: '' }] }])).toBe(true)
    expect(paragraphsBlank([{ runs: [] }, { runs: [{ text: '' }, { text: '' }] }])).toBe(true)
  })

  it('any character, including whitespace, keeps the box', () => {
    expect(paragraphsBlank([{ runs: [{ text: 'a' }] }])).toBe(false)
    expect(paragraphsBlank([{ runs: [{ text: '' }] }, { runs: [{ text: ' ' }] }])).toBe(false)
  })
})

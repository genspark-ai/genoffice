import { describe, it, expect } from 'vitest'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { isPromptPlaceholder, textHitAtPoint } from '../src/renderer/text-hit-area'

const run = (text: string, x: number, widthPx: number) => ({
  text,
  x,
  baselineY: 16,
  fontFamily: 'Arial',
  fontSizePx: 18,
  color: '#000',
  bold: false,
  italic: false,
  underline: false,
  widthPx,
})

function shape(partial: Partial<ShapeRenderNode> = {}, lines = 2): ShapeRenderNode {
  return {
    id: 's',
    sourceId: 's',
    type: 'shape',
    box: { x: 0, y: 0, w: 300, h: 200, rotationDeg: 0 },
    text: {
      lines: Array.from({ length: lines }, (_, i) => ({
        runs: [run('•', 0, 10), run(`line ${i}`, 24, 100)],
        top: 40 + i * 24,
        height: 22,
        paraStart: true,
      })),
      insets: { l: 10, t: 5, r: 10, b: 5 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 48,
      wrap: true,
    },
    ...partial,
  } as ShapeRenderNode
}

const box = { w: 300, h: 200 }

describe('textHitAtPoint (single click on text edits, frame around it selects)', () => {
  it('hits inside a line box spanning the glyph extent', () => {
    expect(textHitAtPoint(shape(), box, { x: 10 + 50, y: 5 + 40 + 10 })).toBe(true)
    // second line
    expect(textHitAtPoint(shape(), box, { x: 10 + 120, y: 5 + 64 + 5 })).toBe(true)
  })

  it('misses the frame around the text (right of the glyphs, above the first line)', () => {
    expect(textHitAtPoint(shape(), box, { x: 250, y: 5 + 40 + 10 })).toBe(false)
    expect(textHitAtPoint(shape(), box, { x: 60, y: 10 })).toBe(false)
    expect(textHitAtPoint(shape(), box, { x: 60, y: 150 })).toBe(false)
  })

  it('pads the glyph extent by a few px so edge clicks still land', () => {
    expect(textHitAtPoint(shape(), box, { x: 10 + 124 + 3, y: 5 + 40 + 10 }, 4)).toBe(true)
    expect(textHitAtPoint(shape(), box, { x: 10 + 124 + 6, y: 5 + 40 + 10 }, 4)).toBe(false)
  })

  it('empty prompt placeholder: the whole frame is text', () => {
    const ph = shape({ placeholder: 'body' }, 0)
    expect(isPromptPlaceholder(ph)).toBe(true)
    expect(textHitAtPoint(ph, box, { x: 280, y: 190 })).toBe(true)
    expect(textHitAtPoint(ph, box, { x: 310, y: 190 })).toBe(false)
    // a plain empty text box has no text region (double-click still edits it)
    expect(textHitAtPoint(shape({}, 0), box, { x: 150, y: 100 })).toBe(false)
    // an empty placeholder usually has no layout object at all
    const bare = shape({ placeholder: 'title' }, 0)
    delete bare.text
    expect(isPromptPlaceholder(bare)).toBe(true)
    expect(textHitAtPoint(bare, box, { x: 150, y: 100 })).toBe(true)
  })

  it('warped text never single-click edits', () => {
    const v = shape()
    v.text!.txWarp = { prst: 'textArchUp' }
    expect(textHitAtPoint(v, box, { x: 60, y: 55 })).toBe(false)
  })

  it('eaVert columns: a click on a column of cells edits, the frame beside the columns selects', () => {
    // two columns 24px wide at x=200 and x=170 (right→left), cells stacked from top 20 to 20+72
    const col = (x: number) => ({
      runs: ['\u7e26', '\u66f8', '\u304d'].map((ch, i) => ({
        ...run(ch, x, 18),
        baselineY: 20 + i * 24 + 16,
      })),
      top: 20,
      height: 72,
      paraStart: x === 200,
    })
    const s = shape({
      text: {
        lines: [col(200), col(170)],
        insets: { l: 10, t: 5, r: 10, b: 5 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 72,
        wrap: true,
        vert: 'eaVert',
      },
    })
    expect(textHitAtPoint(s, box, { x: 10 + 200 + 9, y: 5 + 20 + 40 })).toBe(true)
    expect(textHitAtPoint(s, box, { x: 10 + 170 + 9, y: 5 + 20 + 70 })).toBe(true)
    expect(textHitAtPoint(s, box, { x: 10 + 100, y: 5 + 20 + 40 })).toBe(false)
    expect(textHitAtPoint(s, box, { x: 10 + 200 + 9, y: 5 + 20 + 72 + 10 })).toBe(false)
  })

  it('eaVert rotated Latin word spans the font size left of its rotation anchor', () => {
    const word = { ...run('2024', 212, 60), baselineY: 36, rotate90: true }
    const s = shape({
      text: {
        lines: [{ runs: [word], top: 20, height: 60, paraStart: true }],
        insets: { l: 0, t: 0, r: 0, b: 0 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 60,
        wrap: true,
        vert: 'eaVert',
      },
    })
    expect(textHitAtPoint(s, box, { x: 203, y: 50 }, 0)).toBe(true)
    expect(textHitAtPoint(s, box, { x: 240, y: 50 }, 0)).toBe(false)
  })

  it('whole-block rotation (vert/vert270) keeps the frame-only behavior', () => {
    const s = shape({ text: { ...shape().text!, vert: 'vert' } })
    expect(textHitAtPoint(s, box, { x: 10 + 50, y: 5 + 40 + 10 })).toBe(false)
  })
})

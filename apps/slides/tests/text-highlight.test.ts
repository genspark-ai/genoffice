/**
 * Text highlight (<a:rPr><a:highlight>) → Konva draw conversion: a highlighted glyph
 * run gets a background rect covering its slice of the line box; plain runs get none.
 */
import { describe, it, expect } from 'vitest'
import { layoutGlyphs } from '../src/renderer/konva-adapter'
import type { RenderTextLayout } from '@genoffice/pptx-render'

const baseRun = {
  fontFamily: 'Arial',
  fontSizePx: 20,
  bold: false,
  italic: false,
  underline: false,
}

function layout(): RenderTextLayout {
  return {
    lines: [
      {
        top: 10,
        height: 24,
        runs: [
          {
            ...baseRun,
            text: 'marked',
            x: 5,
            baselineY: 28,
            widthPx: 60,
            color: '#000000',
            highlight: '#FF0000',
          },
          { ...baseRun, text: ' plain', x: 65, baselineY: 28, widthPx: 50, color: '#000000' },
        ],
      },
    ],
    insets: { l: 0, t: 0, r: 0, b: 0 },
    anchor: 'top',
    fontScale: 1,
    contentHeight: 24,
    wrap: true,
  }
}

describe('layoutGlyphs highlight rect', () => {
  it('emits a line-box rect for the highlighted run only', () => {
    const glyphs = layoutGlyphs(layout())
    expect(glyphs).toHaveLength(2)
    expect(glyphs[0]!.highlight).toEqual({ x: 5, y: 10, w: 60, h: 24, color: '#FF0000' })
    expect(glyphs[1]!.highlight).toBeUndefined()
  })

  it('skips vertical layouts, whose lines are full columns', () => {
    const l = layout()
    l.vert = 'eaVert'
    const glyphs = layoutGlyphs(l)
    expect(glyphs[0]!.highlight).toBeUndefined()
  })
})

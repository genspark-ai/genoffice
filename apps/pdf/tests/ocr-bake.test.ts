import { describe, expect, it } from 'vitest'
import { buildOcrBakeInserts, groupLines } from '../src/renderer/ocr-bake'
import { pdfToView } from '../src/renderer/annotations'
import type { OcrWord } from '../src/renderer/ocr-layer'

const geom = { pw: 595, ph: 842, rot: 0, x0: 0, y0: 0 }
const word = (text: string, x: number, y: number, w: number, h: number): OcrWord => ({
  text,
  rect: [x, y, x + w, y + h],
  spaceAfter: true,
})

describe('groupLines', () => {
  it('groups words whose centers share a line, sorting left to right', () => {
    const lines = groupLines([word('b', 100, 700, 20, 12), word('a', 50, 700, 20, 12)])
    expect(lines).toHaveLength(1)
    expect(lines[0]!.map((w) => w.text)).toEqual(['a', 'b'])
  })
  it('splits words on different vertical centers', () => {
    const lines = groupLines([word('top', 50, 780, 20, 12), word('bottom', 50, 700, 20, 12)])
    expect(lines).toHaveLength(2)
  })
})

describe('buildOcrBakeInserts', () => {
  it('emits one white insert per line, origin at the left baseline', () => {
    const inserts = buildOcrBakeInserts(
      0,
      [word('Hello', 50, 700, 40, 12), word('world', 100, 700, 40, 12)],
      geom,
    )
    expect(inserts).toHaveLength(1)
    const insert = inserts[0]!
    expect(insert.pageIndex).toBe(0)
    // 0-255 RGB like every other text insert — [1,1,1] painted near-black
    expect(insert.color).toEqual([255, 255, 255])
    // render mode 3: never painted, whatever a viewer does with the fill
    expect(insert.invisible).toBe(true)
    expect(insert.text).toBe('Hello world')
    expect(insert.origin[0]).toBe(50)
    expect(insert.origin[1]).toBeGreaterThan(700)
    expect(insert.fontSize).toBeGreaterThan(4)
  })

  it('keeps cjk words unspaced', () => {
    const inserts = buildOcrBakeInserts(
      0,
      // cjk overlay words carry spaceAfter=false (the overlay's own no-space-script rule)
      [
        { ...word('\u4f7f', 50, 700, 12, 12), spaceAfter: false },
        { ...word('\u7528', 62, 700, 12, 12), spaceAfter: false },
      ],
      geom,
    )
    expect(inserts[0]!.text).toBe('\u4f7f\u7528')
  })

  it('carries the page rotation', () => {
    const inserts = buildOcrBakeInserts(0, [word('x', 50, 700, 10, 10)], {
      ...geom,
      rot: 90,
    })
    expect(inserts[0]!.rotate).toBe(90)
  })

  it('places the origin through display space on a rotated page', () => {
    // A line that is horizontal in DISPLAY space on a /Rotate 90 page: its
    // user-space box runs vertically (x 80..92, y 100..300), mapping to a
    // display box x 100..300, y 80..92 (y down). The baseline must sit
    // 12*0.22 above the display box bottom on the line's display-left edge;
    // a user-space min-x/min-y origin disagrees with the writer's
    // counter-rotation for every non-zero rotation.
    const g = { pw: 612, ph: 792, rot: 90, x0: 0, y0: 0 }
    const inserts = buildOcrBakeInserts(0, [word('x', 80, 100, 12, 200)], g)
    const [ux, uy] = inserts[0]!.origin
    const [vx, vy] = pdfToView(g, ux, uy)
    expect(vx).toBeCloseTo(100)
    expect(vy).toBeCloseTo(92 - 12 * 0.22)
    expect(inserts[0]!.rotate).toBe(90)
  })
})

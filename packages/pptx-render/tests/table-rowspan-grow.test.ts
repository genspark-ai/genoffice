import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const enginePptx = (name: string) =>
  readFileSync(join(here, '..', '..', 'pptx-engine', 'tests', 'fixtures', name))

describe('table rowSpan growth', () => {
  it('grows spanned rows to fit tall spanning text', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const longText = {
      paragraphs: [
        {
          runs: [
            { text: 'spanning cell text that wraps across many lines '.repeat(8), fontSize: 14 },
          ],
        },
      ],
    }
    const el: any = {
      id: 'tbl_span',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1905000, cy: 476250 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [1905000],
      rowHeights: [238125, 238125],
      rows: [[{ rowSpan: 2, text: longText }], [{ merged: true }]],
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    expect(node.type).toBe('table')
    const cell = node.cells[0]
    expect(cell.rowSpan).toBe(2)
    const layout = cell.text
    expect(layout.contentHeight + layout.insets.t + layout.insets.b).toBeLessThanOrEqual(
      cell.h + 0.01,
    )
    const [y0, y1, y2] = node.gridY as number[]
    const rowHeights = [y1! - y0!, y2! - y1!]
    expect(rowHeights[0]).toBeGreaterThan((238125 * 1280) / deck.size.cx)
    expect(rowHeights[0]).toBeCloseTo(rowHeights[1]!, 6)
  }, 30000)
})

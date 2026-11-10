import { describe, expect, it } from 'vitest'
import {
  addChart,
  addTable,
  commitSaved,
  createBlankPptx,
  materializeSlide,
  openPptx,
  savePptx,
  type OpenedPptx,
} from '../src/index'

const OFF = { x: 0, y: 0, cx: 4000000, cy: 3000000 }

const unsaved = (opened: OpenedPptx) => !!opened.deck.slides[0]!.structureDirty

describe('an insert that reparses the slide', () => {
  it.each([
    ['a table', (o: OpenedPptx) => addTable(o, 0, { rows: 2, cols: 2, offset: OFF })],
    [
      'a chart',
      (o: OpenedPptx) =>
        addChart(o, 0, {
          kind: 'bar',
          categories: ['Q1'],
          series: [{ name: 'S', values: [1] }],
          offset: OFF,
        }),
    ],
  ])('leaves %s unsaved until the deck is saved', async (_, insert) => {
    const opened = await openPptx(await createBlankPptx())
    expect(insert(opened)).not.toBeNull()
    expect(unsaved(opened)).toBe(true)

    await savePptx(opened)
    commitSaved(opened)
    expect(unsaved(opened)).toBe(false)
  })

  it('keeps a clean slide clean', async () => {
    const opened = await openPptx(await createBlankPptx())
    materializeSlide(opened, 0)
    expect(unsaved(opened)).toBe(false)
  })
})

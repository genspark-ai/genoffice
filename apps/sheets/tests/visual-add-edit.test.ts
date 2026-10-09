import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  journalSize,
  recordChartEdit,
  recordSheetRemove,
  recordVisualAdd,
  recordVisualEdit,
  removeVisualAdd,
  reorderVisualAdd,
  toSaveChartEdits,
  toSaveVisualEdits,
  updateVisualAdd,
  toSaveVisualAdds,
} from '../src/renderer/edit-journal'
import { projectVisualEdits } from '../src/renderer/univer-sync'
import {
  workbookChartEditSchema,
  workbookVisualAddSchema,
  workbookVisualEditSchema,
  type WorkbookVisualObject,
} from '../src/shared/desktop-api'

const textBox = (): WorkbookVisualObject => ({
  id: 'added-shape-abc-1',
  sheetId: 'sheet-1',
  kind: 'shape',
  anchor: {
    fromRow: 0,
    fromColumn: 0,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 4,
    toColumn: 4,
    toRowOffset: 0,
    toColumnOffset: 0,
  },
  shapeType: 'rect',
  name: 'TextBox',
  fillColor: '#FFFFFF',
  text: 'Text',
})

describe('updateVisualAdd', () => {
  it('moves, resizes, and retexts a session visual; the save payload follows', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())

    const moved = { ...textBox().anchor, fromRow: 2, toRow: 6, fromColumn: 1, toColumn: 5 }
    expect(updateVisualAdd(journal, 'added-shape-abc-1', { anchor: moved })).toBe(true)
    expect(updateVisualAdd(journal, 'added-shape-abc-1', { text: 'Hello' })).toBe(true)

    const visual = journal.visualAdds[0]
    expect(visual?.anchor).toEqual(moved)
    expect(visual?.text).toBe('Hello')
    expect(visual?.fillColor).toBe('#FFFFFF')

    const additions = toSaveVisualAdds(journal)
    expect(additions).toHaveLength(1)
    expect(additions[0]?.anchor).toEqual(moved)
    expect(additions[0]?.shape).toMatchObject({ text: 'Hello', isTextBox: true })
  })

  it('rejects unknown ids (file visuals go through recordVisualEdit instead)', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    expect(updateVisualAdd(journal, 'file-visual-1', { text: 'nope' })).toBe(false)
    expect(journal.visualAdds[0]?.text).toBe('Text')
  })

  it('drops a session visual outright on removal', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    expect(removeVisualAdd(journal, 'added-shape-abc-1')).toBe(true)
    expect(journal.visualAdds).toHaveLength(0)
    expect(removeVisualAdd(journal, 'added-shape-abc-1')).toBe(false)
  })

  it('projects baked chart state (legend/labels/axis titles/colors) into the save payload', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, {
      ...textBox(),
      id: 'added-chart-abc-1',
      kind: 'chart',
      chart: {
        chartTypes: ['pieChart'],
        title: 'Regional Share',
        legend: 'bottom',
        dataLabels: 'category-percent',
        dataLabelPosition: 'outside-end',
        dataLabelFormat: '0.0%',
        axisTitles: { category: 'Month', value: null },
        series: [
          {
            name: 'S1',
            categories: ['a', 'b'],
            values: [1, 2],
            color: '#4472C4',
            pointColors: [
              { index: 0, color: '#ED7D31' },
              { index: 1, color: '#70AD47' },
            ],
          },
        ],
      },
    })
    const chart = toSaveVisualAdds(journal)[0]?.chart
    expect(chart).toMatchObject({
      chartType: 'pie',
      legend: 'bottom',
      dataLabels: 'category-percent',
      dataLabelPosition: 'outside-end',
      dataLabelFormat: '0.0%',
      axisTitles: { category: 'Month' },
    })
    expect(chart?.axisTitles).not.toHaveProperty('value')
    expect(chart?.series[0]).toMatchObject({
      color: '#4472C4',
      pointColors: { '0': '#ED7D31', '1': '#70AD47' },
    })
  })
})

describe('save-wire string clamping', () => {
  const long = (length: number): string => 'x'.repeat(length)

  it('clamps cell-derived chart strings so a long cell never fails the save schema', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, {
      ...textBox(),
      id: 'added-chart-long-1',
      kind: 'chart',
      chart: {
        chartTypes: ['barChart'],
        barDirection: 'col',
        title: long(400),
        axisTitles: { category: long(400) },
        series: [{ name: long(300), categories: ['short', long(3_000)], values: [1, 2] }],
      },
    })
    const addition = toSaveVisualAdds(journal)[0]
    expect(() => workbookVisualAddSchema.parse(addition)).not.toThrow()
    expect(addition?.chart?.title).toHaveLength(255)
    expect(addition?.chart?.axisTitles?.category).toHaveLength(255)
    const series = addition?.chart?.series[0]
    expect(series?.name).toHaveLength(255)
    expect(series?.categories).toEqual(['short', long(1_024)])
  })

  it('clamps chart-edit series rewrites the same way', () => {
    const journal = createEditJournal()
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      title: long(400),
      series: [{ index: 0, name: long(300), categories: [long(3_000)], values: [1] }],
    })
    const edit = toSaveChartEdits(journal)[0]
    expect(() => workbookChartEditSchema.parse(edit)).not.toThrow()
    expect(edit?.title).toHaveLength(255)
    expect(edit?.series?.[0]?.name).toHaveLength(255)
    expect(edit?.series?.[0]?.categories?.[0]).toHaveLength(1_024)
  })
})

describe('recordVisualEdit', () => {
  const fileImage = (): WorkbookVisualObject => ({
    ...textBox(),
    id: 'visual-3',
    kind: 'image',
    drawingPath: 'xl/drawings/drawing1.xml',
    drawingIndex: 3,
  })

  it('merges moves and removals per visual, removal winning', () => {
    const journal = createEditJournal()
    const moved = { ...fileImage().anchor, fromRow: 8, toRow: 12 }
    expect(recordVisualEdit(journal, fileImage(), { anchor: moved })).toBe(true)
    expect(recordVisualEdit(journal, fileImage(), { remove: true })).toBe(true)
    expect(journalSize(journal)).toBe(1)
    expect(toSaveVisualEdits(journal)).toEqual([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 3,
        remove: true,
        anchor: moved,
      },
    ])
  })

  it('records paint edits alone and merges them with a later move', () => {
    const journal = createEditJournal()
    const shape = { ...fileImage(), kind: 'shape' as const, shapeType: 'rect' }
    expect(recordVisualEdit(journal, shape, { fillColor: '#ABCDEF' })).toBe(true)
    expect(recordVisualEdit(journal, shape, { lineColor: 'none' })).toBe(true)
    const moved = { ...shape.anchor, fromRow: 8, toRow: 12 }
    expect(recordVisualEdit(journal, shape, { anchor: moved })).toBe(true)
    const [edit] = toSaveVisualEdits(journal)
    expect(edit).toEqual({
      drawingPath: 'xl/drawings/drawing1.xml',
      drawingIndex: 3,
      anchor: moved,
      fillColor: '#ABCDEF',
      lineColor: 'none',
    })
    expect(workbookVisualEditSchema.safeParse(edit).success).toBe(true)
    expect(
      workbookVisualEditSchema.safeParse({
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 3,
        lineColor: 'none',
      }).success,
    ).toBe(true)
  })

  it('refuses visuals without a drawing locator', () => {
    const journal = createEditJournal()
    const visual = { ...fileImage(), drawingPath: undefined, drawingIndex: undefined }
    expect(recordVisualEdit(journal, visual, { remove: true })).toBe(false)
    expect(toSaveVisualEdits(journal)).toEqual([])
  })

  it('drops edits on removed sheets from the save payload', () => {
    const journal = createEditJournal()
    recordVisualEdit(journal, fileImage(), { remove: true })
    recordSheetRemove(journal, 'sheet-1')
    expect(toSaveVisualEdits(journal)).toEqual([])
  })
})

describe('arrange properties', () => {
  const fileShape = (): WorkbookVisualObject => ({
    ...textBox(),
    id: 'visual-2',
    drawingPath: 'xl/drawings/drawing1.xml',
    drawingIndex: 2,
  })

  it('merges rotation, flips, alt text, placement, link and z-order into one file edit', () => {
    const journal = createEditJournal()
    expect(recordVisualEdit(journal, fileShape(), { rotation: 90, flipH: true })).toBe(true)
    expect(recordVisualEdit(journal, fileShape(), { altText: 'Logo', zIndex: 0 })).toBe(true)
    expect(recordVisualEdit(journal, fileShape(), { editAs: 'oneCell', hyperlink: '' })).toBe(true)
    expect(recordVisualEdit(journal, fileShape(), { rotation: 0 })).toBe(true)
    expect(journalSize(journal)).toBe(1)
    expect(toSaveVisualEdits(journal)).toEqual([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 2,
        zIndex: 0,
        rotation: 0,
        flipH: true,
        altText: 'Logo',
        editAs: 'oneCell',
        hyperlink: '',
      },
    ])
    expect(workbookVisualEditSchema.safeParse(toSaveVisualEdits(journal)[0]).success).toBe(true)
    recordVisualEdit(journal, fileShape(), { remove: true })
    expect(toSaveVisualEdits(journal)).toEqual([
      { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 2, remove: true },
    ])
  })

  it('session visuals carry the properties into the add payload', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    expect(
      updateVisualAdd(journal, 'added-shape-abc-1', {
        rotation: 45,
        frameSize: { width: 914400, height: 457200 },
        flipV: true,
        altText: 'Note',
        hyperlink: 'https://example.com',
        editAs: 'absolute',
      }),
    ).toBe(true)
    const [addition] = toSaveVisualAdds(journal)
    expect(addition?.shape).toMatchObject({
      rotation: 45,
      frameSize: { width: 914400, height: 457200 },
      flipV: true,
      altText: 'Note',
      hyperlink: 'https://example.com',
      editAs: 'absolute',
    })
    expect(addition?.shape).not.toHaveProperty('flipH')
    expect(workbookVisualAddSchema.safeParse(addition).success).toBe(true)
  })

  it('reorders session visuals within the add list', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    recordVisualAdd(journal, { ...textBox(), id: 'added-shape-abc-2' })
    recordVisualAdd(journal, { ...textBox(), id: 'added-shape-abc-3' })
    expect(reorderVisualAdd(journal, 'added-shape-abc-3', 0)).toBe(true)
    expect(journal.visualAdds.map((visual) => visual.id)).toEqual([
      'added-shape-abc-3',
      'added-shape-abc-1',
      'added-shape-abc-2',
    ])
    expect(reorderVisualAdd(journal, 'missing', 0)).toBe(false)
  })

  it('projects edits onto file visuals and sorts a drawing part by zIndex', () => {
    const journal = createEditJournal()
    const a = { ...fileShape(), id: 'visual-1', drawingIndex: 0 }
    const b = { ...fileShape(), id: 'visual-2', drawingIndex: 1 }
    const c = { ...fileShape(), id: 'visual-3', drawingIndex: 2 }
    recordVisualEdit(journal, a, { zIndex: 2, altText: 'first', flipH: true })
    recordVisualEdit(journal, b, { zIndex: 0 })
    recordVisualEdit(journal, c, { zIndex: 1, remove: true })
    const projected = projectVisualEdits([a, b, c], journal.visualEdits)
    expect(projected.map((visual) => visual.id)).toEqual(['visual-2', 'visual-1'])
    expect(projected[1]).toMatchObject({ altText: 'first', flipH: true })
  })
})

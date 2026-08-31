import type { TableRenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import {
  tableCellAtPoint,
  tableCellOverlayBox,
  tableLocalPointFromStage,
} from '../src/renderer/table-hit'

function makeTable(): TableRenderNode {
  return {
    type: 'table',
    sourceId: 't1',
    id: 'node_t1',
    box: { x: 100, y: 50, w: 120, h: 80, rotationDeg: 0 },
    cells: [
      { x: 0, y: 0, w: 60, h: 40, row: 0, col: 0, fill: { kind: 'solid', color: '#fff' } },
      { x: 60, y: 0, w: 60, h: 40, row: 0, col: 1, fill: { kind: 'solid', color: '#fff' } },
      { x: 0, y: 40, w: 60, h: 40, row: 1, col: 0, fill: { kind: 'solid', color: '#fff' } },
      { x: 60, y: 40, w: 60, h: 40, row: 1, col: 1, fill: { kind: 'solid', color: '#fff' } },
    ],
    gridX: [0, 60, 120],
    gridY: [0, 40, 80],
  } as unknown as TableRenderNode
}

describe('table cell hit testing', () => {
  it('resolves a cell in unrotated table coordinates', () => {
    const table = makeTable()
    expect(tableCellAtPoint(table, { x: 20, y: 10 })).toMatchObject({ row: 0, col: 0 })
    expect(tableCellAtPoint(table, { x: 95, y: 12 })).toMatchObject({ row: 0, col: 1 })
  })

  it('rotates stage hits back into table-local coordinates', () => {
    const table = makeTable()
    const rad = (45 * Math.PI) / 180
    const local = { x: 90, y: 60 }
    const stagePoint = {
      x: 160 + (local.x - 60) * Math.cos(rad) - (local.y - 40) * Math.sin(rad),
      y: 90 + (local.x - 60) * Math.sin(rad) + (local.y - 40) * Math.cos(rad),
    }
    const inverse = tableLocalPointFromStage(stagePoint, {
      x: table.box.x,
      y: table.box.y,
      w: table.box.w,
      h: table.box.h,
      rotationDeg: 45,
    })
    expect(inverse).toMatchObject({ x: 90, y: 60 })
    expect(tableCellAtPoint(table, inverse)).toMatchObject({ row: 1, col: 1 })
  })

  it('places the edit overlay at the rotated cell center', () => {
    const overlay = tableCellOverlayBox(
      { x: 100, y: 50, w: 120, h: 80, rotationDeg: 90 },
      { x: 60, y: 40, w: 60, h: 40 },
    )
    expect(overlay).toMatchObject({ x: 110, y: 100, w: 60, h: 40, rotationDeg: 90 })
  })
})

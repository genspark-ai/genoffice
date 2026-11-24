// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'
import {
  gridGeometry,
  gridlinesCss,
  setActiveDrawingGrid,
  setSnapsToGrid,
  snapToStep,
  snappedDragPointer,
  snappedPoint,
  snapsToGrid,
} from '../src/renderer/editor/drawing-grid'
import {
  drawingGridChanges,
  drawingGridFieldsOf,
} from '../src/renderer/components/DrawingGridDialog'

const pageAt = (left: number, top: number, padding: number) => {
  document.body.innerHTML =
    '<div class="doc-zoom"><div class="doc-page"><p id="p">x</p></div></div>'
  const page = document.querySelector('.doc-page') as HTMLElement
  page.style.padding = `${padding}px`
  page.getBoundingClientRect = () =>
    ({ left, top, right: 0, bottom: 0, width: 0, height: 0 }) as DOMRect
  return document.getElementById('p')!
}
const at = (clientX: number, clientY: number) => ({ clientX, clientY })

afterEach(() => {
  setSnapsToGrid(false)
  setActiveDrawingGrid(null)
})

describe('the drawing grid', () => {
  it('draws a file without a grid as before: 24 px squares from the margins', () => {
    expect(gridGeometry(null)).toEqual({ stepX: 12, stepY: 12, everyX: 2, everyY: 2, origin: null })
    expect(gridlinesCss(null)).toBe('')
  })

  it('takes spacing, shown lines and origin from the file', () => {
    const grid = { verticalSpacing: 156, displayHorizontalEvery: 0, displayVerticalEvery: 2 }
    expect(gridGeometry(grid)).toMatchObject({ stepX: 12, stepY: 10.4, everyX: 0, everyY: 2 })
    expect(gridlinesCss(grid)).toBe(
      '--docs-grid-x:12px; --docs-grid-y:20.8px; --docs-grid-columns:none',
    )
    expect(
      gridlinesCss({
        horizontalSpacing: 360,
        ownOrigin: true,
        horizontalOrigin: 1800,
        verticalOrigin: 1440,
      }),
    ).toBe(
      '--docs-grid-x:48px; --docs-grid-y:24px; --docs-grid-box:padding-box; --docs-grid-at:120px 96px',
    )
  })

  it('reads a zero spacing as the default and keeps a horizontal line when the file says none', () => {
    expect(gridGeometry({ horizontalSpacing: 0, displayVerticalEvery: 0 })).toMatchObject({
      stepX: 12,
      everyY: 1,
    })
  })

  it('snaps to the nearest grid point from an origin', () => {
    expect(snapToStep(131, 100, 24)).toBe(124)
    expect(snapToStep(137, 100, 24)).toBe(148)
    expect(snapToStep(90, 100, 24)).toBe(100)
    expect(snapToStep(5, 0, 0)).toBe(5)
  })

  it('leaves the pointer alone until snapping is switched on, and remembers the switch', () => {
    const el = pageAt(10, 20, 96)
    setActiveDrawingGrid({ horizontalSpacing: 360, verticalSpacing: 360 })
    expect(snapsToGrid()).toBe(false)
    expect(snappedPoint(el, at(140, 150))).toEqual({ x: 140, y: 150 })
    setSnapsToGrid(true)
    expect(snapsToGrid()).toBe(true)
    // margin corner at (106, 116), 24 px steps
    expect(snappedPoint(el, at(140, 150))).toEqual({ x: 130, y: 140 })
  })

  it('moves the pointer of a drag so the corner of the object lands on the grid', () => {
    const el = pageAt(0, 0, 100)
    setActiveDrawingGrid({ horizontalSpacing: 300, verticalSpacing: 300 })
    setSnapsToGrid(true)
    const corner = { left: 113, top: 127 }
    const start = { x: 150, y: 150 }
    const pointer = snappedDragPointer(el, start, corner, at(161, 150))
    // the corner would be at 124: the 20 px grid from 100 puts it at 120
    expect(corner.left + pointer.x - start.x).toBe(120)
    expect(corner.top + pointer.y - start.y).toBe(120)
  })

  it('measures from the page edge when the file gives the grid its own origin', () => {
    const el = pageAt(0, 0, 100)
    setActiveDrawingGrid({
      horizontalSpacing: 300,
      ownOrigin: true,
      horizontalOrigin: 150,
      verticalOrigin: 0,
    })
    setSnapsToGrid(true)
    expect(snappedPoint(el, at(33, 0)).x).toBe(30)
  })
})

describe('the Drawing Grid dialog', () => {
  it('shows a file without a grid at the defaults', () => {
    expect(drawingGridFieldsOf(null)).toEqual({
      hSpacing: 9,
      vSpacing: 9,
      useMargins: true,
      hOrigin: 90,
      vOrigin: 72,
      verticalLines: true,
      vEvery: 2,
      hEvery: 2,
    })
  })

  it('writes nothing when no grid field changed', () => {
    const grid = { verticalSpacing: 156, displayHorizontalEvery: 0 }
    const fields = drawingGridFieldsOf(grid)
    expect(drawingGridChanges(grid, fields, { ...fields })).toBeUndefined()
  })

  it('writes the changed fields over the ones the file has', () => {
    const grid = { verticalSpacing: 156, displayHorizontalEvery: 0, displayVerticalEvery: 2 }
    const fields = drawingGridFieldsOf(grid)
    expect(
      drawingGridChanges(grid, fields, { ...fields, hSpacing: 18, verticalLines: true, vEvery: 3 }),
    ).toEqual({
      horizontalSpacing: 360,
      verticalSpacing: 156,
      displayHorizontalEvery: 3,
      displayVerticalEvery: 2,
    })
  })

  it('writes the origin when the margins are given up and drops the flag when they come back', () => {
    const fields = drawingGridFieldsOf(null)
    const own = drawingGridChanges(null, fields, { ...fields, useMargins: false })
    expect(own).toEqual({ ownOrigin: true, horizontalOrigin: 1800, verticalOrigin: 1440 })
    const back = drawingGridChanges(own!, drawingGridFieldsOf(own!), {
      ...drawingGridFieldsOf(own!),
      useMargins: true,
    })
    expect(back).toEqual({ horizontalOrigin: 1800, verticalOrigin: 1440 })
  })
})

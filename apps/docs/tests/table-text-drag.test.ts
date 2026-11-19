import { describe, expect, it } from 'vitest'
import { shouldBlockCellSelection } from '../src/renderer/editor/extensions'

/** `a` is where the drag starts, `b` its neighbour; a page-gap widget has no cell */
function cells() {
  const table = document.createElement('table')
  const row = table.insertRow()
  const a = row.insertCell()
  const b = row.insertCell()
  return { a, b }
}

describe('table text drag guard', () => {
  it('blocks a CellSelection while the drag stays inside the starting cell', () => {
    const { a } = cells()
    expect(shouldBlockCellSelection(a, a, a, true)).toBe(true)
  })

  it('also blocks when the drag ends on a page-gap widget (no cell of its own)', () => {
    const { a } = cells()
    expect(shouldBlockCellSelection(a, a, null, true)).toBe(true)
  })

  it('allows a genuine multi-cell drag', () => {
    const { a, b } = cells()
    expect(shouldBlockCellSelection(a, a, b, true)).toBe(false)
  })

  it('leaves command-driven selections alone (no live gesture)', () => {
    const { a, b } = cells()
    expect(shouldBlockCellSelection(null, a, a, true)).toBe(false)
    expect(shouldBlockCellSelection(a, b, b, true)).toBe(false)
  })

  it('leaves a collapsed DOM selection alone', () => {
    const { a } = cells()
    expect(shouldBlockCellSelection(a, a, a, false)).toBe(false)
  })
})

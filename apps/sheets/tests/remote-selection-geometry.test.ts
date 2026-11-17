import { describe, expect, it } from 'vitest'
import { labelGoesInside, membersOnSheet } from '../src/renderer/remote-selection-geometry'

describe('the remote selection overlay', () => {
  it('draws only the members on the sheet that is on screen', () => {
    const members = [
      { id: 1, name: 'A', sheetId: 'sheet-1', range: 'A1:B2' },
      { id: 2, name: 'B', sheetId: 'sheet-2', range: 'C3' },
      { id: 3, name: 'C', sheetId: 'sheet-1' },
    ]
    expect(membersOnSheet(members, 'sheet-1').map((member) => member.id)).toEqual([1])
  })

  it('draws nothing while there is no sheet on screen', () => {
    const members = [{ id: 1, name: 'A', sheetId: 'sheet-1', range: 'A1' }]
    expect(membersOnSheet(members, undefined)).toEqual([])
  })

  it('puts the name tag inside the box when it would cross the top edge', () => {
    expect(labelGoesInside(100, 0)).toBe(false)
    expect(labelGoesInside(100, 95)).toBe(true)
    expect(labelGoesInside(100, 100)).toBe(true)
  })
})

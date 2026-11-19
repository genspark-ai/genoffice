import { describe, expect, it } from 'vitest'

import { mapProtectedRanges } from '../src/renderer/protected-ranges'

const range = (name: string, sqref: string) => ({ name, sqref, hasPassword: false })

describe('mapProtectedRanges', () => {
  it('shifts areas through row inserts and column removals', () => {
    const mapped = mapProtectedRanges(
      [range('Data', 'B3:D6'), range('Cell', 'B2')],
      [
        { kind: 'insert-rows', index: 1, count: 2 },
        { kind: 'remove-cols', index: 0, count: 1 },
      ],
    )
    expect(mapped).toEqual([range('Data', 'A5:C8'), range('Cell', 'A4')])
  })

  it('shrinks partially deleted areas and drops fully deleted ones', () => {
    const mapped = mapProtectedRanges(
      [range('Shrinks', 'A2:A5'), range('Gone', 'A3:A4'), range('Multi', 'A1 A3:A4')],
      [{ kind: 'remove-rows', index: 2, count: 2 }],
    )
    expect(mapped).toEqual([range('Shrinks', 'A2:A3'), range('Multi', 'A1')])
  })

  it('splits a partially moved range into exact runs instead of the envelope', () => {
    // Rows 2-3 (file) move to the tail: file rows 1,2,3 land on screen 4,5,1.
    const mapped = mapProtectedRanges(
      [range('Data', 'A2:A4')],
      [{ kind: 'move-rows', index: 1, count: 2, before: 6 }],
    )
    expect(mapped).toEqual([range('Data', 'A2 A5:A6')])
  })

  it('returns the input unchanged without ops and keeps unparseable parts', () => {
    const untouched = [range('Data', 'B3:D6')]
    expect(mapProtectedRanges(untouched, [])).toEqual(untouched)
    expect(
      mapProtectedRanges([range('Odd', 'NotARef')], [{ kind: 'insert-rows', index: 0, count: 1 }]),
    ).toEqual([range('Odd', 'NotARef')])
  })
})

/**
 * Copying C2:E2 and pasting into C2:C14 must repeat the row down the
 * selection, spilling the source's full width (Excel bulk-fill). The
 * expansion rule decides when the paste target widens to the source extent
 * so upstream's both-multiples tiling can take over.
 */
import { describe, expect, it } from 'vitest'
import { RANGE_TYPE } from '@univerjs/core'
import { anchorTileExpansion } from '../src/renderer/clipboard-anchor-tile'

const rect = (startRow: number, endRow: number, startColumn: number, endColumn: number) => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

describe('anchorTileExpansion', () => {
  it('widens a single-column multi-row target to the source width', () => {
    // C2:C14 (13×1) over a 1×3 source → C2:E14
    expect(anchorTileExpansion(rect(1, 13, 2, 2), 1, 3)).toEqual({
      ...rect(1, 13, 2, 4),
      rangeType: RANGE_TYPE.NORMAL,
    })
  })

  it('heightens a single-row multi-column target to the source height', () => {
    // B2:H2 (1×7... use 6 columns, source 3×2) → multiple on columns
    expect(anchorTileExpansion(rect(1, 1, 1, 6), 3, 2)).toEqual({
      ...rect(1, 3, 1, 6),
      rangeType: RANGE_TYPE.NORMAL,
    })
  })

  it('leaves exact-multiple targets to upstream tiling', () => {
    expect(anchorTileExpansion(rect(1, 13, 2, 4), 1, 3)).toBeNull()
  })

  it('leaves single-cell and same-size targets to upstream anchoring', () => {
    expect(anchorTileExpansion(rect(1, 1, 2, 2), 1, 3)).toBeNull()
    expect(anchorTileExpansion(rect(1, 1, 2, 4), 1, 3)).toBeNull()
  })

  it('leaves non-multiple targets alone (Excel refuses those too)', () => {
    // 5 rows over a 2-row source is not a multiple
    expect(anchorTileExpansion(rect(1, 5, 2, 2), 2, 3)).toBeNull()
  })

  it('ignores whole-row, whole-column and whole-sheet selections', () => {
    expect(
      anchorTileExpansion({ ...rect(1, 13, 0, 25), rangeType: RANGE_TYPE.ROW }, 1, 3),
    ).toBeNull()
    expect(
      anchorTileExpansion({ ...rect(0, 999, 2, 2), rangeType: RANGE_TYPE.COLUMN }, 1, 3),
    ).toBeNull()
    expect(
      anchorTileExpansion({ ...rect(0, 999, 0, 25), rangeType: RANGE_TYPE.ALL }, 1, 3),
    ).toBeNull()
  })

  it('does not widen past the sheet edge', () => {
    const bounds = { rowCount: 20, columnCount: 5 }
    // C2:C14 over a 1×3 source fits (→ column E, index 4)
    expect(anchorTileExpansion(rect(1, 13, 2, 2), 1, 3, bounds)).not.toBeNull()
    // D2:D14 over the same source would need column F, past the last column
    expect(anchorTileExpansion(rect(1, 13, 3, 3), 1, 3, bounds)).toBeNull()
    // B18:G18 over a 3×2 source would need row 20, past the last row
    expect(anchorTileExpansion(rect(17, 17, 1, 6), 3, 2, bounds)).toBeNull()
  })
})

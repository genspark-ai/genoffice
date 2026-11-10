import { describe, expect, it } from 'vitest'

import { preloadInstallRange } from '../src/renderer/univer-sync'

const block = { startRow: 1000, endRow: 1999, startColumn: 0, endColumn: 49 }

describe('preloadInstallRange', () => {
  it('installs nothing for a block without cells, links or banding', () => {
    expect(preloadInstallRange(block, [], [], [], [], [])).toBeNull()
  })

  it('shrinks a column of styled blanks to that column', () => {
    const cells = Array.from({ length: 1000 }, (_unused, i) => ({ row: 1000 + i, column: 3 }))
    expect(preloadInstallRange(block, cells, [], [], [], [])).toEqual({
      startRow: 1000,
      endRow: 1999,
      startColumn: 3,
      endColumn: 3,
    })
  })

  it('spans cells, hyperlinks and the table overlap clipped to the block', () => {
    const tables = [
      { range: { startRow: 0, endRow: 1200, startColumn: 10, endColumn: 12 } },
    ] as unknown as Parameters<typeof preloadInstallRange>[4]
    expect(
      preloadInstallRange(
        block,
        [{ row: 1500, column: 2 }],
        [{ row: 1100, column: 7 }],
        [],
        tables,
        [],
      ),
    ).toEqual({ startRow: 1000, endRow: 1500, startColumn: 2, endColumn: 12 })
  })

  it('widens the box to whole merges so the write never cuts through one', () => {
    const merges = [{ startRow: 1000, endRow: 1000, startColumn: 0, endColumn: 9 }]
    expect(preloadInstallRange(block, [{ row: 1000, column: 0 }], [], merges, [], [])).toEqual({
      startRow: 1000,
      endRow: 1000,
      startColumn: 0,
      endColumn: 9,
    })
  })

  it('ignores banding that lies outside the block', () => {
    const tables = [
      { range: { startRow: 0, endRow: 10, startColumn: 0, endColumn: 5 } },
    ] as unknown as Parameters<typeof preloadInstallRange>[4]
    expect(preloadInstallRange(block, [], [], [], tables, [])).toBeNull()
  })
})

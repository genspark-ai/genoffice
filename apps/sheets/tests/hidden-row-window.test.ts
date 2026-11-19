import { describe, expect, it } from 'vitest'

import { nextHiddenRowWindow } from '../src/renderer/univer-sync'

const window = { startRow: 1000, endRow: 1199, startColumn: 0, endColumn: 11 }
const base = { window, targetVisibleRows: 200, screenRowCount: 373_218, cellCap: 300_000 }

function hiddenEveryRowBut(keepEvery: number): Set<number> {
  const hidden = new Set<number>()
  for (let row = 0; row < 400_000; row += 1) if (row % keepEvery !== 0) hidden.add(row)
  return hidden
}

describe('nextHiddenRowWindow', () => {
  it('doubles the window while the loaded rows show too few visible rows', () => {
    const next = nextHiddenRowWindow({ ...base, hidden: hiddenEveryRowBut(10), loaded: window })
    expect(next).toEqual({ ...window, endRow: 1399 })
  })

  it('stops once the window holds a viewport of visible rows', () => {
    expect(nextHiddenRowWindow({ ...base, hidden: new Set([5000]), loaded: window })).toBeNull()
  })

  it('stops when another load has moved the loaded range away from this window', () => {
    const elsewhere = { startRow: 5000, endRow: 5199, startColumn: 0, endColumn: 11 }
    expect(
      nextHiddenRowWindow({ ...base, hidden: hiddenEveryRowBut(10), loaded: elsewhere }),
    ).toBeNull()
    expect(
      nextHiddenRowWindow({ ...base, hidden: hiddenEveryRowBut(10), loaded: undefined }),
    ).toBeNull()
  })

  it('clamps the doubling to the residency cap and then stops', () => {
    const hidden = hiddenEveryRowBut(1000)
    const wide = { ...window, endRow: 1000 + 20_000 - 1 }
    const next = nextHiddenRowWindow({ ...base, window: wide, hidden, loaded: wide })
    expect(next).toEqual({ ...wide, endRow: 1000 + 25_000 - 1 })
    expect(nextHiddenRowWindow({ ...base, window: next!, hidden, loaded: next! })).toBeNull()
  })

  it('stops at the sheet end', () => {
    const tail = { ...window, startRow: 373_000, endRow: 373_217 }
    expect(
      nextHiddenRowWindow({ ...base, window: tail, hidden: hiddenEveryRowBut(10), loaded: tail }),
    ).toBeNull()
  })
})

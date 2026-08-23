import { describe, expect, it } from 'vitest'
import type { IRange } from '@univerjs/core'

import { revealCellBelowFreeze } from '../src/renderer/univer-sync'

function range(startRow: number, startColumn: number): IRange {
  return { startRow, startColumn, endRow: startRow + 20, endColumn: startColumn + 10 } as IRange
}

describe('revealCellBelowFreeze', () => {
  it('corrects the broken freeze offset until the aim row is visible', async () => {
    // scrollToCell overshoots by 2 rows (custom-height frozen pane)
    const calls: Array<[number, number]> = []
    let scrolledRow = 0
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
        scrolledRow = row
      },
      getVisibleRange: () => range(scrolledRow + 2, 0),
    }
    await revealCellBelowFreeze(sheet, 10, 0)
    expect(calls).toEqual([
      [9, 0],
      [7, 0],
    ])
  })

  it('stops early when the viewport clamps at the frozen pane', async () => {
    // target on the first scrollable row: the aim (row - 1) sits inside the
    // frozen pane, the viewport can never start above the pane edge
    const calls: Array<[number, number]> = []
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
      },
      getVisibleRange: () => range(2, 0),
    }
    await revealCellBelowFreeze(sheet, 2, 0)
    // initial scroll plus one correction, then the no-progress bail — not
    // four futile attempts marching toward row 0
    expect(calls).toEqual([
      [1, 0],
      [0, 0],
    ])
  })

  it('scrolls forward when a clamped correction leaves the target below the viewport', async () => {
    // a very tall frozen pane: the scroll lands the (short) viewport well
    // above the aim, so the aim is visible but the target row is not
    const calls: Array<[number, number]> = []
    let scrolledRow = 0
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
        scrolledRow = row
      },
      getVisibleRange: () => {
        const startRow = Math.max(0, scrolledRow - 8)
        return { startRow, startColumn: 0, endRow: startRow + 6, endColumn: 10 } as IRange
      },
    }
    await revealCellBelowFreeze(sheet, 15, 0)
    // scroll 14 shows rows 6..12 — no overshoot (start 6 <= aim 14), but the
    // target 15 sits below endRow 12: correct forward by 3, landing 9..15
    expect(calls).toEqual([
      [14, 0],
      [17, 0],
    ])
  })

  it('lets a newer reveal supersede an older correction loop', async () => {
    const calls: Array<[number, number]> = []
    let scrolledRow = 0
    const sheet = {
      scrollToCell(row: number, column: number) {
        calls.push([row, column])
        scrolledRow = row
      },
      getVisibleRange: () => range(scrolledRow + 2, 0),
    }
    const older = revealCellBelowFreeze(sheet, 50, 0)
    const newer = revealCellBelowFreeze(sheet, 10, 0)
    await Promise.all([older, newer])
    // the older loop places its initial scroll, then yields at its first
    // generation check; only the newer loop keeps correcting
    expect(calls).toEqual([
      [49, 0],
      [9, 0],
      [7, 0],
    ])
  })
})

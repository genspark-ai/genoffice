import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  MARK_FILL,
  buildMark,
  readCellFill,
  restoreCellFill,
  tintRange,
} from '../src/renderer/redact-tint'

/**
 * The visible half of a mark: painting the tint, and putting the reader's own
 * fill back when the mark is cleared.
 *
 * This is the one part of the feature that writes to the reader's file, so the
 * things worth pinning are all about not losing anything: the order of read and
 * paint, and which fill ends up where.
 */

/// A grid range double recording what was asked of it.
function fakeRange(currentFill: string | null | 'throws' = null) {
  const state = { fill: currentFill === 'throws' ? null : currentFill, painted: [] as string[] }
  const range = {
    getCellStyleData: vi.fn(() => {
      if (currentFill === 'throws') throw new Error('never-touched cell')
      return state.fill === null ? {} : { bg: { rgb: state.fill } }
    }),
    setBackground: vi.fn((color: string) => {
      state.fill = color
      state.painted.push(color)
    }),
    getRange: vi.fn(() => ({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 })),
  }
  return { range, state }
}

const runtime = { univerAPI: { syncExecuteCommand: vi.fn() } }

// the runtime double is shared, so its call count has to be per-test
beforeEach(() => runtime.univerAPI.syncExecuteCommand.mockClear())

describe('painting a mark', () => {
  it('tints the cell and reports the fill it replaced', () => {
    const { range, state } = fakeRange('#FFEE00')
    expect(tintRange(range as never)).toBe('#FFEE00')
    expect(range.setBackground).toHaveBeenCalledWith(MARK_FILL)
    expect(state.painted).toEqual([MARK_FILL])
  })

  it('reads the fill BEFORE painting it', () => {
    // The two orders leave the tinted cell looking identical, so only the
    // reported value tells them apart — and getting it backwards would store
    // the tint as the reader's colour, so clearing could never take it off.
    const { range } = fakeRange('#FFEE00')
    expect(tintRange(range as never)).not.toBe(MARK_FILL)
  })

  it('reports no fill for a cell that never had one', () => {
    const { range } = fakeRange(null)
    expect(tintRange(range as never)).toBeNull()
  })

  it('still marks a cell whose style cannot be read', () => {
    // An untouched cell throws on style resolution; that means "no fill", not
    // "refuse the mark".
    const { range } = fakeRange('throws')
    expect(tintRange(range as never)).toBeNull()
    expect(range.setBackground).toHaveBeenCalledWith(MARK_FILL)
  })

  it('uses a colour the cell value stays readable on', () => {
    // The value is still the reader's, and they have to be able to check it.
    expect(MARK_FILL).toMatch(/^#[0-9A-F]{6}$/i)
    const channel = (offset: number) => parseInt(MARK_FILL.slice(offset, offset + 2), 16)
    const luminance = (0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5)) / 255
    expect(luminance).toBeGreaterThan(0.75)
  })
})

describe('building the mark', () => {
  const BOUNDS = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }

  it('records the displaced fill as part of the mark itself', () => {
    // The failure this closes: a mark assembled without the field paints the
    // tint and can never put the reader's colour back, and nothing looks wrong
    // until the mark is cleared.
    const { range, state } = fakeRange('#FFEE00')
    const mark = buildMark({ getRange: () => range } as never, 'B2:B2', '客户电话', BOUNDS)
    expect(mark.previousFill).toBe('#FFEE00')
    expect(mark.label).toBe('客户电话')
    expect(state.fill).toBe(MARK_FILL)
  })

  it('records a null for a cell that had no fill', () => {
    const { range } = fakeRange(null)
    expect(buildMark({ getRange: () => range } as never, 'B2', 'x', BOUNDS).previousFill).toBeNull()
  })

  it('paints the range it was given, not some other one', () => {
    const asked: string[] = []
    const { range } = fakeRange(null)
    buildMark(
      {
        getRange: (a: string) => {
          asked.push(a)
          return range
        },
      } as never,
      'B2:B500',
      'x',
      { startRow: 1, endRow: 499, startColumn: 1, endColumn: 1 },
    )
    expect(asked).toEqual(['B2:B500'])
  })
})

describe('reading a cell fill', () => {
  it('normalizes the colour the grid reports', () => {
    expect(readCellFill(fakeRange('#ffee00').range as never)).toBe('#FFEE00')
  })

  it('reads a cell with no fill as null', () => {
    expect(readCellFill(fakeRange(null).range as never)).toBeNull()
  })
})

describe('clearing a mark', () => {
  it('puts the reader’s own colour back', () => {
    const { range, state } = fakeRange(null)
    state.fill = MARK_FILL
    restoreCellFill(runtime as never, 'sh1', range as never, '#FFEE00', 'wb')
    expect(state.fill).toBe('#FFEE00')
    expect(runtime.univerAPI.syncExecuteCommand).not.toHaveBeenCalled()
  })

  it('blanks the cell when there was no fill to put back', () => {
    const { range } = fakeRange(null)
    restoreCellFill(runtime as never, 'sh1', range as never, null, 'wb')
    // NOT setBackground(null): that reaches the mutation as a MISSING bg, and
    // a <col style=> fill on the column composes straight back through the
    // cell, so a "cleared" cell would look unchanged.
    expect(runtime.univerAPI.syncExecuteCommand).toHaveBeenCalledWith(
      'sheet.command.set-style',
      expect.objectContaining({
        unitId: 'wb',
        subUnitId: 'sh1',
        style: { type: 'bg', value: { rgb: '' } },
      }),
    )
  })

  it('blanks the cell when the part predates the tint', () => {
    // An absent previousFill means there is nothing recorded to restore, so the
    // only correct action is the same as an explicit null.
    const { range } = fakeRange(null)
    restoreCellFill(runtime as never, 'sh1', range as never, undefined, 'wb')
    expect(runtime.univerAPI.syncExecuteCommand).toHaveBeenCalledTimes(1)
  })

  it('takes the tint off the cell it was on', () => {
    // The whole point: after clearing, the cell must not still look marked.
    const { range, state } = fakeRange(MARK_FILL)
    restoreCellFill(runtime as never, 'sh1', range as never, null, 'wb')
    expect(state.painted).toEqual([])
  })
})

import { CellValueType, type ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { clippedCopyEmfMerge, copyEmfCellDisplay } from '../src/renderer/copy-emf'

describe('copyEmfCellDisplay', () => {
  it('leaves empty cells empty and left-aligned', () => {
    expect(copyEmfCellDisplay(null)).toEqual({ v: '', a: 'l' })
    expect(copyEmfCellDisplay({ v: null } as ICellData)).toEqual({ v: '', a: 'l' })
    expect(copyEmfCellDisplay({ v: '' } as ICellData)).toEqual({ v: '', a: 'l' })
  })

  it('renders booleans as centered TRUE/FALSE', () => {
    expect(copyEmfCellDisplay({ v: true, t: CellValueType.BOOLEAN } as ICellData)).toEqual({
      v: 'TRUE',
      a: 'c',
    })
    expect(copyEmfCellDisplay({ v: 0, t: CellValueType.BOOLEAN } as ICellData)).toEqual({
      v: 'FALSE',
      a: 'c',
    })
  })

  it('renders numbers right-aligned, honoring the number format', () => {
    expect(copyEmfCellDisplay({ v: 42 } as ICellData)).toEqual({ v: '42', a: 'r' })
    const formatted = copyEmfCellDisplay({
      v: 1234.5,
      s: { n: { pattern: '#,##0.00' } },
    } as ICellData)
    expect(formatted.a).toBe('r')
    expect(formatted.v).toBe('1,234.50')
  })

  it('renders strings left-aligned and verbatim', () => {
    expect(copyEmfCellDisplay({ v: 'Name' } as ICellData)).toEqual({ v: 'Name', a: 'l' })
  })
})

describe('clippedCopyEmfMerge', () => {
  const mergeAt = (row: number, column: number, height: number, width: number) => ({
    getRow: () => row,
    getColumn: () => column,
    getHeight: () => height,
    getWidth: () => width,
  })

  it('keeps a merge fully inside the selection', () => {
    expect(clippedCopyEmfMerge(mergeAt(5, 3, 2, 2), 0, 0, 10, 10)).toEqual({
      r: 5,
      c: 3,
      rs: 2,
      cs: 2,
    })
  })

  it('clips a merge that hangs over the selection edge', () => {
    expect(clippedCopyEmfMerge(mergeAt(9, 8, 4, 4), 0, 0, 10, 10)).toEqual({
      r: 9,
      c: 8,
      rs: 1,
      cs: 2,
    })
  })

  it('clips a merge that starts above the selection', () => {
    expect(clippedCopyEmfMerge(mergeAt(2, 0, 5, 2), 4, 0, 10, 10)).toEqual({
      r: 0,
      c: 0,
      rs: 3,
      cs: 2,
    })
  })

  it('drops merges outside the selection and degenerate single cells', () => {
    expect(clippedCopyEmfMerge(mergeAt(20, 0, 2, 2), 0, 0, 10, 10)).toBeNull()
    expect(clippedCopyEmfMerge(mergeAt(1, 1, 1, 1), 0, 0, 10, 10)).toBeNull()
  })
})

/** @vitest-environment jsdom */
describe('emitCopyEmf gating', () => {
  it('asks for the helper once, then skips the payload build off Windows', async () => {
    const { emitCopyEmf } = await import('../src/renderer/copy-emf')
    const { vi } = await import('vitest')
    const available = vi.fn().mockResolvedValue(false)
    const append = vi.fn()
    vi.stubGlobal('window', {
      desktopApi: { isCopyEmfAvailable: available, appendCopyEmf: append },
    })
    const tick = () => new Promise((resolve) => setTimeout(resolve, 0))
    const runtime = {} as Parameters<typeof emitCopyEmf>[0]

    emitCopyEmf(runtime)
    await tick()
    emitCopyEmf(runtime)
    await tick()

    // one ask across copies; the payload build and the IPC never ran
    expect(available).toHaveBeenCalledTimes(1)
    expect(append).not.toHaveBeenCalled()
    vi.unstubAllGlobals()
  })
})

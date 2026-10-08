import { describe, expect, it } from 'vitest'
import { rowFillsSig } from '../src/renderer/editor/pagination-gaps'

describe('rowFillsSig', () => {
  it('is stable across sub-pixel jitter of the same fills', () => {
    const a = rowFillsSig([{ blockTop: 60622.2, row: 7, targetPx: 605.3, extraPx: 48.7 }])
    const b = rowFillsSig([{ blockTop: 60622.4, row: 7, targetPx: 605.1, extraPx: 48.9 }])
    expect(a).toBe(b)
  })

  it('changes when a fill appears, moves rows or is dropped', () => {
    const none = rowFillsSig([])
    const one = rowFillsSig([{ blockTop: 100, row: 7, targetPx: 605 }])
    const other = rowFillsSig([{ blockTop: 100, row: 8, targetPx: 605 }])
    expect(one).not.toBe(none)
    expect(one).not.toBe(other)
    expect(none).toBe('')
  })
})

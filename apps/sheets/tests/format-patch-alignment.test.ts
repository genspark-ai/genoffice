import { describe, expect, it } from 'vitest'

import { applyFormatPatchToRange } from '../src/renderer/univer-sync'

type Range = Parameters<typeof applyFormatPatchToRange>[0]

function makeRange() {
  const calls: { align?: string } = {}
  const range = {
    setHorizontalAlignment: (value: string) => {
      calls.align = value
    },
  } as unknown as Range
  return { range, calls }
}

/// The DSL speaks 'left' | 'center' | 'right'; the facade only accepts
/// 'left' | 'center' | 'normal' (and 'normal' is RIGHT). A raw 'right' used to
/// hit the facade's throwing default branch.
describe('applyFormatPatchToRange horizontal alignment', () => {
  it('passes left and center through untouched', () => {
    for (const align of ['left', 'center'] as const) {
      const { range, calls } = makeRange()
      applyFormatPatchToRange(range, { horizontalAlign: align })
      expect(calls.align).toBe(align)
    }
  })

  it('translates right to the facade vocabulary (normal)', () => {
    const { range, calls } = makeRange()
    applyFormatPatchToRange(range, { horizontalAlign: 'right' })
    expect(calls.align).toBe('normal')
  })

  it('keeps the null fallback to normal', () => {
    const { range, calls } = makeRange()
    applyFormatPatchToRange(range, { horizontalAlign: null })
    expect(calls.align).toBe('normal')
  })
})

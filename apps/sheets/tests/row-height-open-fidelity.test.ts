import { SheetInterceptorService } from '@univerjs/sheets'
import { afterEach, describe, expect, it } from 'vitest'

import {
  countSkeletonTextLines,
  effectiveTextDrawProps,
  excelClipOffsetTop,
  installCellClipAnchorFix,
  shouldAnchorTextTop,
} from '../src/renderer/cell-clip-anchor-fix'
import { installLoadAutoHeightGate, loadAutoHeightSuppression } from '../src/renderer/univer-state'

describe('installLoadAutoHeightGate', () => {
  afterEach(() => {
    loadAutoHeightSuppression.active = false
  })

  // The gate patches the prototype (command handlers resolve the real
  // service); calling the patched method with a stub `this` exercises both
  // the pass-through and the suppressed path against the real Univer code.
  const callGenerate = () => {
    const proto = SheetInterceptorService.prototype as unknown as {
      generateMutationsOfAutoHeight(ctx: unknown): { redos: unknown[]; undos: unknown[] }
    }
    const stub = {
      _autoHeightInterceptors: [
        {
          getMutations: () => ({ redos: [{ id: 'measure' }], undos: [{ id: 'unmeasure' }] }),
        },
      ],
    }
    return proto.generateMutationsOfAutoHeight.call(stub, {})
  }

  it('passes auto-height mutations through when no file content is installing', () => {
    installLoadAutoHeightGate()
    expect(callGenerate().redos).toEqual([{ id: 'measure' }])
  })

  it('yields no mutations while a file patch is installing', () => {
    installLoadAutoHeightGate()
    loadAutoHeightSuppression.active = true
    const result = callGenerate()
    expect(result.redos).toEqual([])
    expect(result.undos).toEqual([])
  })
})

describe('excelClipOffsetTop', () => {
  it('anchors an overflowing multi-line block to the cell top', () => {
    // Bottom/middle alignment with text taller than the cell yields a
    // negative offset (last line visible) — Excel shows the first line.
    expect(excelClipOffsetTop(-42, 3)).toBe(0)
  })

  it('keeps a near-fit single line on its native alignment', () => {
    // Excel clips a bottom-aligned single line at its top when the row is
    // slightly short — the negative offset is the correct rendering.
    expect(excelClipOffsetTop(-3, 1)).toBe(-3)
  })

  it('keeps non-overflowing alignment offsets untouched', () => {
    expect(excelClipOffsetTop(0, 1)).toBe(0)
    expect(excelClipOffsetTop(7.5, 2)).toBe(7.5)
  })

  it('installs idempotently on the Documents prototype', async () => {
    const { Documents } = await import('@univerjs/engine-render')
    installCellClipAnchorFix()
    const patched = (Documents.prototype as unknown as Record<string, unknown>)._verticalHandler
    installCellClipAnchorFix()
    expect((Documents.prototype as unknown as Record<string, unknown>)._verticalHandler).toBe(
      patched,
    )
    const call = patched as (h: number, pt: number, pb: number, v: number) => number
    const docStub = (lineCount: number) => ({
      height: 20,
      getSkeleton: () => ({
        getSkeletonData: () => ({
          pages: [
            {
              sections: [{ columns: [{ lines: Array.from({ length: lineCount }, () => ({})) }] }],
            },
          ],
        }),
      }),
    })
    // Bottom-aligned overflow through the real patched method: height 20,
    // content 60 → raw offset −40 → clamped to the cell top for 3 lines...
    expect(call.call(docStub(3), 60, 0, 0, 3)).toBe(0)
    // ...but a single rich line keeps its native (negative) bottom offset.
    expect(call.call(docStub(1), 24, 0, 0, 3)).toBe(-4)
  })
})

describe('countSkeletonTextLines', () => {
  it('counts text lines across pages/sections/columns and skips BLOCK lines', () => {
    const data = {
      pages: [
        { sections: [{ columns: [{ lines: [{ type: 0 }, { type: 1 }] }] }] },
        { sections: [{ columns: [{ lines: [{ type: 0 }] }] }] },
      ],
    }
    expect(countSkeletonTextLines(data)).toBe(2)
    expect(countSkeletonTextLines(undefined)).toBe(0)
  })
})

describe('shouldAnchorTextTop', () => {
  it('re-anchors an overflowing multi-line plain-text block to the top', () => {
    expect(shouldAnchorTextTop(2, 42, 21)).toBe(true)
  })

  it('keeps single lines and fitting blocks on their native alignment', () => {
    // A slightly-short row clips a bottom-aligned single line at its top in
    // Excel too, so one-liners never re-anchor.
    expect(shouldAnchorTextTop(1, 24, 21)).toBe(false)
    expect(shouldAnchorTextTop(2, 40, 46)).toBe(false)
  })

  it('flips the vAlign handed to Text.drawWith only for overflowing blocks', () => {
    const props = {
      text: 'a b',
      fontStyle: '12px Arial',
      warp: true,
      vAlign: 3, // BOTTOM
      width: 40,
      height: 21,
    }
    expect(effectiveTextDrawProps(props as never, 2, 42)).toMatchObject({ vAlign: 1 })
    // Fitting block: same object back, native alignment untouched.
    expect(effectiveTextDrawProps(props as never, 2, 18)).toBe(props)
  })
})

import { describe, expect, it } from 'vitest'
import { carryStreamedSamples, fillLineBoxes } from '../src/renderer/pagination-lines'
import type { BlockBox } from '../src/renderer/pagination-types'

const rectOf = (top: number, height: number, width = 100) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: width,
    width,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

/** an el-bearing block with the given geometry and optional prior samples */
function blockOf(el: HTMLElement, top: number, height: number, prev?: Partial<BlockBox>): BlockBox {
  el.getBoundingClientRect = () => rectOf(top, height)
  return { top, height, el, ...(prev ?? {}) } as BlockBox
}

describe('carryStreamedSamples', () => {
  it('carries line/row samples across the element-identity prefix', () => {
    const els = [document.createElement('p'), document.createElement('p')]
    const prev = [
      blockOf(els[0], 0, 100, { lineBoxes: [{ offsetInBlock: 0, height: 12 }] as never }),
      blockOf(els[1], 100, 80, { lineLeadPx: 3 }),
    ]
    const next = [
      blockOf(els[0], 0, 100),
      blockOf(els[1], 100, 80),
      blockOf(document.createElement('p'), 180, 50),
    ]
    carryStreamedSamples(next, prev, { pending: true, dirty: 2, lastPassChildCount: 2 })
    expect(next[0].lineBoxes).toEqual(prev[0].lineBoxes)
    expect(next[0].lineLeadPx).toBeUndefined()
    expect(next[1].lineLeadPx).toBe(3)
    expect(next[1].lineBoxes).toBeUndefined()
    // the appended tail has no previous sample to carry
    expect(next[2].lineBoxes).toBeUndefined()
  })

  it('stops at the element-identity frontier (an inserted block shifts the tail)', () => {
    const head = document.createElement('p')
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const prev = [blockOf(head, 0, 100, { lineBoxes: carried })]
    const inserted = document.createElement('p')
    const after = document.createElement('p')
    const next = [blockOf(head, 0, 100), blockOf(inserted, 100, 10), blockOf(after, 110, 100)]
    carryStreamedSamples(next, prev, { pending: true, dirty: 1, lastPassChildCount: 1 })
    expect(next[0].lineBoxes).toEqual(carried)
    // the arrays no longer align past the insertion: nothing carries
    expect(next[1].lineBoxes).toBeUndefined()
    expect(next[2].lineBoxes).toBeUndefined()
  })

  it('a changed block keeps its identity slot while later unchanged blocks still carry', () => {
    const els = [document.createElement('p'), document.createElement('p')]
    const carriedA = [{ offsetInBlock: 0, height: 12 }] as never
    const carriedB = [{ offsetInBlock: 0, height: 15 }] as never
    const prev = [
      blockOf(els[0], 0, 100, { lineBoxes: carriedA }),
      blockOf(els[1], 100, 80, { lineBoxes: carriedB }),
    ]
    // the first block grew (its own line layout changed), the second is untouched
    const next = [blockOf(els[0], 0, 140), blockOf(els[1], 140, 80)]
    carryStreamedSamples(next, prev, { pending: true, dirty: 0, lastPassChildCount: 0 })
    expect(next[0].lineBoxes).toBeUndefined()
    expect(next[1].lineBoxes).toEqual(carriedB)
  })

  it('an edit below the append frontier disables the carry wholesale', () => {
    const el = document.createElement('p')
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const prev = [blockOf(el, 0, 100, { lineBoxes: carried })]
    const next = [blockOf(el, 0, 100)]
    carryStreamedSamples(next, prev, { pending: true, dirty: 0, lastPassChildCount: 2 })
    expect(next[0].lineBoxes).toBeUndefined()
  })

  it('a webfont-load pass (dirty = null) disables the carry', () => {
    const el = document.createElement('p')
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const prev = [blockOf(el, 0, 100, { lineBoxes: carried })]
    const next = [blockOf(el, 0, 100)]
    carryStreamedSamples(next, prev, { pending: true, dirty: null, lastPassChildCount: 1 })
    expect(next[0].lineBoxes).toBeUndefined()
  })

  it('a settled pass (pending = false) disables the carry', () => {
    const el = document.createElement('p')
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const prev = [blockOf(el, 0, 100, { lineBoxes: carried })]
    const next = [blockOf(el, 0, 100)]
    carryStreamedSamples(next, prev, { pending: false, dirty: 1, lastPassChildCount: 1 })
    expect(next[0].lineBoxes).toBeUndefined()
  })

  it('a block whose width changed re-samples (wrap points moved at equal height)', () => {
    const el = document.createElement('p')
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const prev = [blockOf(el, 0, 100, { lineBoxes: carried, widthPx: 400 })]
    const next = [blockOf(el, 0, 100)]
    next[0].widthPx = 300
    carryStreamedSamples(next, prev, { pending: true, dirty: 1, lastPassChildCount: 1 })
    expect(next[0].lineBoxes).toBeUndefined()
  })

  it('carried samples let fillLineBoxes skip the block', () => {
    const el = document.createElement('p')
    el.textContent = 'hello world of pagination'
    const carried = [{ offsetInBlock: 0, height: 12 }] as never
    const next = [blockOf(el, 0, 100)]
    next[0].lineBoxes = carried
    // jsdom lays out nothing: without the carried sample this block would gain
    // none either, but the skip is observable in that the sample survives as-is
    fillLineBoxes(next, [{ contentHeight: 80, forceBreak: false }], 1)
    expect(next[0].lineBoxes).toEqual(carried)
  })
})

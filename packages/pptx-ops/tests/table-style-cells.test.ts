/**
 * setTableStyle "cells" validation (packages/pptx-ops/src/ops/table-ops.ts
 * resolveStyleCells). The field used to pass straight through to the engine,
 * where a non-array made [...cells] throw a TypeError, a null entry threw on
 * destructuring, and an oversized array reached a Math.max spread. Every
 * malformed shape must now come back as a guided validation error.
 */
import { describe, it, expect, beforeAll } from 'vitest'
import { createBlankPptx, openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let tableId: string

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const r = runTxn(opened, {
    ops: [
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 2,
        cols: 2,
        offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
      },
    ],
  })
  expect(r.applied).toBe(true)
  const table = opened.deck.slides[0]!.elements.find((el) => el.type === 'table')
  tableId = table!.id
})

const run = (cells: unknown) =>
  runTxn(opened, {
    ops: [
      {
        op: 'setTableStyle',
        target: { slide: 0, el: tableId },
        firstRow: true,
        cells,
      },
    ],
  })

const expectRejected = (cells: unknown, pattern: RegExp) => {
  const r = run(cells)
  expect(r.applied).toBe(false)
  expect(r.failures).toHaveLength(1)
  expect(r.failures![0]!.error).toMatch(pattern)
}

describe('setTableStyle cells validation', () => {
  it('accepts a normal list of row/column pairs', () => {
    const r = run([
      { row: 0, col: 0 },
      { row: 1, col: 1 },
    ])
    expect(r.applied).toBe(true)
  })

  it('accepts an empty list', () => {
    expect(run([]).applied).toBe(true)
  })

  it('rejects a non-array cells value', () => {
    expectRejected(7, /cells must be an array/)
    expectRejected('0,0', /cells must be an array/)
    expectRejected({ row: 0, col: 0 }, /cells must be an array/)
  })

  it('rejects null and primitive entries', () => {
    expectRejected([null], /cells\[0\] must be a \{ row, col \} object/)
    expectRejected([{ row: 0, col: 0 }, null], /cells\[1\] must be a \{ row, col \} object/)
    expectRejected([3], /cells\[0\] must be a \{ row, col \} object/)
  })

  it('rejects non-integer row/column', () => {
    expectRejected([{ row: 0.5, col: 1 }], /must be integers/)
    expectRejected([{ row: 0, col: 1.5 }], /must be integers/)
    expectRejected([{ row: '0', col: 1 }], /must be integers/)
    expectRejected([{ row: NaN, col: 1 }], /must be integers/)
    expectRejected([{ col: 1 }], /must be integers/)
    expectRejected([{ row: 0 }], /must be integers/)
  })

  it('rejects negative row/column', () => {
    expectRejected([{ row: -1, col: 0 }], /must be >= 0/)
    expectRejected([{ row: 0, col: -2 }], /must be >= 0/)
  })

  it('rejects an oversized list before it reaches the engine', () => {
    const huge = Array.from({ length: 200_000 }, (_, i) => ({ row: 0, col: i % 50 }))
    expectRejected(huge, /cells takes at most 4096 entries, got 200000/)
  })
})

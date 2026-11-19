import { describe, expect, it } from 'vitest'
import { BlockIndex, CapacityWindows, SortedYs } from '../src/renderer/pagination-index'
import type { BlockBox } from '../src/renderer/pagination-types'

const box = (top: number, height: number, extra: Partial<BlockBox> = {}): BlockBox => ({
  top,
  height,
  ...extra,
})

describe('SortedYs', () => {
  const ys = new SortedYs([900, 0, 1800, 900, 2700])
  it('dedups and answers range/near queries', () => {
    expect(ys.length).toBe(4)
    expect(ys.hasBetween(850, 950)).toBe(true)
    expect(ys.hasBetween(900, 950)).toBe(false)
    expect(ys.hasBetween(850, 900)).toBe(false)
    expect(ys.hasBetween(1000, 1700)).toBe(false)
    expect(ys.hasNear(900.3, 0.5)).toBe(true)
    expect(ys.hasNear(901, 0.5)).toBe(false)
    // strict, like the Math.abs(...) < 0.5 scan it replaces
    expect(ys.hasNear(900.5, 0.5)).toBe(false)
    expect(ys.hasNear(899.5, 0.5)).toBe(false)
    expect(ys.hasNear(900.49, 0.5)).toBe(true)
    expect(ys.floor(1000)).toBe(900)
    expect(ys.floor(-1)).toBeUndefined()
  })
  it('matches the linear scans it replaces', () => {
    const vals = [3, 7.5, 12, 12, 40, 41.2]
    const s = new SortedYs(vals)
    for (let a = 0; a < 45; a += 0.5) {
      for (let b = a; b < 45; b += 1.3) {
        expect(s.hasBetween(a, b)).toBe(vals.some((y) => a < y && y < b))
      }
      expect(s.hasNear(a, 0.5)).toBe(vals.some((y) => Math.abs(a - y) < 0.5))
    }
  })
})

describe('BlockIndex', () => {
  // document order differs from top order: a lifted table (index 2) starts above block 1
  const blocks = [
    box(0, 100),
    box(100, 50),
    box(90, 200, { liftPx: 10 }),
    box(300, 0, { emptyPara: true }),
    box(300, 40),
    box(340, 600),
  ]
  const index = new BlockIndex(blocks)

  it('returns the first block in document order at a top', () => {
    expect(index.firstAtTop(300)).toBe(3)
    expect(index.firstAtTop(300.4)).toBe(3)
    expect(index.firstAtTop(300, (b) => !b.emptyPara)).toBe(4)
    expect(index.firstAtTop(90)).toBe(2)
    expect(index.firstAtTop(200)).toBe(-1)
  })

  it('returns the first containing block in document order', () => {
    // 120 lies inside block 1 (100..150) and block 2 (90..290): document order wins
    expect(index.containing(120)).toBe(blocks[1])
    expect(index.containing(120, (b) => !!b.liftPx)).toBe(blocks[2])
    expect(index.containing(500)).toBe(blocks[5])
    // a block's own top and the last half pixel are not "inside"
    expect(index.containing(340)).toBe(undefined)
    expect(index.containing(939.8)).toBe(undefined)
    expect(index.containing(939.4)).toBe(blocks[5])
  })

  it('agrees with the linear predicates on random layouts', () => {
    const rnd = (n: number) => Math.floor(Math.random() * n)
    for (let round = 0; round < 20; round++) {
      const bs: BlockBox[] = []
      let y = 0
      for (let i = 0; i < 60; i++) {
        const h = rnd(3) === 0 ? 0 : rnd(120)
        bs.push(box(y - (rnd(4) === 0 ? rnd(30) : 0), h))
        y += h
      }
      const idx = new BlockIndex(bs)
      for (let q = -5; q < y + 5; q += 3.1) {
        expect(idx.firstAtTop(q)).toBe(bs.findIndex((b) => Math.abs(b.top - q) < 0.5))
        expect(idx.containing(q)).toBe(bs.find((b) => b.top < q && q < b.top + b.height - 0.5))
      }
    }
  })
})

describe('CapacityWindows', () => {
  it('applies every window containing a Y, duplicate starts included', () => {
    // a titlePg first page and its first column share a start; a balanced
    // region's empty trailing column shares its start with the next region
    const w = new CapacityWindows([
      { start: 0, end: 900, cap: 700 },
      { start: 0, end: 300, cap: 300 },
      { start: 300, end: 300, cap: 300 },
      { start: 300, end: 900, cap: 500 },
      { start: 900, end: 1800, cap: 900 },
    ])
    expect(w.capAt(100, 1000)).toBe(300)
    expect(w.capAt(299.6, 1000)).toBe(500)
    expect(w.capAt(600, 1000)).toBe(500)
    expect(w.capAt(1200, 1000)).toBe(900)
    expect(w.capAt(2000, 1000)).toBe(1000)
    // the half-pixel tolerance is inclusive on the start side (start - 0.5 <= y)
    expect(w.capAt(899.5, 1000)).toBe(900)
    expect(w.capAt(899.4, 1000)).toBe(500)
    expect(w.capAt(-0.5, 1000)).toBe(300)
  })
  it('matches the linear scan on random windows', () => {
    const rnd = (n: number) => Math.floor(Math.random() * n)
    for (let round = 0; round < 30; round++) {
      const ws = Array.from({ length: 12 }, () => {
        const start = rnd(20) * 50
        return { start, end: start + rnd(4) * 100, cap: 100 + rnd(900) }
      })
      const w = new CapacityWindows(ws)
      for (let y = -10.5; y < 1400; y += 3.5) {
        let cap = 1000
        for (const x of ws) if (x.start - 0.5 <= y && y < x.end - 0.5) cap = Math.min(cap, x.cap)
        expect(w.capAt(y, 1000)).toBe(cap)
      }
    }
  })
})

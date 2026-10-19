/**
 * The graph view's layout: deterministic, inside the viewport, and shaped —
 * linked nodes end up nearer each other than strangers do.
 */
import { describe, expect, it } from 'vitest'

import { layoutGraph } from '../src/renderer/components/graph-layout'

const dist = (a: { x: number; y: number }, b: { x: number; y: number }): number =>
  Math.hypot(a.x - b.x, a.y - b.y)

describe('layoutGraph', () => {
  it('places every node inside the viewport, never at NaN', () => {
    const nodes = Array.from({ length: 12 }, (_, i) => ({ path: `n${i}` }))
    const edges = [
      { from: 'n0', to: 'n1' },
      { from: 'n1', to: 'n2' },
      { from: 'n0', to: 'n2' },
    ]
    const positions = layoutGraph(nodes, edges, 900, 600)
    expect(positions.size).toBe(12)
    for (const p of positions.values()) {
      expect(Number.isFinite(p.x)).toBe(true)
      expect(Number.isFinite(p.y)).toBe(true)
      expect(p.x).toBeGreaterThanOrEqual(0)
      expect(p.x).toBeLessThanOrEqual(900)
      expect(p.y).toBeGreaterThanOrEqual(0)
      expect(p.y).toBeLessThanOrEqual(600)
    }
  })

  it('a linked pair settles nearer than an unlinked stranger', () => {
    const nodes = [{ path: 'a' }, { path: 'b' }, { path: 'far' }]
    const linked = layoutGraph(nodes, [{ from: 'a', to: 'b' }], 800, 600)
    const stranger = dist(linked.get('a')!, linked.get('far')!)
    expect(dist(linked.get('a')!, linked.get('b')!)).toBeLessThan(stranger)
  })

  it('the same input draws the same picture', () => {
    const nodes = [{ path: 'a' }, { path: 'b' }, { path: 'c' }]
    const edges = [
      { from: 'a', to: 'b' },
      { from: 'b', to: 'c' },
    ]
    const first = layoutGraph(nodes, edges, 700, 500)
    const second = layoutGraph(nodes, edges, 700, 500)
    for (const [path, p] of first) {
      expect(second.get(path)).toEqual(p)
    }
  })

  it('empty input answers an empty layout', () => {
    expect(layoutGraph([], [], 500, 500).size).toBe(0)
    const nodes = [{ path: 'a' }]
    expect(layoutGraph(nodes, [], 0, 500).size).toBe(0)
  })
})

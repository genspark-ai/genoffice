import { describe, expect, it } from 'vitest'
import { layoutHierTree, parseHierConstraints, type HierTreeNode } from '../src/dgm-hier'

const n = (text: string, children: HierTreeNode[] = []): HierTreeNode => ({
  texts: [text],
  children,
})

const CONS = parseHierConstraints(undefined)
const FRAME_CX = 8000000
const FRAME_CY = 6000000
const WIDE = 130_000
const TIMEOUT = 180_000

describe('layoutHierTree fan-out and depth', () => {
  it('lays out a normal small org chart unchanged', () => {
    const tree = n('A', [n('B1', [n('C1'), n('C2')]), n('B2', [n('C3')])])
    const g = layoutHierTree([tree], CONS, FRAME_CX, FRAME_CY)!
    expect(g.boxes.map((b) => b.node.texts[0]).sort()).toEqual(['A', 'B1', 'B2', 'C1', 'C2', 'C3'])
    expect(g.boxW).toBeGreaterThan(0)
  })

  it('lays out several roots and picks the deepest row count', () => {
    const roots = [n('r0', [n('a', [n('b')])]), n('r1', [n('c')]), n('r2')]
    const g = layoutHierTree(roots, CONS, FRAME_CX, FRAME_CY)!
    expect(g.boxes).toHaveLength(6)
  })

  it(
    'lays out a root fan-out wider than the engine argument limit',
    () => {
      const kids = Array.from({ length: WIDE }, (_, i) => n(`k${i}`))
      const g = layoutHierTree([n('root', kids)], CONS, FRAME_CX, FRAME_CY)!
      expect(g.boxes).toHaveLength(WIDE + 1)
      expect(g.lines.length).toBeGreaterThan(0)
    },
    TIMEOUT,
  )

  it('lays out a deep balanced hierarchy completely', () => {
    const DEPTH = 8
    const FANOUT = 3
    const build = (level: number): HierTreeNode =>
      level >= DEPTH
        ? n(`L${level}`)
        : n(
            `L${level}`,
            Array.from({ length: FANOUT }, () => build(level + 1)),
          )
    const total = (FANOUT ** (DEPTH + 1) - 1) / (FANOUT - 1)
    const g = layoutHierTree([build(0)], CONS, FRAME_CX, FRAME_CY)!
    expect(g.boxes).toHaveLength(total)
    expect(g.boxW).toBeGreaterThan(0)
  })

  it('lays out a deep chain completely', () => {
    const DEPTH = 400
    let ref = n(`d${DEPTH}`)
    for (let i = DEPTH - 1; i >= 0; i--) ref = n(`d${i}`, [ref])
    const g = layoutHierTree([ref], CONS, FRAME_CX, FRAME_CY)!
    expect(g.boxes).toHaveLength(DEPTH + 1)
    expect(g.lines.length).toBeGreaterThan(0)
  })
})

/**
 * The desktop selection layer's pure rules (src/renderer/src/desktop-selection.ts):
 * the marquee's rectangle geometry and modifier merge semantics, the icon
 * click modifiers, the Shift range over the grid's visible-geometry order,
 * the arrow keys' geometry walk, and the Ask AI path harvest. No bridge, no DOM.
 */
import { describe, expect, it } from 'vitest'
import {
  EMPTY_SELECTION,
  askPathsOf,
  geometryOrderOf,
  marqueeHits,
  marqueeRect,
  mergeMarqueeSelection,
  moveGridFocus,
  rangeSelection,
  rectsIntersect,
  toggleSelectionMember,
  type GridCell,
  type SelectionTarget,
} from '../src/renderer/src/desktop-selection'
import type { DesktopItem } from '../src/renderer/src/desktop-items'

const target = (id: string, x: number, y: number, w = 96, h = 104): SelectionTarget => ({
  id,
  x,
  y,
  w,
  h,
})

describe('marqueeRect', () => {
  it('normalizes either drag direction into a top-left origin', () => {
    expect(marqueeRect(100, 200, 300, 400)).toEqual({ x: 100, y: 200, w: 200, h: 200 })
    expect(marqueeRect(300, 400, 100, 200)).toEqual({ x: 100, y: 200, w: 200, h: 200 })
    // the degenerate press-and-hold: a zero-area band at the press point
    expect(marqueeRect(50, 60, 50, 60)).toEqual({ x: 50, y: 60, w: 0, h: 0 })
  })
})

describe('rectsIntersect', () => {
  it('needs shared positive area on both axes', () => {
    expect(rectsIntersect({ x: 0, y: 0, w: 10, h: 10 }, { x: 5, y: 5, w: 10, h: 10 })).toBe(true)
    expect(rectsIntersect({ x: 0, y: 0, w: 10, h: 10 }, { x: 20, y: 20, w: 5, h: 5 })).toBe(false)
    // containment counts
    expect(rectsIntersect({ x: 0, y: 0, w: 100, h: 100 }, { x: 10, y: 10, w: 5, h: 5 })).toBe(true)
  })

  it('edges that merely touch do not intersect', () => {
    expect(rectsIntersect({ x: 0, y: 0, w: 10, h: 10 }, { x: 10, y: 0, w: 10, h: 10 })).toBe(false)
    expect(rectsIntersect({ x: 0, y: 0, w: 10, h: 10 }, { x: 0, y: 10, w: 10, h: 10 })).toBe(false)
    // a zero-area band claims nothing, whatever it sits on
    expect(rectsIntersect({ x: 5, y: 5, w: 0, h: 0 }, { x: 0, y: 0, w: 10, h: 10 })).toBe(false)
  })
})

describe('marqueeHits', () => {
  const icons = [
    target('a', 12, 12),
    target('b', 108, 12),
    target('c', 12, 116),
    target('d', 204, 12),
  ]

  it('returns the intersecting ids in the order given', () => {
    expect(marqueeHits({ x: 100, y: 0, w: 60, h: 100 }, icons)).toEqual(['a', 'b'])
    expect(marqueeHits({ x: 0, y: 0, w: 300, h: 300 }, icons)).toEqual(['a', 'b', 'c', 'd'])
  })

  it('an empty band hits nothing, a band between icons hits nothing', () => {
    expect(marqueeHits({ x: 5, y: 5, w: 0, h: 0 }, icons)).toEqual([])
    // the gutter between column 0 (ends 108) and column 1 (starts 108) is zero-wide,
    // but a band clear of every cell on one axis misses everything
    expect(marqueeHits({ x: 300, y: 300, w: 40, h: 40 }, icons)).toEqual([])
  })
})

describe('mergeMarqueeSelection', () => {
  const base = new Set(['a', 'b'])

  it('replace: the band is the whole selection', () => {
    expect(mergeMarqueeSelection(base, ['c', 'd'], 'replace')).toEqual(new Set(['c', 'd']))
    expect(mergeMarqueeSelection(base, [], 'replace')).toEqual(new Set())
  })

  it('add (Shift): the hits union onto what was selected', () => {
    expect(mergeMarqueeSelection(base, ['b', 'c'], 'add')).toEqual(new Set(['a', 'b', 'c']))
  })

  it('toggle (Cmd/Ctrl): hits flip — the already-selected leave, the rest join', () => {
    expect(mergeMarqueeSelection(base, ['b', 'c'], 'toggle')).toEqual(new Set(['a', 'c']))
    // the band shrank and no longer covers c: recomputed from the same base, c leaves
    expect(mergeMarqueeSelection(base, [], 'toggle')).toEqual(new Set(['a', 'b']))
  })

  it('never mutates the base set', () => {
    mergeMarqueeSelection(base, ['b'], 'toggle')
    expect([...base]).toEqual(['a', 'b'])
  })
})

describe('toggleSelectionMember', () => {
  it('flips one id and keeps the rest', () => {
    expect(toggleSelectionMember(new Set(['a', 'b']), 'b')).toEqual(new Set(['a']))
    expect(toggleSelectionMember(new Set(['a']), 'b')).toEqual(new Set(['a', 'b']))
    expect(toggleSelectionMember(EMPTY_SELECTION, 'a')).toEqual(new Set(['a']))
  })
})

describe('rangeSelection', () => {
  const order = ['a', 'b', 'c', 'd', 'e']

  it('selects the flow-order span between the anchor and the target, either way', () => {
    expect([...rangeSelection(order, 'a', 'c')]).toEqual(['a', 'b', 'c'])
    expect([...rangeSelection(order, 'c', 'a')]).toEqual(['a', 'b', 'c'])
    expect([...rangeSelection(order, 'b', 'b')]).toEqual(['b'])
  })

  it('a missing anchor falls back to the click alone', () => {
    expect([...rangeSelection(order, null, 'c')]).toEqual(['c'])
    // an anchor that left the surface (an icon removed mid-flight) too
    expect([...rangeSelection(order, 'gone', 'c')]).toEqual(['c'])
  })

  it('a target outside the order answers just the target', () => {
    expect([...rangeSelection(order, 'a', 'gone')]).toEqual(['gone'])
  })
})

describe('geometryOrderOf — the visible-geometry reading order', () => {
  // a custom placement: the provider order (a,b,c,d,e) is scattered across
  // the grid, and the order must follow the cells instead
  const scattered: GridCell[] = [
    { id: 'a', col: 0, row: 0 },
    { id: 'b', col: 0, row: 2 },
    { id: 'c', col: 1, row: 0 },
    { id: 'd', col: 2, row: 1 },
    { id: 'e', col: 0, row: 1 },
  ]

  it('sorts rows first, columns within a row', () => {
    expect(geometryOrderOf(scattered)).toEqual(['a', 'c', 'e', 'd', 'b'])
  })

  it('a range over the geometry order spans the visually contiguous run', () => {
    // anchor a (0,0) to target e (0,1): row 0 in between plus e itself —
    // the provider-order span a..e would have claimed all five
    expect([...rangeSelection(geometryOrderOf(scattered), 'a', 'e')]).toEqual(['a', 'c', 'e'])
  })

  it('an empty cell set answers an empty order', () => {
    expect(geometryOrderOf([])).toEqual([])
  })
})

describe('moveGridFocus', () => {
  // a column-major flow: three columns of two, like the desktop's fill order
  const cells: GridCell[] = [
    { id: 'a', col: 0, row: 0 },
    { id: 'b', col: 0, row: 1 },
    { id: 'c', col: 1, row: 0 },
    { id: 'd', col: 1, row: 1 },
    { id: 'e', col: 2, row: 0 },
  ]

  it('walks the geometry: down/up step rows, right/left step columns', () => {
    expect(moveGridFocus(cells, 'a', 'down')).toBe('b')
    expect(moveGridFocus(cells, 'b', 'up')).toBe('a')
    expect(moveGridFocus(cells, 'a', 'right')).toBe('c')
    expect(moveGridFocus(cells, 'c', 'left')).toBe('a')
    expect(moveGridFocus(cells, 'e', 'left')).toBe('c')
  })

  it('an edge in the stepped direction keeps the focus (null answer)', () => {
    expect(moveGridFocus(cells, 'b', 'down')).toBeNull()
    expect(moveGridFocus(cells, 'a', 'up')).toBeNull()
    expect(moveGridFocus(cells, 'e', 'right')).toBeNull()
  })

  it('a ragged grid falls over to the nearest cell of the next column/row', () => {
    // e (2,0) down: no row below in its column, so the nearest lower cell overall wins
    expect(moveGridFocus(cells, 'e', 'down')).toBe('d')
    // d (1,1) down: no row below in any column — null; right from d: column 2
    expect(moveGridFocus(cells, 'd', 'down')).toBeNull()
    expect(moveGridFocus(cells, 'd', 'right')).toBe('e')
    // b (0,1) right: nearest of column 1 by row distance is d
    expect(moveGridFocus(cells, 'b', 'right')).toBe('d')
    // c (1,0) down: nearest of row 1 by column distance is d
    expect(moveGridFocus(cells, 'c', 'down')).toBe('d')
  })

  it('with nothing focused, down/right enter from the top-left, up/left from the bottom', () => {
    expect(moveGridFocus(cells, null, 'down')).toBe('a')
    expect(moveGridFocus(cells, null, 'right')).toBe('a')
    expect(moveGridFocus(cells, null, 'up')).toBe('d')
    expect(moveGridFocus(cells, null, 'left')).toBe('e')
    expect(moveGridFocus([], null, 'down')).toBeNull()
  })

  it('a focus that left the grid re-enters from the corner', () => {
    expect(moveGridFocus(cells, 'gone', 'down')).toBe('a')
  })

  it('follows saved positions, not the list order', () => {
    // the same ids, dragged into a different shape: b pinned to (2,2)
    const rearranged: GridCell[] = [
      { id: 'a', col: 0, row: 0 },
      { id: 'c', col: 1, row: 0 },
      { id: 'b', col: 2, row: 2 },
    ]
    expect(moveGridFocus(rearranged, 'a', 'right')).toBe('c')
    expect(moveGridFocus(rearranged, 'c', 'right')).toBe('b')
    expect(moveGridFocus(rearranged, 'b', 'down')).toBeNull()
  })
})

describe('askPathsOf', () => {
  const item = (over: Partial<DesktopItem>): DesktopItem => ({
    id: 'x',
    kind: 'file',
    title: 'x',
    tile: null,
    ext: 'docx',
    path: null,
    alias: false,
    draggable: true,
    open: () => {},
    ...over,
  })

  it('harvests the real paths of files and folders, in the given order', () => {
    expect(
      askPathsOf([
        item({ id: 'app', kind: 'app', path: null }),
        item({ id: 'f1', path: '/docs/a.docx' }),
        item({ id: 'folder', kind: 'folder', path: '/docs/Work' }),
        item({ id: 'f2', kind: 'folder', path: null }),
      ]),
    ).toEqual(['/docs/a.docx', '/docs/Work'])
  })

  it('skips app tiles and pathless items, drops duplicate paths', () => {
    expect(
      askPathsOf([
        item({ id: 'ask', kind: 'app', tile: 'ask', path: null }),
        item({ id: 'a', path: '/docs/a.docx' }),
        item({ id: 'a-alias', path: '/docs/a.docx' }),
      ]),
    ).toEqual(['/docs/a.docx'])
    expect(askPathsOf([item({ id: 'app', kind: 'app' })])).toEqual([])
  })
})

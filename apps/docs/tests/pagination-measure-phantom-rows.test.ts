/**
 * A cut table carries phantom rows on the canvas (page-gap row, repeated
 * header clones) that the measuring state collapses and the preview clone
 * removes. When a collapsed row answers a zero rect while the table still
 * keeps height for it, every block after the table measured one line below
 * where the clone paints it (008: two cut tables, ~19 px each, the empty
 * paragraph after each one started the drift). The residue between the real
 * rows the phantoms separate is subtracted like any in-block gap.
 */
import { describe, expect, it } from 'vitest'
import { fillLineBoxes, measureBlocks, phantomRowResiduePx } from '../src/renderer/pagination'

const rectOf = (top: number, height: number, width = 400) =>
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

interface RowSpec {
  top: number
  height: number
  phantom?: 'gap' | 'header'
}

/** a table whose rows report the given rects; phantom rows carry the canvas classes */
function tableOf(top: number, height: number, rows: RowSpec[], attach = true): HTMLTableElement {
  const table = document.createElement('table')
  table.className = 'doc-table'
  // .doc-table collapses borders (jsdom has no stylesheet: 2 px spacing otherwise)
  table.style.borderCollapse = 'collapse'
  const tbody = document.createElement('tbody')
  for (const r of rows) {
    const tr = document.createElement('tr')
    if (r.phantom === 'gap') tr.className = 'page-gap page-gap-inline page-gap-table'
    if (r.phantom === 'header') tr.className = 'page-gap-inline page-repeat-header'
    tr.innerHTML = '<td>x</td>'
    tr.getBoundingClientRect = () => rectOf(r.top, r.height)
    tbody.appendChild(tr)
  }
  table.appendChild(tbody)
  table.getBoundingClientRect = () => rectOf(top, height)
  if (attach) document.body.appendChild(table)
  return table
}

describe('phantomRowResiduePx', () => {
  it('is the space between the real rows a phantom separates, less the phantom rects', () => {
    // rows 0-20 | phantom (rect 0) | 38.4-58.4: the table kept 18.4 px for the collapsed row
    const t = tableOf(0, 58.4, [
      { top: 0, height: 20 },
      { top: 20, height: 0, phantom: 'gap' },
      { top: 20, height: 0, phantom: 'header' },
      { top: 38.4, height: 20 },
    ])
    expect(phantomRowResiduePx(t)).toBeCloseTo(18.4, 5)
  })

  it('is zero when the phantoms collapse cleanly or keep their own height', () => {
    const clean = tableOf(0, 40, [
      { top: 0, height: 20 },
      { top: 20, height: 0, phantom: 'gap' },
      { top: 20, height: 20 },
    ])
    expect(phantomRowResiduePx(clean)).toBe(0)
    // the phantom rect carries the full height: the inline-gap sum already subtracts it
    const kept = tableOf(0, 140, [
      { top: 0, height: 20 },
      { top: 20, height: 100, phantom: 'gap' },
      { top: 120, height: 20 },
    ])
    expect(phantomRowResiduePx(kept)).toBe(0)
    const plain = tableOf(0, 40, [
      { top: 0, height: 20 },
      { top: 20, height: 20 },
    ])
    expect(phantomRowResiduePx(plain)).toBe(0)
  })

  it('ignores a nested table’s rows', () => {
    const outer = tableOf(0, 40, [
      { top: 0, height: 20 },
      { top: 20, height: 20 },
    ])
    const inner = tableOf(
      0,
      20,
      [
        { top: 0, height: 0, phantom: 'gap' },
        { top: 18, height: 2 },
      ],
      false,
    )
    inner.className = 'doc-table doc-nested-table'
    outer.querySelector('td')!.appendChild(inner)
    expect(phantomRowResiduePx(outer)).toBe(0)
  })
})

describe('measureBlocks — two tables, the first cut (phantom rows kept 18.4 px)', () => {
  it('the blocks after the cut table move up by the residue; a clean table changes nothing', () => {
    const pm = document.createElement('div')
    const cut = tableOf(
      0,
      58.4,
      [
        { top: 0, height: 20 },
        { top: 20, height: 0, phantom: 'gap' },
        { top: 38.4, height: 20 },
      ],
      false,
    )
    const p = document.createElement('p')
    p.className = 'doc-p-empty'
    p.innerHTML = '<br>'
    p.getBoundingClientRect = () => rectOf(58.4, 18)
    const second = tableOf(
      76.4,
      40,
      [
        { top: 76.4, height: 20 },
        { top: 96.4, height: 20 },
      ],
      false,
    )
    const after = document.createElement('p')
    after.textContent = 'after'
    after.getBoundingClientRect = () => rectOf(116.4, 18)
    for (const el of [cut, p, second, after]) pm.appendChild(el)
    document.body.appendChild(pm)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(
      blocks.map((b) => [Math.round(b.top * 10) / 10, Math.round(b.height * 10) / 10]),
    ).toEqual([
      [0, 40],
      [40, 18],
      [58, 40],
      [98, 18],
    ])
    // the residue sits between the rows, not in the row before the cut nor
    // taken from the last one: the rows still tile 20 / 20
    // jsdom has no Range.getClientRects (rowCutYs scans cell lines through it)
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = () => [] as unknown as DOMRectList
    try {
      fillLineBoxes(
        blocks,
        [{ contentHeight: 800, forceBreak: false }],
        1,
        undefined,
        undefined,
        new Set([blocks[0], blocks[2]]),
      )
    } finally {
      Range.prototype.getClientRects = orig
    }
    const rowsOf = (b: (typeof blocks)[number]) =>
      b.tableRows?.map((r) => Math.round(r.height * 10) / 10)
    expect(rowsOf(blocks[0])).toEqual([20, 20])
    expect(rowsOf(blocks[2])).toEqual([20, 20])
  })
})

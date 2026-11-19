import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { EditorView } from '@tiptap/pm/view'
import {
  nestedRowAtCut,
  nestedRowSplit,
  syncNestedTableGaps,
  tableGridCols,
  type NestedTableGap,
  type PageGapSpec,
} from '../src/renderer/editor/pagination-gaps'
import type { LineAnchor } from '../src/renderer/pagination'

// jsdom has no layout: every text line and box takes its top from the nearest
// data-top attribute, so a test lays rows and lines out by hand
const topOf = (n: Node | null): number => {
  const el = n instanceof Element ? n : (n?.parentElement ?? null)
  const at = el?.closest('[data-top]')?.getAttribute('data-top')
  return at == null ? 0 : Number(at)
}
const rect = (top: number) =>
  ({ top, bottom: top + 18, left: 0, right: 100, width: 100, height: 18, x: 0, y: top }) as DOMRect

let restore: Array<() => void> = []
beforeEach(() => {
  // jsdom's Range has no getClientRects at all: stub both through a loose view
  const range = Range.prototype as unknown as { getClientRects?: () => DOMRect[] }
  const elRect = Element.prototype.getBoundingClientRect
  const rangeRects = range.getClientRects
  Element.prototype.getBoundingClientRect = function (this: Element) {
    return rect(topOf(this))
  }
  range.getClientRects = function (this: Range) {
    return [rect(topOf(this.startContainer))]
  }
  restore = [
    () => (Element.prototype.getBoundingClientRect = elRect),
    () => (range.getClientRects = rangeRects),
  ]
})
afterEach(() => restore.forEach((f) => f()))

/** an outer cell holding a nested-table view; rows/lines carry data-top */
function nested(html: string): { view: EditorView; host: HTMLElement } {
  const pm = document.createElement('div')
  pm.innerHTML = `<table class="doc-table"><tbody><tr><td>
    <div class="doc-nested-table" contenteditable="false">
      <table class="doc-table"><tbody>${html}</tbody></table>
    </div></td></tr></tbody></table>`
  document.body.replaceChildren(pm)
  return {
    view: { dom: pm } as unknown as EditorView,
    host: pm.querySelector('.doc-nested-table') as HTMLElement,
  }
}
const textIn = (host: HTMLElement, sel: string): Text => host.querySelector(sel)!.firstChild as Text
const at = (node: Text | Element, charOffset = 0): LineAnchor => ({ node, charOffset })

const metrics = { marginTop: 96, marginBottom: 72, marginLeft: 96, marginRight: 96 }
const tableGap = (cols: number): PageGapSpec => ({ pos: 0, kind: 'table', cols, page: 1, metrics })
const cellGap = (liftPx = 0): PageGapSpec => ({
  pos: 0,
  kind: 'cell',
  page: 1,
  metrics: { ...metrics, marginBottom: metrics.marginBottom + liftPx },
})

describe('nestedRowAtCut', () => {
  const rows = `
    <tr data-top="100"><td><p id="h1">Peneliti</p></td><td><p id="h2">Peran</p></td></tr>
    <tr data-top="140"><td><p id="a" data-top="140">SRM</p></td>
      <td><p id="b" data-top="140">Ketua</p><p id="b2" data-top="160">Peneliti</p></td></tr>`

  it('takes the row a turn between two rows opens', () => {
    const { host } = nested(rows)
    const row = nestedRowAtCut(at(textIn(host, '#b')), 6)
    expect(row?.querySelector('#a')).not.toBeNull()
  })

  it('leaves a turn between the lines of one row to the cut marker', () => {
    const { host } = nested(rows)
    expect(nestedRowAtCut(at(textIn(host, '#b2')), 6)).toBeNull()
  })

  it('only answers inside a nested-table view', () => {
    nested(rows)
    const p = document.createElement('p')
    p.textContent = 'top level'
    document.body.appendChild(p)
    expect(nestedRowAtCut(at(p.firstChild as Text), 6)).toBeNull()
  })

  it('climbs to the view’s own table row from a deeper table', () => {
    const { host } = nested(`
      <tr data-top="100"><td><p>above</p></td></tr>
      <tr data-top="140"><td><table><tbody><tr data-top="140">
        <td><p id="deep">inner</p></td></tr></tbody></table></td></tr>`)
    const row = nestedRowAtCut(at(textIn(host, '#deep')), 6)
    expect(row?.parentElement?.parentElement).toBe(host.querySelector('table'))
  })
})

describe('nestedRowSplit', () => {
  it('finds each cell’s first line past the turn and how far above it starts', () => {
    const { host } = nested(`
      <tr data-top="100">
        <td><p data-top="100">SRM</p></td>
        <td><p data-top="100">Supervisi</p><p id="c2" data-top="140">kolaborasi</p></td>
        <td><p data-top="100">Arsitektur</p><p id="c3" data-top="136">collaboration</p></td>
      </tr>`)
    // the turn's own line is c2 at 140; c3 starts 4px higher (larger font), SRM ended above
    const lines: Record<string, LineAnchor> = {
      c2: at(textIn(host, '#c2')),
      c3: at(textIn(host, '#c3')),
    }
    const split = nestedRowSplit(lines.c2, 1, 6, (cell) => {
      const p = cell.querySelector('#c2, #c3')
      return p ? lines[p.id] : null
    })
    expect(split?.cells.map((c) => c.mark.cell.cellIndex)).toEqual([1, 2])
    expect(split?.cells.map((c) => c.liftPx)).toEqual([0, 4])
    // marks are character offsets into the cell's text, not node references
    expect(split?.cells[0].mark).toMatchObject({ text: 'Supervisi'.length })
  })
})

describe('syncNestedTableGaps', () => {
  it('puts a gap row before the row, keeps it on a settled pass, clears it', () => {
    const { view, host } = nested(`
      <tr><td>1</td><td>a</td></tr><tr id="r2"><td>2</td><td>b</td></tr>`)
    const row = host.querySelector('#r2') as HTMLTableRowElement
    const gap: NestedTableGap = { row, gap: tableGap(2) }
    syncNestedTableGaps(view, [gap])
    const el = row.previousElementSibling as HTMLElement
    expect(el.matches('tr.page-gap.page-gap-table.page-gap-nested')).toBe(true)
    expect((el.firstElementChild as HTMLTableCellElement).colSpan).toBe(2)
    syncNestedTableGaps(view, [gap])
    expect(row.previousElementSibling).toBe(el)
    syncNestedTableGaps(view, [])
    expect(host.querySelector('.page-gap')).toBeNull()
  })

  it('splits a row cell by cell and rejoins the text when the turn moves', () => {
    const { view, host } = nested(`
      <tr id="r"><td><p>Supervisi metodologi</p></td><td><p>Arsitektur integrasi</p></td></tr>`)
    const row = host.querySelector('#r') as HTMLTableRowElement
    const [c0, c1] = Array.from(row.cells)
    const split = (n0: number, n1: number): NestedTableGap => ({
      row,
      gap: cellGap(),
      cells: [
        { mark: { cell: c0, text: n0 }, gap: cellGap() },
        { mark: { cell: c1, text: n1 }, gap: cellGap(4) },
      ],
    })
    syncNestedTableGaps(view, [split(10, 11)])
    const gaps = Array.from(host.querySelectorAll('.page-gap-nested'))
    expect(gaps).toHaveLength(2)
    expect(gaps.every((g) => g.matches('.page-gap-cell'))).toBe(true)
    // each gap opens its cell's line: the text after it starts there
    expect(gaps[0].nextSibling?.textContent).toBe('metodologi')
    expect(gaps[1].nextSibling?.textContent).toBe('integrasi')
    expect(c0.textContent).toBe('Supervisi metodologi')

    // a settled pass keeps the same elements even though the text is split now
    syncNestedTableGaps(view, [split(10, 11)])
    expect(Array.from(host.querySelectorAll('.page-gap-nested'))).toEqual(gaps)

    // the turn moves: old gaps out, text rejoined, new ones in at the new marks
    syncNestedTableGaps(view, [split(0, 0)])
    expect(gaps.some((g) => g.isConnected)).toBe(false)
    expect(c0.querySelector('p')!.childNodes[1].textContent).toBe('Supervisi metodologi')

    syncNestedTableGaps(view, [])
    expect(c0.querySelector('p')!.childNodes).toHaveLength(1)
    expect(c1.querySelector('p')!.childNodes).toHaveLength(1)
  })

  it('leaves a detached row alone', () => {
    const { view, host } = nested(`<tr id="r"><td>x</td></tr>`)
    const row = host.querySelector('#r') as HTMLTableRowElement
    row.remove()
    syncNestedTableGaps(view, [{ row, gap: tableGap(1) }])
    expect(host.querySelector('.page-gap')).toBeNull()
  })
})

describe('tableGridCols', () => {
  it('counts the real grid, not gap rows or repeated headers', () => {
    const table = document.createElement('table')
    table.innerHTML = `<tbody>
      <tr class="page-gap"><td colspan="9"></td></tr>
      <tr class="page-repeat-header"><td colspan="7"></td></tr>
      <tr><td colspan="2"></td><td></td></tr>
      <tr><td></td><td></td></tr></tbody>`
    expect(tableGridCols(table)).toBe(3)
    expect(tableGridCols(null)).toBe(1)
  })
})

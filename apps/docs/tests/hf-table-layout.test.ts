import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfTableCell, HfTableRow } from '@genoffice/docx-engine'
import {
  hfCellGeometry,
  hfCellParaStyle,
  hfCellSegStyle,
  hfCellTabLines,
  hfRowStyle,
  makeGapHfEl,
} from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

const single = { style: 'single', szEighths: 12, color: 'auto' }

/** the right-hand "03.PP / Página 1 de 1<tab> de" cell of a 1980-twip column */
const pageCell: HfTableCell = {
  paras: [
    [
      { text: 'Página 1 de 1', bold: true, sizeHalfPoints: 24 },
      { text: '\t de  ', bold: true, sizeHalfPoints: 24 },
    ],
  ],
  paraProps: [{ indentLeft: 142 }],
  widthTwips: 1980,
  marTwips: { left: 70, right: 70 },
  vAlign: 'center',
  borders: { top: single, bottom: single, right: single },
}

describe('header layout-table cell geometry', () => {
  it('fixed column: declared width as px, margins as padding, borders, vAlign', () => {
    const geom = hfCellGeometry(pageCell)
    expect(geom.style.width).toBe('132.0px')
    expect(geom.style.flex).toBe('0 0 auto')
    expect(geom.style.padding).toBe('0.0px 4.7px 0.0px 4.7px')
    expect(geom.style.justifyContent).toBe('center')
    expect(geom.style.borderTop).toBe('2px solid #000')
    expect(geom.style.borderLeft).toBeUndefined()
    expect(geom.borders).toEqual({ t: '2px solid #000', b: '2px solid #000', r: '2px solid #000' })
    expect(geom.textWidthPx).toBeCloseTo((1980 - 140) / 15, 3)
  })

  it('proportional column keeps its percentage and no padding', () => {
    const geom = hfCellGeometry({ paras: [[{ text: 'x' }]], widthPct: 40 })
    expect(geom.style).toEqual({ width: '40%' })
    expect(geom.textWidthPx).toBeUndefined()
  })

  it('row: atLeast height floors, exact height clips, only a positive indent offsets', () => {
    expect(hfRowStyle({ heightTwips: 1428, heightRule: 'atLeast', indentTwips: 0 })).toEqual({
      minHeight: '95.2px',
    })
    expect(hfRowStyle({ heightTwips: 300, heightRule: 'exact', indentTwips: -108 })).toEqual({
      height: '20.0px',
      overflow: 'hidden',
    })
    expect(hfRowStyle({ indentTwips: 70 })).toEqual({ marginLeft: '4.7px' })
    expect(hfRowStyle({ indentTwips: 70, bidiVisual: true })).toEqual({
      marginRight: '4.7px',
      flexDirection: 'row-reverse',
    })
    expect(hfRowStyle(undefined)).toEqual({})
  })
})

describe('tabs inside a fixed-width cell', () => {
  const geom = hfCellGeometry(pageCell)
  const runs = pageCell.paras[0]

  it('legacy layout drops a segment that starts past the cell edge (Word clips it)', () => {
    const row: HfTableRow = { tabOverflow: 'clip' }
    const lines = hfCellTabLines(runs, pageCell.paraProps![0], geom, row)!
    expect(lines).toHaveLength(1)
    expect(lines[0].lead.map((r) => r.text)).toEqual(['Página 1 de 1'])
    expect(lines[0].segments).toHaveLength(0)
  })

  it('a segment crossing the cell edge keeps its box inside the cell', () => {
    // text ends at ~113px, the next default stop (144px) lies inside a 2400-twip column
    const wide = hfCellGeometry({ ...pageCell, widthTwips: 2400 })
    const lines = hfCellTabLines(runs, pageCell.paraProps![0], wide, { tabOverflow: 'clip' })!
    expect(lines[0].segments).toHaveLength(1)
    expect(hfCellSegStyle(lines[0].segments[0])).toEqual({
      left: '144.0px',
      maxWidth: 'calc(100% - 144.0px)',
    })
  })

  it('Word 2013+ layout wraps the text after the tab onto a new line at the indent', () => {
    const row: HfTableRow = { tabOverflow: 'wrap' }
    const lines = hfCellTabLines(runs, pageCell.paraProps![0], geom, row)!
    expect(lines).toHaveLength(2)
    expect(lines[0].segments).toHaveLength(0)
    expect(lines[1].lead.map((r) => r.text)).toEqual(['de  '])
    expect(lines[1].segments).toHaveLength(0)
  })

  it('a tab that fits stays on the line at the next default stop', () => {
    const wide = hfCellGeometry({ ...pageCell, widthTwips: 6000 })
    const lines = hfCellTabLines(runs, pageCell.paraProps![0], wide, { tabOverflow: 'wrap' })!
    expect(lines).toHaveLength(1)
    expect(lines[0].segments).toHaveLength(1)
  })

  it('paragraph spacing and first-line indent land once across wrapped lines', () => {
    const props = { spaceBefore: 240, spaceAfter: 120, indentFirstLine: 142, indentLeft: 142 }
    expect(hfCellParaStyle(props, { first: true, last: false })).toEqual({
      paddingLeft: '9.5px',
      textIndent: '9.5px',
      marginTop: '16.0px',
    })
    expect(hfCellParaStyle(props, { first: false, last: false })).toEqual({ paddingLeft: '9.5px' })
    expect(hfCellParaStyle(props, { first: false, last: true })).toEqual({
      paddingLeft: '9.5px',
      marginBottom: '8.0px',
    })
  })

  it('paragraphs without tabs are not tab-laid', () => {
    expect(hfCellTabLines([{ text: 'plain' }], undefined, geom, undefined)).toBeNull()
  })
})

describe('gap strip DOM for a bordered fixed header table', () => {
  const value: HeaderFooter = {
    text: '',
    paras: [
      {
        runs: [],
        row: { heightTwips: 1428, heightRule: 'atLeast', indentTwips: 0, tabOverflow: 'wrap' },
        cells: [
          {
            paras: [[{ text: 'LOGO' }]],
            widthTwips: 3544,
            marTwips: { left: 70, right: 70 },
            borders: { top: single, left: single, bottom: single, right: single },
          },
          pageCell,
        ],
      },
    ],
  }

  it('paints borders and widths on the cells and floors the row height', () => {
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const row = el.querySelector('.page-hf-row') as HTMLElement
    expect(row.style.minHeight).toBe('95.2px')
    const cells = el.querySelectorAll<HTMLElement>('.page-hf-cell')
    expect(cells[0].style.width).toBe('236.3px')
    expect(cells[0].style.borderLeft).toMatch(/^2px solid/)
    expect(cells[1].style.borderLeft).toBe('')
    expect(cells[1].style.borderRight).toMatch(/^2px solid/)
    // the wrapped "de" becomes a second tab-laid line
    const lines = cells[1].querySelectorAll('.page-hf-cell-para.page-hf-tabbed')
    expect(lines).toHaveLength(2)
    expect(lines[1].textContent).toBe('de  ')
  })

  it('estimateHfHeight reserves at least the declared row height', () => {
    expect(estimateHfHeight(value, 680)).toBeGreaterThanOrEqual(1428 / 15)
  })

  it('estimateHfHeight fixes an exact row at its declared height (content is clipped)', () => {
    const tall = { paras: [[{ text: 'one' }], [{ text: 'two' }], [{ text: 'three' }]] }
    const exact = {
      text: '',
      paras: [{ runs: [], row: { heightTwips: 150, heightRule: 'exact' as const }, cells: [tall] }],
    }
    expect(estimateHfHeight(exact, 680)).toBe(10)
    const atLeast = {
      text: '',
      paras: [
        { runs: [], row: { heightTwips: 150, heightRule: 'atLeast' as const }, cells: [tall] },
      ],
    }
    expect(estimateHfHeight(atLeast, 680)).toBeGreaterThan(10)
  })
})

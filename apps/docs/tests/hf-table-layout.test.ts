import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfTableCell, HfTableRow } from '@genoffice/docx-engine'
import {
  hfCellGeometry,
  hfCellParaStyle,
  hfCellSegStyle,
  hfCellTabLines,
  hfIndentOverhangPx,
  hfProbeGeomKey,
  hfRowRightOverhangPx,
  hfRowStyle,
  hfRunFace,
  makeGapHfEl,
  type HfRowGeom,
} from '../src/renderer/editor/hf-dom'
import { fromHfParagraph, toStyleRecord } from '../src/renderer/editor/para-layout-spec'
import { estimateHfHeight, lineHeightFactor } from '../src/renderer/line-metrics'

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

  it('diagonal lines paint as background layers with a dark-page twin for callers', () => {
    const geom = hfCellGeometry({ ...pageCell, borders: { tl2br: single } })
    expect(geom.diagonals).toContain('linear-gradient(to top right')
    expect(geom.style.backgroundImage).toBe(geom.diagonals)
    expect(hfCellGeometry(pageCell).diagonals).toBeUndefined()
  })

  it('proportional column keeps its percentage and no padding', () => {
    const geom = hfCellGeometry({ paras: [[{ text: 'x' }]], widthPct: 40 })
    expect(geom.style).toEqual({ width: '40%' })
    expect(geom.textWidthPx).toBeUndefined()
  })

  it('row: atLeast height floors, exact height clips, a signed indent offsets the edge', () => {
    expect(hfRowStyle({ heightTwips: 1428, heightRule: 'atLeast', indentTwips: 0 })).toEqual({
      minHeight: '95.2px',
    })
    // negative tblInd hangs the table into the margin (prod corpus 20261021 sample 017 footer)
    expect(hfRowStyle({ heightTwips: 300, heightRule: 'exact', indentTwips: -108 })).toEqual({
      marginLeft: '-7.2px',
      height: '20.0px',
      overflow: 'hidden',
    })
    expect(hfIndentOverhangPx([{ runs: [], cells: [], row: { indentTwips: -1276 } }])).toBeCloseTo(
      85.07,
      1,
    )
    expect(hfIndentOverhangPx([{ runs: [], cells: [], row: { indentTwips: 70 } }])).toBe(0)
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
    // the line box (factor, line-height) is the adapter's; only the geometry varies per line
    const geometry = (style: Record<string, string>) => {
      const { ['--doc-line-factor']: _f, lineHeight: _lh, direction: _d, ...rest } = style
      return rest
    }
    expect(geometry(hfCellParaStyle(props, { first: true, last: false }))).toEqual({
      paddingLeft: '9.5px',
      textIndent: '9.5px',
      marginTop: '16.0px',
    })
    expect(geometry(hfCellParaStyle(props, { first: false, last: false }))).toEqual({
      paddingLeft: '9.5px',
    })
    expect(geometry(hfCellParaStyle(props, { first: false, last: true }))).toEqual({
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

describe('positioned header tables (w:tblpPr / w:jc)', () => {
  // A4, 1 inch margins: 595 pt page, 451 pt column, header strip top 17 pt
  const geom: HfRowGeom = {
    kind: 'header',
    columnPx: 601.7,
    stripLeftPx: 96,
    pageWPx: 793.7,
    marginRightPx: 96,
    marginTopPx: 96,
    headerStripTopPx: 22.7,
  }
  const pageAnchored: HfTableRow = {
    widthTwips: 11058,
    float: { xTwips: 464, yTwips: 346, horzAnchor: 'page', vertAnchor: 'page' },
    source: { tbl: 0, tr: 0 },
  }

  it('a page-anchored table sits at its page x, hanging into the margin and past the column', () => {
    const style = hfRowStyle(pageAnchored, geom, true)
    expect(style.marginLeft).toBe('-65.1px')
    expect(style.marginTop).toBeUndefined()
    const paras = [{ runs: [], cells: [], row: pageAnchored }]
    expect(hfIndentOverhangPx(paras, geom)).toBeCloseTo(65.1, 1)
    // 30.9 + 737.2 - 601.7 - 96 = 70.4 past the column, within the right margin
    expect(hfRowRightOverhangPx(paras, geom)).toBeCloseTo(70.4, 1)
  })

  it('w:tblpY below the strip top becomes a gap above the first row only', () => {
    const low: HfTableRow = { ...pageAnchored, float: { ...pageAnchored.float, yTwips: 1200 } }
    expect(hfRowStyle(low, geom, true).marginTop).toBe('57.3px')
    expect(hfRowStyle(low, geom, false).marginTop).toBeUndefined()
    expect(hfRowStyle(low, { ...geom, kind: 'footer' }, true).marginTop).toBeUndefined()
  })

  it('the strip-height cache keys a page-anchored table on header distance and top margin', () => {
    const value: HeaderFooter = { text: '', paras: [{ runs: [], cells: [], row: pageAnchored }] }
    const a = hfProbeGeomKey(value, { headerDistPx: 22.7, marginTopPx: 96 })
    const b = hfProbeGeomKey(value, { headerDistPx: 48, marginTopPx: 96 })
    const c = hfProbeGeomKey(value, { headerDistPx: 22.7, marginTopPx: 120 })
    expect(new Set([a, b, c]).size).toBe(3)
    const inFlow: HeaderFooter = {
      text: '',
      paras: [{ runs: [], cells: [], row: { widthTwips: 11058 } }],
    }
    expect(hfProbeGeomKey(inFlow, { headerDistPx: 22.7, marginTopPx: 96 })).toBe('')
    expect(hfProbeGeomKey(inFlow, { headerDistPx: 48, marginTopPx: 96 })).toBe('')
  })

  it('a margin-anchored centred table centres on the column; without geometry it stays in flow', () => {
    const centred: HfTableRow = {
      widthTwips: 10975,
      indentTwips: -200,
      float: { xSpec: 'center', horzAnchor: 'margin', vertAnchor: 'page', yTwips: 15252 },
    }
    expect(hfRowStyle(centred, geom).marginLeft).toBe('-65.0px')
    expect(hfRowStyle(centred).marginLeft).toBe('-13.3px')
  })

  it('w:jc centres the row in the column and reports the overhang of an over-wide table', () => {
    const jc: HfTableRow = { align: 'center', widthTwips: 9185 }
    expect(hfRowStyle(jc, geom).justifyContent).toBe('center')
    expect(hfRowStyle(jc, geom).marginLeft).toBeUndefined()
    const wide = [{ runs: [], cells: [], row: { align: 'center' as const, widthTwips: 10000 } }]
    expect(hfIndentOverhangPx(wide, geom)).toBeCloseTo(32.5, 1)
    expect(hfRowRightOverhangPx(wide, geom)).toBeCloseTo(32.5, 1)
  })
})

describe('strip run faces from the paragraph style', () => {
  const effective = {
    cs: false,
    bold: false,
    italic: false,
    sizeHalfPoints: 20,
    fonts: { ascii: 'Malgun Gothic', eastAsia: 'Malgun Gothic' },
    underline: false,
    strike: false,
    caps: 'none' as const,
    vanish: false,
    charSpacingTwips: 0,
    source: {
      sizeHalfPoints: 'paraStyle',
      ascii: 'paraStyle',
      eastAsia: 'paraStyle',
      bold: 'builtin',
      italic: 'builtin',
    },
  }

  it('a run without its own w:sz / w:rFonts takes the footer style size and face', () => {
    expect(hfRunFace({ text: 'page', effective } as never)).toEqual({
      sizeHalfPoints: 20,
      fontAscii: 'Malgun Gothic',
      font: 'Malgun Gothic',
      eastAsiaFont: 'Malgun Gothic',
    })
  })

  it('a style declaring both slots keeps the East Asian face for the hangul line factor', () => {
    const runs = [{ text: '\ud1b5\ud569\uc0ac\ud68c 2', bold: true, effective }] as never[]
    expect(toStyleRecord(fromHfParagraph({ runs: runs as never }))['--doc-line-factor']).toBe(
      String(lineHeightFactor('Malgun Gothic')),
    )
  })

  it('direct properties win and docDefaults-sourced values stay with the strip default', () => {
    expect(
      hfRunFace({ text: 'x', sizeHalfPoints: 28, bold: true, effective } as never),
    ).toMatchObject({ sizeHalfPoints: 28, bold: true })
    const dd = { ...effective, source: { ...effective.source, sizeHalfPoints: 'docDefaults' } }
    expect(hfRunFace({ text: 'x', effective: dd } as never).sizeHalfPoints).toBeUndefined()
  })

  it('a w:rtl run of only digits, dashes and spaces keeps the strip default size', () => {
    // footer "- 9 -" under Normal sz=32 szCs=22: Word draws it 16pt, not the 11pt szCs
    const cs = {
      ...effective,
      cs: true,
      sizeHalfPoints: 22,
      fonts: { ascii: 'Traditional Arabic', cs: 'Traditional Arabic' },
      source: { ...effective.source, cs: 'direct' },
    }
    expect(hfRunFace({ text: '\u2014 9 \u2014', effective: cs } as never)).toEqual({})
    expect(hfRunFace({ text: '  ', effective: cs } as never)).toEqual({})
    expect(hfRunFace({ text: 'page', effective: cs } as never)).toEqual({
      sizeHalfPoints: 22,
      fontAscii: 'Traditional Arabic',
    })
  })

  it('an empty paragraph is sized by its mark face, keeping its own line multiple', () => {
    const empty = (para: object) => toStyleRecord(fromHfParagraph({ runs: [], ...para }))
    const latin = empty({ emptyRunFont: 'Calibri', lineRule: 'auto', lineRawTwips: 300 })
    expect(Number(latin['--doc-line-factor'])).toBeCloseTo(1.22, 2)
    expect(latin.lineHeight).toBe('calc(var(--doc-line-factor,1.2) * 1.25)')
    expect(empty({ emptyRunFont: 'Calibri' }).lineHeight).toContain('--doc-line-factor')
    expect(empty({})).toEqual({
      direction: 'ltr',
      '--doc-line-factor': 'var(--doc-line-factor-latin,1.2)',
      lineHeight: 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    })
  })
})

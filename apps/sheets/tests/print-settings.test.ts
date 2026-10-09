import { describe, expect, it } from 'vitest'

import {
  assemblePrintHtml,
  buildSheetPrintPayload,
  layoutSheetPrint,
  pageVariant,
  renderHeaderFooterSection,
  sectionPictures,
} from '../src/renderer/print-html'
import {
  decodeHeaderFooter,
  printAreasFromFormula,
  printTitleRowsFromFormula,
  resolveEffectivePageSetup,
  type EffectivePageSetup,
} from '../src/renderer/print-settings'
import type { PrintWorksheet } from '../src/renderer/print-html'
import type { PrintVisualSnapshot } from '../src/renderer/print-visuals'

describe('decodeHeaderFooter', () => {
  it('splits left/center/right sections and keeps resolvable field codes', () => {
    expect(decodeHeaderFooter('&L&A&CSeite &P von &N')).toEqual({
      left: '&A',
      center: 'Seite &P von &N',
    })
  })

  it('defaults unmarked text to the center section', () => {
    expect(decodeHeaderFooter('Quarterly Report')).toEqual({ center: 'Quarterly Report' })
  })

  it('strips font, size, color and unsupported codes', () => {
    expect(decodeHeaderFooter('&C&"Broadway,Bold Italic"&12&KFF0000Big &BRed&B Title&Z')).toEqual({
      center: 'Big Red Title',
    })
    expect(decodeHeaderFooter('&C&OOut&O &HShade&H')).toEqual({ center: 'Out Shade' })
  })

  it('drops the path code but keeps the file and picture codes', () => {
    expect(decodeHeaderFooter('&C&P / &N&R&Z&F&G')).toEqual({
      center: '&P / &N',
      right: '&F&G',
    })
    expect(decodeHeaderFooter('&L&G')).toEqual({ left: '&G' })
  })

  it('keeps escaped ampersands verbatim', () => {
    expect(decodeHeaderFooter('&LProfit && Loss')).toEqual({ left: 'Profit && Loss' })
  })

  it('returns null when everything strips away', () => {
    expect(decodeHeaderFooter('&L&Z')).toBeNull()
    expect(decodeHeaderFooter('')).toBeNull()
  })

  it('accepts lowercase section markers', () => {
    expect(decodeHeaderFooter('&lLeft&rRight')).toEqual({ left: 'Left', right: 'Right' })
  })

  it('keeps the character after an unrecognised code as literal text', () => {
    expect(decodeHeaderFooter('Tom&Jerry')).toEqual({ center: 'Tom&Jerry' })
    expect(decodeHeaderFooter('R&D')).toEqual({ center: 'R&D' })
    expect(decodeHeaderFooter('Tom & Jerry')).toEqual({ center: 'Tom & Jerry' })
    expect(decodeHeaderFooter('&L50% & up')).toEqual({ left: '50% & up' })
    expect(decodeHeaderFooter('&LProfit && Loss')).toEqual({ left: 'Profit && Loss' })
    expect(decodeHeaderFooter('&C&P of &N in &F')).toEqual({ center: '&P of &N in &F' })
    expect(decodeHeaderFooter('&C&"Broadway,Bold"&12&KFF0000Big &BRed&B Title&Z')).toEqual({
      center: 'Big Red Title',
    })
    expect(decodeHeaderFooter('A&B')).toEqual({ center: 'A' })
  })
})

describe('printAreasFromFormula', () => {
  it('parses a quoted sheet-qualified absolute range', () => {
    expect(printAreasFromFormula("'W PS Mustermann Hans'!$A$1:$K$84")).toEqual(['A1:K84'])
  })

  it('parses multiple areas and quoted commas', () => {
    expect(printAreasFromFormula("'a,b'!$A$1:$B$2,'a,b'!$D$3:$E$4")).toEqual(['A1:B2', 'D3:E4'])
  })

  it('expands a single-cell area', () => {
    expect(printAreasFromFormula('Sheet1!$B$2')).toEqual(['B2:B2'])
  })

  it('falls back to the used range for refs it cannot crop to', () => {
    expect(printAreasFromFormula("'S'!$A:$C")).toEqual([])
    expect(printAreasFromFormula("'S'!#REF!")).toEqual([])
  })

  it('skips #REF! parts but keeps the used-range fallback for uncroppable ones', () => {
    expect(printAreasFromFormula("'S'!$A$1:$B$2,'S'!$C:$D")).toEqual([])
    expect(printAreasFromFormula("'S'!#REF!,'S'!$D$3")).toEqual(['D3:D3'])
  })

  it('returns [] when absent', () => {
    expect(printAreasFromFormula(undefined)).toEqual([])
  })

  it('normalises an area whose corners are reversed', () => {
    expect(printAreasFromFormula("'S'!$B$4:$A$1")).toEqual(['A1:B4'])
    expect(printAreasFromFormula("'S'!$D$3:$B$9")).toEqual(['B3:D9'])
    expect(printAreasFromFormula("'S'!$B$2:$B$2")).toEqual(['B2:B2'])
  })
})

describe('printTitleRowsFromFormula', () => {
  it('extracts the repeated row span', () => {
    expect(printTitleRowsFromFormula("'S'!$17:$17")).toBe('17:17')
  })

  it('skips a column-repeat part and finds the rows', () => {
    expect(printTitleRowsFromFormula("'S'!$A:$B,'S'!$1:$3")).toBe('1:3')
  })

  it('drops spans beyond the layout title cap', () => {
    expect(printTitleRowsFromFormula("'S'!$1:$40")).toBeNull()
  })

  it('returns null when absent or column-only', () => {
    expect(printTitleRowsFromFormula(undefined)).toBeNull()
    expect(printTitleRowsFromFormula("'S'!$A:$B")).toBeNull()
  })
})

describe('resolveEffectivePageSetup', () => {
  it('defaults to A4 portrait at 100% with normal margins', () => {
    const setup = resolveEffectivePageSetup({}, null, null)
    expect(setup.orientation).toBe('portrait')
    expect(setup.paperSize).toBe(9)
    expect(setup.scale).toBe(100)
    expect(setup.fitToPage).toBe(false)
    expect(setup.margins).toEqual({
      left: 0.7,
      right: 0.7,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3,
    })
    expect(setup.printAreas).toEqual([])
    expect(setup.header).toBeNull()
    expect(setup.footer).toBeNull()
  })

  it('applies the saved file settings when the session touched nothing', () => {
    const setup = resolveEffectivePageSetup(
      {},
      {
        orientation: 'landscape',
        paperSize: 1,
        scale: 65,
        margins: { left: 0.98, right: 0.98, top: 5, bottom: 0.79, header: 0, footer: 0.51 },
        printGridlines: true,
        oddFooter: '&CSeite &P von &N',
      },
      { printArea: "'S'!$A$1:$K$84", printTitles: "'S'!$17:$17" },
    )
    expect(setup.orientation).toBe('landscape')
    expect(setup.paperSize).toBe(1)
    expect(setup.scale).toBe(65)
    // margins clamp to the export wire's 3in cap
    expect(setup.margins.top).toBe(3)
    expect(setup.margins.left).toBe(0.98)
    expect(setup.printGridlines).toBe(true)
    expect(setup.printAreas).toEqual(['A1:K84'])
    expect(setup.printTitles).toBe('17:17')
    expect(setup.footer).toEqual({ center: 'Seite &P von &N' })
    expect(setup.header).toBeNull()
    // scaleWithDoc is Excel's default; only an explicit "0" pins the size
    expect(setup.headerFooterScaleWithDoc).toBe(true)
    expect(
      resolveEffectivePageSetup({}, { headerFooterFixedSize: true }, null).headerFooterScaleWithDoc,
    ).toBe(false)
  })

  it('lets the session journal win over the file', () => {
    const setup = resolveEffectivePageSetup(
      {
        orientation: 'portrait',
        printArea: 'B2:C3',
        printTitles: null,
        header: null,
        margins: 'narrow',
      },
      {
        orientation: 'landscape',
        margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
        oddHeader: '&CFile Header',
      },
      { printArea: "'S'!$A$1:$K$84", printTitles: "'S'!$1:$2" },
    )
    expect(setup.orientation).toBe('portrait')
    expect(setup.printAreas).toEqual(['B2:C3'])
    expect(setup.printTitles).toBeNull()
    expect(setup.header).toBeNull()
    expect(setup.margins.left).toBe(0.25)
  })

  it('clears the file print area when the session cleared it', () => {
    const setup = resolveEffectivePageSetup({ printArea: null }, null, {
      printArea: "'S'!$A$1:$K$84",
    })
    expect(setup.printAreas).toEqual([])
  })

  it('defaults fitToWidth/fitToHeight to one page when the file only sets fitToPage', () => {
    const setup = resolveEffectivePageSetup({}, { fitToPage: true }, null)
    expect(setup.fitToPage).toBe(true)
    expect(setup.fitToWidth).toBe(1)
    expect(setup.fitToHeight).toBe(1)
    const heightOnly = resolveEffectivePageSetup(
      {},
      { fitToPage: true, fitToWidth: 0, fitToHeight: 1 },
      null,
    )
    expect(heightOnly.fitToWidth).toBe(0)
    expect(heightOnly.fitToHeight).toBe(1)
    expect(resolveEffectivePageSetup({}, null, null).firstPage).toBeNull()
    expect(resolveEffectivePageSetup({}, null, null).evenPages).toBeNull()
    expect(resolveEffectivePageSetup({}, null, null).headerFooterPictures).toEqual([])
  })

  it('exposes the first/even page variants only when their flag is set', () => {
    const file = {
      oddHeader: '&L&G&COdd',
      oddFooter: '&P',
      evenHeader: '&CEven',
      firstFooter: '&RFirst &D',
      headerFooterPictures: [
        {
          id: 'hf-picture-0-lh',
          position: 'LH',
          widthPt: 442.5,
          heightPt: 43.5,
          mediaType: 'image/png',
        },
      ],
    }
    const off = resolveEffectivePageSetup({}, file, null)
    expect(off.header).toEqual({ left: '&G', center: 'Odd' })
    expect(off.firstPage).toBeNull()
    expect(off.evenPages).toBeNull()
    expect(off.headerFooterPictures).toEqual(file.headerFooterPictures)

    const on = resolveEffectivePageSetup(
      {},
      { ...file, differentOddEven: true, differentFirst: true },
      null,
    )
    // differentFirst with no firstHeader: page 1 prints no header at all.
    expect(on.firstPage).toEqual({ header: null, footer: { right: 'First &D' } })
    expect(on.evenPages).toEqual({ header: { center: 'Even' }, footer: null })
  })

  it('keeps the file variants when the session edits the odd header', () => {
    const setup = resolveEffectivePageSetup(
      { header: { center: 'Session' }, footer: null },
      { differentFirst: true, oddHeader: '&COdd', firstHeader: '&CFirst', oddFooter: '&P' },
      null,
    )
    expect(setup.header).toEqual({ center: 'Session' })
    expect(setup.footer).toBeNull()
    expect(setup.firstPage).toEqual({ header: { center: 'First' }, footer: null })
  })

  it('shifts the file print area through rows inserted above', () => {
    const setup = resolveEffectivePageSetup({}, null, { printArea: "'S'!$A$1:$K$84" }, [
      { kind: 'insert-rows', index: 0, count: 2 },
    ])
    expect(setup.printAreas).toEqual(['A3:K86'])
  })

  it('shrinks the file print area when a column inside it is deleted', () => {
    const setup = resolveEffectivePageSetup({}, null, { printArea: "'S'!$A$1:$K$84" }, [
      { kind: 'remove-cols', index: 2, count: 1 },
    ])
    expect(setup.printAreas).toEqual(['A1:J84'])
  })

  it('falls back to the used range when the edits delete the whole area', () => {
    const setup = resolveEffectivePageSetup({}, null, { printArea: "'S'!$B$2:$C$3" }, [
      { kind: 'remove-rows', index: 1, count: 2 },
    ])
    expect(setup.printAreas).toEqual([])
  })

  it('shifts the file title rows through structural edits', () => {
    const setup = resolveEffectivePageSetup({}, null, { printTitles: "'S'!$17:$17" }, [
      { kind: 'insert-rows', index: 0, count: 3 },
    ])
    expect(setup.printTitles).toBe('20:20')
  })

  it('keeps title rows through column edits and drops them when deleted', () => {
    const columnEdit = resolveEffectivePageSetup({}, null, { printTitles: "'S'!$1:$2" }, [
      { kind: 'remove-cols', index: 0, count: 3 },
    ])
    expect(columnEdit.printTitles).toBe('1:2')
    const deleted = resolveEffectivePageSetup({}, null, { printTitles: "'S'!$1:$2" }, [
      { kind: 'remove-rows', index: 0, count: 2 },
    ])
    expect(deleted.printTitles).toBeNull()
  })

  it('does not remap session-set print areas (already screen space)', () => {
    const setup = resolveEffectivePageSetup({ printArea: 'B2:C3' }, null, null, [
      { kind: 'insert-rows', index: 0, count: 5 },
    ])
    expect(setup.printAreas).toEqual(['B2:C3'])
  })

  it('prints a reversed print area instead of refusing the export', () => {
    const setup = resolveEffectivePageSetup({}, null, { printArea: "'S'!$B$4:$A$1" })
    expect(setup.printAreas).toEqual(['A1:B4'])
    const payload = buildSheetPrintPayload(fakeWorksheet(), setup, 'Book.pdf', 'S1')
    expect(payload.html).toContain('<table>')
    expect(payload.html).toContain('A1')
    expect(payload.html).toContain('B3')
  })

  it('drops title rows stretched past the cap by inserts between them', () => {
    const setup = resolveEffectivePageSetup({}, null, { printTitles: "'S'!$1:$2" }, [
      { kind: 'insert-rows', index: 1, count: 25 },
    ])
    expect(setup.printTitles).toBeNull()
    // the export payload still builds instead of throwing on the span
    const payload = buildSheetPrintPayload(fakeWorksheet(), setup, 'Book.pdf', 'S1')
    expect(payload.html).toContain('<table>')
  })
})

describe('renderHeaderFooterSection', () => {
  const now = new Date(2026, 0, 2, 3, 4, 5)
  const fields = { page: 3, total: 7, date: now, fileName: 'Book', sheetName: 'S1' }

  it('resolves page codes for the page being rendered', () => {
    expect(renderHeaderFooterSection('Seite &P von &N', fields)).toBe('Seite 3 von 7')
  })

  it('resolves static codes and escapes markup', () => {
    const html = renderHeaderFooterSection('&A <&F> && more', {
      ...fields,
      fileName: 'Bud<get',
      sheetName: 'Sh&eet',
    })
    expect(html).toBe('Sh&amp;eet &lt;Bud&lt;get&gt; &amp; more')
  })

  it('replaces &G with the section picture at its declared size', () => {
    const picture = { dataUrl: 'data:image/png;base64,AAAA', widthPt: 72, heightPt: 36 }
    expect(renderHeaderFooterSection('&G Logo', fields, picture)).toBe(
      '<img src="data:image/png;base64,AAAA" style="width:96px;height:48px"> Logo',
    )
  })

  it('prints nothing for &G when the slot has no picture', () => {
    expect(renderHeaderFooterSection('&GTitle', fields)).toBe('Title')
  })
})

describe('sectionPictures', () => {
  const picture = (name: string) => ({ dataUrl: `data:${name}`, widthPt: 1, heightPt: 1 })
  const pictures = new Map([
    ['LH', picture('LH')],
    ['RF', picture('RF')],
    ['CHFIRST', picture('CHFIRST')],
    ['LFEVEN', picture('LFEVEN')],
  ])

  it('maps the VML slots onto the sections of each variant', () => {
    expect(sectionPictures(pictures, 'header', 'odd')).toEqual({ left: picture('LH') })
    expect(sectionPictures(pictures, 'footer', 'odd')).toEqual({ right: picture('RF') })
    expect(sectionPictures(pictures, 'header', 'first')).toEqual({ center: picture('CHFIRST') })
    expect(sectionPictures(pictures, 'footer', 'first')).toEqual({})
    expect(sectionPictures(pictures, 'footer', 'even')).toEqual({ left: picture('LFEVEN') })
    expect(sectionPictures(pictures, 'header', 'even')).toEqual({})
  })
})

const usedGrid = [
  ['A1', 'B1'],
  ['A2', 'B2'],
  ['A3', 'B3'],
]

function fakeWorksheet(): PrintWorksheet {
  return {
    getLastRow: () => 2,
    getLastColumn: () => 1,
    getRowHeight: () => 20,
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        usedGrid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () =>
        usedGrid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getCellStyleData: () => null,
    })) as PrintWorksheet['getRange'],
  }
}

function payloadSetup(overrides: Partial<EffectivePageSetup>): EffectivePageSetup {
  return {
    orientation: 'portrait',
    paperSize: 9,
    scale: 100,
    fitToWidth: 0,
    fitToHeight: 0,
    fitToPage: false,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    printGridlines: false,
    printHeadings: false,
    printAreas: [],
    printTitles: null,
    printTitleColumns: null,
    header: null,
    footer: null,
    firstPage: null,
    evenPages: null,
    headerFooterScaleWithDoc: true,
    headerFooterPictures: [],
    ...overrides,
  }
}

/// `rows` rows of 20px (15pt) with a value in column A.
function tallWorksheet(rows: number): PrintWorksheet {
  return {
    getLastRow: () => rows - 1,
    getLastColumn: () => 0,
    getRowHeight: () => 20,
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: ((row: number, _column: number, numRows?: number) => ({
      getDisplayValues: () => Array.from({ length: numRows ?? 1 }, (_, i) => [`R${row + i}`]),
      getValues: () => Array.from({ length: numRows ?? 1 }, (_, i) => [`R${row + i}`]),
      getCellStyleData: () => null,
    })) as PrintWorksheet['getRange'],
  }
}

const A4 = { widthIn: 8.27, heightIn: 11.69 }

function layout(
  worksheet: PrintWorksheet,
  setup: EffectivePageSetup,
  pictures = new Map<string, { dataUrl: string; widthPt: number; heightPt: number }>(),
  visuals: PrintVisualSnapshot = { visuals: [], css: '' },
  options = {},
) {
  return layoutSheetPrint(worksheet, setup, 'Book', 'S1', pictures, visuals, {
    now: new Date(2026, 0, 2),
    ...options,
  })
}

describe('layoutSheetPrint', () => {
  it('maps common OOXML paper sizes instead of falling back to A4', () => {
    expect(layout(fakeWorksheet(), payloadSetup({ paperSize: 12 })).paper.pageSize).toEqual({
      width: 9.84,
      height: 13.9,
    })
    expect(layout(fakeWorksheet(), payloadSetup({ paperSize: 13 })).paper.pageSize).toEqual({
      width: 7.17,
      height: 10.12,
    })
    expect(layout(fakeWorksheet(), payloadSetup({ paperSize: 14 })).paper.pageSize).toEqual({
      width: 8.5,
      height: 13,
    })
    expect(layout(fakeWorksheet(), payloadSetup({ paperSize: 6 })).paper.pageSize).toEqual({
      width: 5.5,
      height: 8.5,
    })
    const landscape = layout(fakeWorksheet(), payloadSetup({ orientation: 'landscape' })).paper
    expect(landscape.widthIn).toBe(A4.heightIn)
    expect(landscape.heightIn).toBe(A4.widthIn)
  })

  it('crops the layout to the print area', () => {
    const html = assemblePrintHtml([
      layout(fakeWorksheet(), payloadSetup({ printAreas: ['A1:A2'] })),
    ]).html
    expect(html).toContain('A1')
    expect(html).toContain('A2')
    expect(html).not.toContain('B1')
    expect(html).not.toContain('A3')
  })

  it('prints a selection instead of the print area', () => {
    const html = assemblePrintHtml([
      layout(fakeWorksheet(), payloadSetup({ printAreas: ['A1:A2'] }), new Map(), undefined, {
        selection: 'B2:B3',
      }),
    ]).html
    expect(html).toContain('B2')
    expect(html).toContain('B3')
    expect(html).not.toContain('A1')
  })

  it('emits one page per print area', () => {
    const document = layout(fakeWorksheet(), payloadSetup({ printAreas: ['A1:A1', 'B2:B2'] }))
    expect(document.pages).toHaveLength(2)
    const html = assemblePrintHtml([document]).html
    expect(html.match(/<div class="page /g)).toHaveLength(2)
  })

  it('paginates wide sheets across columns and honours manual breaks', () => {
    // 20 columns of 100px = 1500pt on a ~495pt wide A4 portrait page.
    const wide: PrintWorksheet = {
      ...fakeWorksheet(),
      getLastColumn: () => 19,
      getLastRow: () => 9,
    }
    const plain = layout(wide, payloadSetup({}))
    expect(plain.pages.length).toBeGreaterThan(1)
    expect(plain.pages.every((page) => page.page.colEnd - page.page.colStart < 7)).toBe(true)
    const broken = layout(wide, payloadSetup({}), new Map(), undefined, {
      breaks: { rowBreaks: [5], colBreaks: [] },
    })
    expect(broken.pages.length).toBe(plain.pages.length * 2)
    expect(broken.pages[0]?.page.rowEnd).toBe(4)
    expect(broken.pages[1]?.page.rowStart).toBe(5)
  })

  it('repeats title rows and columns on every page and skips hidden rows', () => {
    const grid = Array.from({ length: 30 }, (_, r) =>
      Array.from({ length: 3 }, (_, c) => `R${r + 1}C${c + 1}`),
    )
    const worksheet: PrintWorksheet = {
      getLastRow: () => 29,
      getLastColumn: () => 2,
      getRowHeight: () => 400,
      getColumnWidth: () => 100,
      getSheet: () => ({ getRowVisible: (row) => row !== 2, getColVisible: () => true }),
      getMergedRanges: () => [],
      getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
        getDisplayValues: () =>
          grid
            .slice(row, row + (numRows ?? 1))
            .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
        getValues: () => [],
        getCellStyleData: () => null,
      })) as PrintWorksheet['getRange'],
    }
    const document = layout(
      worksheet,
      payloadSetup({ printTitles: '1:1', printTitleColumns: 'A:A', printHeadings: true }),
    )
    // 300pt rows on a 733pt page: two body rows per page.
    expect(document.pages.length).toBeGreaterThan(10)
    const html = assemblePrintHtml([document]).html
    expect(html.match(/R1C2/g)?.length).toBe(document.pages.length)
    expect(html.match(/R1C1/g)?.length).toBe(document.pages.length)
    expect(html.match(/R5C1/g)?.length).toBe(1)
    expect(html).not.toContain('R3C2')
    expect(html).toContain('<th class="hd hd-top">B</th>')
  })

  it('gives every shared edge one owner when gridlines and custom borders meet', () => {
    const worksheet: PrintWorksheet = {
      ...fakeWorksheet(),
      getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
        getDisplayValues: () =>
          usedGrid
            .slice(row, row + (numRows ?? 1))
            .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
        getValues: () => [],
        getCellStyleData: () =>
          row === 1 && column === 0 ? { bd: { t: { s: 1, cl: { rgb: '#ff0000' } } } } : null,
      })) as PrintWorksheet['getRange'],
    }
    const html = assemblePrintHtml([
      layout(worksheet, payloadSetup({ printGridlines: true, printHeadings: true })),
    ]).html
    const cells = [...html.matchAll(/<td style="([^"]*)">/g)].map((match) => match[1] ?? '')
    // A1: no bottom gridline (A2 owns that edge), no top/left (heading strip owns them)
    expect(cells[0]).toContain('inset -0.75pt 0 0 #c0c0c0')
    expect(cells[0]).not.toContain('inset 0 -0.75pt')
    expect(cells[0]).not.toContain('inset 0 0.75pt')
    expect(cells[0]).not.toContain('inset 0.75pt 0')
    // A2 paints its red top and the shared bottom/right gridlines
    expect(cells[2]).toContain('inset 0 0.75pt 0 #ff0000')
    expect(cells[2]).toContain('inset 0 -0.75pt 0 #c0c0c0')
    // the heading strip owns the outer frame: top on the letter row, left on the number column
    expect(html).toContain('<th class="hd hd-top hd-left"></th>')
    expect(html).toContain('<th class="hd hd-top">A</th>')
    expect(html).toContain('<th class="hd hd-left">1</th>')
  })

  it('spans merged cells inside repeated title rows', () => {
    const worksheet: PrintWorksheet = {
      ...tallWorksheet(40),
      getLastColumn: () => 2,
      getMergedRanges: () => [
        { getRow: () => 0, getColumn: () => 0, getWidth: () => 3, getHeight: () => 1 },
      ],
    }
    const html = assemblePrintHtml([layout(worksheet, payloadSetup({ printTitles: '1:1' }))]).html
    expect(html).toContain('<thead><tr style="height:15pt"><td rowspan="1" colspan="3"')
    // the title row merge also spans when the print area starts below it
    const cropped = assemblePrintHtml([
      layout(worksheet, payloadSetup({ printTitles: '1:1', printAreas: ['A5:C10'] })),
    ]).html
    expect(cropped).toContain('<thead><tr style="height:15pt"><td rowspan="1" colspan="3"')
  })

  it('places floating visuals at their anchor and widens the used range to cover them', () => {
    const visual = {
      id: 'v1',
      fromRow: 1,
      fromColumn: 1,
      toRow: 6,
      toColumn: 4,
      offsetXPx: 10,
      offsetYPx: 4,
      widthPx: 400,
      heightPx: 200,
      html: '<div class="xlsx-print-visual">chart</div>',
    }
    const outside = { ...visual, id: 'v2', fromRow: 40, fromColumn: 30, toRow: 41, toColumn: 31 }
    const html = assemblePrintHtml([
      layout(fakeWorksheet(), payloadSetup({ printAreas: ['A1:B3'] }), new Map(), {
        visuals: [visual, outside],
        css: '.xlsx-chart { color: black; }',
      }),
    ]).html
    // column B starts after one 100px (75pt) column; row 2 after one 20px (15pt) row
    expect(html).toContain(
      '<div class="pv" style="left:82.5pt;top:18pt;width:300pt;height:150pt"><div style="width:400px;height:200px"><div class="xlsx-print-visual">chart</div></div></div>',
    )
    expect(html.match(/class="pv"/g)).toHaveLength(1)
    expect(html).toContain('<style>.xlsx-chart { color: black; }</style>')

    const widened = assemblePrintHtml([
      layout(fakeWorksheet(), payloadSetup({}), new Map(), { visuals: [visual], css: '' }),
    ]).html
    // A:B of data, but the chart reaches column E (index 4)
    expect(widened.match(/<col /g)).toHaveLength(5)
    expect(widened).not.toContain('<style></style>')
  })

  it('carries the page geometry and renders the header/footer inside the page', () => {
    const document = layout(
      fakeWorksheet(),
      payloadSetup({
        orientation: 'landscape',
        paperSize: 1,
        scale: 65,
        margins: { left: 0.98, right: 0.98, top: 0.98, bottom: 0.79, header: 0, footer: 0.51 },
        footer: { center: 'Seite &P von &N' },
      }),
    )
    expect(document.paper.landscape).toBe(true)
    expect(document.paper.pageSize).toBe('Letter')
    expect(document.scale).toBeCloseTo(0.65, 5)
    const html = assemblePrintHtml([document]).html
    expect(html).toContain('style="width:11in;height:8.5in"')
    expect(html).toContain('left:0.98in;top:0.98in;width:9.04in;height:6.73in')
    expect(html).toContain('zoom:0.65')
    expect(html).toContain('bottom:0.51in')
    expect(html).toContain('Seite 1 von 1')
    expect(html).toContain('@page p0 { size: 11in 8.5in; margin: 0; }')
  })

  it('numbers pages across the sheets of a job and keeps &N at the job total', () => {
    const first = layout(tallWorksheet(120), payloadSetup({ footer: { right: '&P/&N' } }))
    const second = layout(fakeWorksheet(), payloadSetup({ footer: { right: '&P/&N' } }))
    const job = assemblePrintHtml([first, second])
    const total = first.pages.length + 1
    expect(job.pages).toHaveLength(total)
    expect(job.html).toContain(`>1/${total}<`)
    expect(job.html).toContain(`>${total}/${total}<`)
    const subset = assemblePrintHtml([first, second], (page) => page === total)
    expect(subset.pages).toEqual([{ document: 1, page: 0 }])
    expect(subset.html).not.toContain(`>1/${total}<`)
  })

  it('shrinks to fit the width when fit-to-page is on', () => {
    // content 150pt wide (two 100px columns at 0.75), A4 printable ~493pt
    expect(layout(fakeWorksheet(), payloadSetup({ fitToPage: true, fitToWidth: 1 })).scale).toBe(1)
  })

  it('shrinks to fit the height when fitToHeight is set and ignores manual breaks', () => {
    // 200 rows of 15pt need 3000pt; A4 portrait leaves 733.5pt between 0.75in margins.
    const onePage = layout(
      tallWorksheet(200),
      payloadSetup({ fitToPage: true, fitToWidth: 0, fitToHeight: 1 }),
      new Map(),
      undefined,
      { breaks: { rowBreaks: [50, 100], colBreaks: [] } },
    )
    expect(onePage.scale).toBeLessThan(0.25)
    expect(onePage.scale).toBeGreaterThan(0.2)
    expect(onePage.pages).toHaveLength(1)
    const twoPages = layout(
      tallWorksheet(200),
      payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 2 }),
    )
    expect(twoPages.scale).toBeGreaterThan(onePage.scale)
    expect(twoPages.scale).toBeLessThan(0.5)
    expect(twoPages.pages).toHaveLength(2)
    // Width alone is satisfied at 100%; the height axis decides.
    expect(
      layout(tallWorksheet(200), payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 }))
        .scale,
    ).toBe(1)
    // The saved scale is ignored while fit-to-page is on.
    expect(
      layout(
        tallWorksheet(3),
        payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 1, scale: 50 }),
      ).scale,
    ).toBe(1)
  })

  it('uses first/even page header/footer variants only when active', () => {
    const plain = layout(tallWorksheet(120), payloadSetup({ header: { center: 'Odd' } }))
    expect(plain.headerFooter.first).toBeNull()
    expect(plain.headerFooter.even).toBeNull()

    const variants = layout(
      tallWorksheet(120),
      payloadSetup({
        header: { center: 'Odd' },
        footer: { center: '&P' },
        firstPage: { header: null, footer: { right: 'First' } },
        evenPages: { header: { left: 'Even' }, footer: null },
      }),
    )
    expect(variants.pages).toHaveLength(3)
    expect(
      variants.pages.map((_page, index) => pageVariant(variants.headerFooter, index, index + 1)),
    ).toEqual(['first', 'even', 'odd'])
    // differentFirst with a blank first header: page 1 prints no header.
    expect(variants.headerFooter.first?.header).toBeNull()
    expect(variants.headerFooter.even?.footer).toBeNull()
    const html = assemblePrintHtml([variants]).html
    expect(html).toContain('First')
    expect(html).toContain('Even')
    // A second sheet after an odd page count: its first page is job page 4
    // (even), its second job page 5 (odd).
    const second = layout(
      tallWorksheet(60),
      payloadSetup({ evenPages: { header: { left: 'Even' }, footer: null } }),
    )
    expect(pageVariant(second.headerFooter, 0, 4)).toBe('even')
    expect(pageVariant(second.headerFooter, 1, 5)).toBe('odd')
  })

  it('puts the &G picture of each variant into its own header', () => {
    const pictures = new Map([
      ['LH', { dataUrl: 'data:image/png;base64,ODD', widthPt: 300, heightPt: 30 }],
      ['CHFIRST', { dataUrl: 'data:image/png;base64,FIRST', widthPt: 100, heightPt: 50 }],
    ])
    const html = assemblePrintHtml([
      layout(
        tallWorksheet(120),
        payloadSetup({
          header: { left: '&G' },
          firstPage: { header: { center: '&G' }, footer: null },
        }),
        pictures,
      ),
    ]).html
    expect(html).toContain('src="data:image/png;base64,ODD"')
    expect(html).toContain('width:400px;height:40px')
    expect(html.indexOf('FIRST')).toBeLessThan(html.indexOf('ODD'))
  })

  it('scales the header/footer with the document unless scaleWithDoc is off', () => {
    const pictures = new Map([
      ['LH', { dataUrl: 'data:image/png;base64,LOGO', widthPt: 442.5, heightPt: 43.5 }],
    ])
    const scaled = assemblePrintHtml([
      layout(
        fakeWorksheet(),
        payloadSetup({ scale: 65, header: { left: '&G' }, footer: { center: 'Seite &P' } }),
        pictures,
      ),
    ]).html
    // 442.5pt * 0.65 = 287.6pt = 383.5px; 43.5pt * 0.65 = 28.3pt = 37.7px
    expect(scaled).toContain('width:383.5px;height:37.7px')
    expect(scaled).toContain('font-size:5.85pt')

    const fixed = assemblePrintHtml([
      layout(
        fakeWorksheet(),
        payloadSetup({
          scale: 65,
          header: { left: '&G' },
          footer: { center: 'Seite &P' },
          headerFooterScaleWithDoc: false,
        }),
        pictures,
      ),
    ]).html
    expect(fixed).toContain('width:590px;height:58px')
    expect(fixed).toContain('font-size:9pt')
  })

  it('sanitizes non-finite workbook dimensions instead of emitting NaNpt', () => {
    const hostile: PrintWorksheet = {
      ...fakeWorksheet(),
      getRowHeight: () => NaN,
      getColumnWidth: () => Infinity,
    }
    const payload = buildSheetPrintPayload(hostile, payloadSetup({}), 'Book.pdf', 'S1')
    expect(payload.html).toContain('<table>')
    expect(payload.html).not.toContain('NaN')
    expect(payload.html).not.toContain('Infinity')
    expect(payload.margins).toEqual({ top: 0, bottom: 0, left: 0, right: 0 })
    expect(payload.scale).toBe(1)
  })
})

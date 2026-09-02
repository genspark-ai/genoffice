import { describe, expect, it } from 'vitest'

import {
  buildSheetPrintPayload,
  renderHeaderFooterHtml,
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
    expect(printAreasFromFormula("'S'!$A$1:$B$2,'S'!$C:$D")).toEqual([])
  })

  it('returns [] when absent', () => {
    expect(printAreasFromFormula(undefined)).toEqual([])
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

describe('renderHeaderFooterHtml', () => {
  const now = new Date(2026, 0, 2, 3, 4, 5)

  it('turns &P/&N into live Chromium spans', () => {
    expect(renderHeaderFooterHtml('Seite &P von &N', 'Book', 'S1', now)).toBe(
      'Seite <span class="pageNumber"></span> von <span class="totalPages"></span>',
    )
  })

  it('resolves static codes and escapes markup', () => {
    const html = renderHeaderFooterHtml('&A <&F> && more', 'Bud<get', 'Sh&eet', now)
    expect(html).toBe('Sh&amp;eet &lt;Bud&lt;get&gt; &amp; more')
  })

  it('replaces &G with the section picture at its declared size', () => {
    const picture = { dataUrl: 'data:image/png;base64,AAAA', widthPt: 72, heightPt: 36 }
    expect(renderHeaderFooterHtml('&G Logo', 'Book', 'S1', now, picture)).toBe(
      '<img src="data:image/png;base64,AAAA" ' +
        'style="width:96px;height:48px;vertical-align:bottom"> Logo',
    )
  })

  it('prints nothing for &G when the slot has no picture', () => {
    expect(renderHeaderFooterHtml('&GTitle', 'Book', 'S1', now)).toBe('Title')
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

describe('buildSheetPrintPayload', () => {
  it('crops the layout to the print area', () => {
    const payload = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({ printAreas: ['A1:A2'] }),
      'Book.pdf',
      'S1',
    )
    expect(payload.html).toContain('A1')
    expect(payload.html).toContain('A2')
    expect(payload.html).not.toContain('B1')
    expect(payload.html).not.toContain('A3')
  })

  it('emits one page-broken table per print area', () => {
    const payload = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({ printAreas: ['A1:A1', 'B2:B2'] }),
      'Book.pdf',
      'S1',
    )
    expect(payload.html.match(/<table>/g)).toHaveLength(2)
    expect(payload.html).toContain('table + table { break-before: page; }')
  })

  it('carries the page geometry and header/footer templates', () => {
    const payload = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({
        orientation: 'landscape',
        paperSize: 1,
        scale: 65,
        margins: { left: 0.98, right: 0.98, top: 0.98, bottom: 0.79, header: 0, footer: 0.51 },
        footer: { center: 'Seite &P von &N' },
      }),
      'Book.pdf',
      'S1',
    )
    expect(payload.landscape).toBe(true)
    expect(payload.pageSize).toBe('Letter')
    expect(payload.scale).toBeCloseTo(0.65, 5)
    expect(payload.margins).toEqual({ top: 0.98, bottom: 0.79, left: 0.98, right: 0.98 })
    expect(payload.headerTemplate).toBeUndefined()
    expect(payload.footerTemplate).toContain('<span class="pageNumber"></span>')
    expect(payload.footerTemplate).toContain('<span class="totalPages"></span>')
    expect(payload.footerTemplate).toContain('padding-bottom:0.51in')
    // the template document is content-box: without an inline border-box the
    // padded 100%-wide row overflows the page and shifts/clips the sections
    expect(payload.footerTemplate).toContain('box-sizing:border-box')
  })

  it('shrinks to fit the width when fit-to-page is on', () => {
    // content 150pt wide (two 100px columns at 0.75), A4 printable ~493pt
    const payload = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({ fitToPage: true, fitToWidth: 1 }),
      'Book.pdf',
      'S1',
    )
    expect(payload.scale).toBe(1)
  })

  it('shrinks to fit the height when fitToHeight is set', () => {
    // 200 text rows: printed at ~15.75pt each (11pt line + padding) they
    // need ~3150pt; A4 portrait leaves 733.5pt between 0.75in margins.
    const onePage = buildSheetPrintPayload(
      tallWorksheet(200),
      payloadSetup({ fitToPage: true, fitToWidth: 0, fitToHeight: 1 }),
      'Book.pdf',
      'S1',
    )
    expect(onePage.scale).toBeLessThan(0.24)
    expect(onePage.scale).toBeGreaterThan(0.2)
    const twoPages = buildSheetPrintPayload(
      tallWorksheet(200),
      payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 2 }),
      'Book.pdf',
      'S1',
    )
    expect(twoPages.scale).toBeGreaterThan(onePage.scale)
    expect(twoPages.scale).toBeLessThan(0.47)
    // Width alone is satisfied at 100%; the height axis decides.
    expect(
      buildSheetPrintPayload(
        tallWorksheet(200),
        payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 }),
        'Book.pdf',
        'S1',
      ).scale,
    ).toBe(1)
    // The saved scale is ignored while fit-to-page is on.
    expect(
      buildSheetPrintPayload(
        tallWorksheet(3),
        payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 1, scale: 50 }),
        'Book.pdf',
        'S1',
      ).scale,
    ).toBe(1)
  })

  it('emits first/even page templates only for active variants', () => {
    const plain = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({ header: { center: 'Odd' } }),
      'Book.pdf',
      'S1',
    )
    expect(plain.firstPage).toBeUndefined()
    expect(plain.evenPages).toBeUndefined()

    const variants = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({
        header: { center: 'Odd' },
        footer: { center: '&P' },
        firstPage: { header: null, footer: { right: 'First' } },
        evenPages: { header: { left: 'Even' }, footer: null },
      }),
      'Book.pdf',
      'S1',
    )
    expect(variants.headerTemplate).toContain('Odd')
    expect(variants.footerTemplate).toContain('pageNumber')
    // differentFirst with a blank first header: page 1 prints no header.
    expect(variants.firstPage).toEqual({
      footerTemplate: expect.stringContaining('First') as string,
    })
    expect(variants.evenPages).toEqual({
      headerTemplate: expect.stringContaining('Even') as string,
    })
  })

  it('puts the &G picture of each variant into its own template', () => {
    const pictures = new Map([
      ['LH', { dataUrl: 'data:image/png;base64,ODD', widthPt: 300, heightPt: 30 }],
      ['CHFIRST', { dataUrl: 'data:image/png;base64,FIRST', widthPt: 100, heightPt: 50 }],
    ])
    const payload = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({
        header: { left: '&G' },
        firstPage: { header: { center: '&G' }, footer: null },
      }),
      'Book.pdf',
      'S1',
      pictures,
    )
    expect(payload.headerTemplate).toContain('src="data:image/png;base64,ODD"')
    expect(payload.headerTemplate).toContain('width:400px;height:40px')
    expect(payload.headerTemplate).not.toContain('FIRST')
    expect(payload.firstPage?.headerTemplate).toContain('src="data:image/png;base64,FIRST"')
    expect(payload.firstPage?.headerTemplate).not.toContain('ODD')
  })

  it('scales the header/footer with the document unless scaleWithDoc is off', () => {
    // phpss_issue.1767: scale 65, a 442.5pt x 43.5pt logo in the left header.
    const pictures = new Map([
      ['LH', { dataUrl: 'data:image/png;base64,LOGO', widthPt: 442.5, heightPt: 43.5 }],
    ])
    const scaled = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({ scale: 65, header: { left: '&G' }, footer: { center: 'Seite &P' } }),
      'Book.pdf',
      'S1',
      pictures,
    )
    // 442.5pt * 0.65 = 287.6pt = 383.5px; 43.5pt * 0.65 = 28.3pt = 37.7px
    expect(scaled.headerTemplate).toContain('width:383.5px;height:37.7px')
    expect(scaled.headerTemplate).toContain('font-size:5.85pt')
    expect(scaled.footerTemplate).toContain('font-size:5.85pt')

    const fixed = buildSheetPrintPayload(
      fakeWorksheet(),
      payloadSetup({
        scale: 65,
        header: { left: '&G' },
        footer: { center: 'Seite &P' },
        headerFooterScaleWithDoc: false,
      }),
      'Book.pdf',
      'S1',
      pictures,
    )
    expect(fixed.headerTemplate).toContain('width:590px;height:58px')
    expect(fixed.footerTemplate).toContain('font-size:9pt')
  })
})

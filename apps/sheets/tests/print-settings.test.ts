import { describe, expect, it } from 'vitest'

import { buildSheetPrintPayload, renderHeaderFooterHtml } from '../src/renderer/print-html'
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
    expect(decodeHeaderFooter('&C&"Broadway,Bold Italic"&12&KFF0000Big &BRed&B Title&G')).toEqual({
      center: 'Big Red Title',
    })
  })

  it('drops the picture/path codes but keeps the file code', () => {
    expect(decodeHeaderFooter('&C&P / &N&R&Z&F')).toEqual({
      center: '&P / &N',
      right: '&F',
    })
  })

  it('keeps escaped ampersands verbatim', () => {
    expect(decodeHeaderFooter('&LProfit && Loss')).toEqual({ left: 'Profit && Loss' })
  })

  it('returns null when everything strips away', () => {
    expect(decodeHeaderFooter('&L&G')).toBeNull()
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

  it('defaults fitToWidth to one page when the file only sets fitToPage', () => {
    const setup = resolveEffectivePageSetup({}, { fitToPage: true }, null)
    expect(setup.fitToPage).toBe(true)
    expect(setup.fitToWidth).toBe(1)
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
    fitToPage: false,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    printGridlines: false,
    printHeadings: false,
    printAreas: [],
    printTitles: null,
    header: null,
    footer: null,
    ...overrides,
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
})

import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import { hfHasPageField, hfWithoutPageMarks, makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { restingHfAreaVariant } from '../src/renderer/doc-state'

describe('page-number substitution (PAGE_MARK)', () => {
  it('fills the PAGE field position and keeps a literal # intact', () => {
    const value: HeaderFooter = {
      text: `[Course #] | Page ${PAGE_MARK}`,
      pageNumber: true,
      paras: [{ align: 'center', runs: [{ text: '[Course #] | Page ' }, { text: PAGE_MARK }] }],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 3, pageTotal: 7 })
    expect(el.textContent).toBe('[Course #] | Page 3')
  })

  it('legacy user-typed # (no PAGE_MARK anywhere) still substitutes', () => {
    const value: HeaderFooter = {
      text: '- # -',
      pageNumber: true,
      paras: [{ align: 'center', runs: [{ text: '- # -' }] }],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 2, pageTotal: 5 })
    expect(el.textContent).toBe('- 2 -')
  })

  it('substitutes every PAGE_MARK and the NUMPAGES total', () => {
    const value: HeaderFooter = {
      text: `${PAGE_MARK} / ${TOTAL_PAGES_MARK}`,
      pageNumber: true,
      paras: [{ runs: [{ text: `${PAGE_MARK} / ${TOTAL_PAGES_MARK}` }] }],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 4, pageTotal: 9 })
    expect(el.textContent).toBe('4 / 9')
  })

  it('pageNumber without any marker appends a PAGE_MARK run (legacy single line)', () => {
    const value: HeaderFooter = { text: 'Confidential', pageNumber: true }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 6, pageTotal: 8 })
    expect(el.textContent).toBe('Confidential 6')
  })
})

describe('restingHfAreaVariant (canvas area variant when no chip is picked)', () => {
  it('header area sits on page 1: first variant when titlePg is on', () => {
    expect(restingHfAreaVariant('header', { titlePg: true, evenOddHf: false, pageCount: 2 })).toBe(
      'first',
    )
    expect(restingHfAreaVariant('header', { titlePg: false, evenOddHf: false, pageCount: 2 })).toBe(
      'default',
    )
    expect(restingHfAreaVariant('header', { titlePg: false, evenOddHf: true, pageCount: 2 })).toBe(
      'default',
    )
  })

  it('footer area sits on the last page', () => {
    expect(restingHfAreaVariant('footer', { titlePg: true, evenOddHf: false, pageCount: 1 })).toBe(
      'first',
    )
    expect(restingHfAreaVariant('footer', { titlePg: true, evenOddHf: false, pageCount: 2 })).toBe(
      'default',
    )
    expect(
      restingHfAreaVariant('footer', {
        titlePg: false,
        evenOddHf: true,
        pageCount: 2,
        lastPageNo: 2,
      }),
    ).toBe('even')
    expect(
      restingHfAreaVariant('footer', {
        titlePg: false,
        evenOddHf: true,
        pageCount: 3,
        lastPageNo: 3,
      }),
    ).toBe('default')
  })

  it('parity follows the displayed last page number, not the physical count', () => {
    expect(
      restingHfAreaVariant('footer', {
        titlePg: false,
        evenOddHf: true,
        pageCount: 3,
        lastPageNo: 4,
      }),
    ).toBe('even')
    expect(
      restingHfAreaVariant('footer', {
        titlePg: false,
        evenOddHf: true,
        pageCount: 4,
        lastPageNo: 3,
      }),
    ).toBe('default')
  })
})

describe('hfWithoutPageMarks (ribbon Remove Page Numbers)', () => {
  it('strips only PAGE_MARK fields, keeping literal # and rich formatting', () => {
    const value: HeaderFooter = {
      text: `[Course #] | Page ${PAGE_MARK}`,
      pageNumber: true,
      paras: [
        { align: 'right', runs: [{ text: '[Course #] | Page ', bold: true }, { text: PAGE_MARK }] },
      ],
    }
    expect(hfWithoutPageMarks(value)).toEqual({
      text: '[Course #] | Page ',
      pageNumber: false,
      paras: [{ align: 'right', runs: [{ text: '[Course #] | Page ', bold: true }] }],
    })
  })

  it('legacy part (no sentinel): strips only the first user-typed #', () => {
    const value: HeaderFooter = {
      text: 'p# of #',
      pageNumber: true,
      paras: [{ align: 'center', runs: [{ text: 'p# of #' }] }],
    }
    expect(hfWithoutPageMarks(value)).toEqual({
      text: 'p of #',
      pageNumber: false,
      paras: [{ align: 'center', runs: [{ text: 'p of #' }] }],
    })
  })

  it('page-number-only footer removes as empty (no leftover paragraph shell)', () => {
    const value: HeaderFooter = {
      text: PAGE_MARK,
      pageNumber: true,
      paras: [{ align: 'center', runs: [{ text: PAGE_MARK }] }],
    }
    const next = hfWithoutPageMarks(value)
    expect(next).toEqual({ text: '', pageNumber: false })
    expect(next.paras).toBeUndefined()
  })

  it('multi-paragraph part: the dedicated page-number paragraph is dropped, blank lines stay', () => {
    const value: HeaderFooter = {
      text: `Confidential${PAGE_MARK}`,
      pageNumber: true,
      paras: [
        { align: 'left', runs: [{ text: 'Confidential', bold: true }] },
        { align: 'left', runs: [] },
        { align: 'center', runs: [{ text: PAGE_MARK }] },
      ],
    }
    expect(hfWithoutPageMarks(value)).toEqual({
      text: 'Confidential',
      pageNumber: false,
      paras: [
        { align: 'left', runs: [{ text: 'Confidential', bold: true }] },
        { align: 'left', runs: [] },
      ],
    })
  })

  it('layout-table rows survive removal untouched (display-only; save keeps the w:tbl bytes)', () => {
    const cells = [
      { runs: [{ text: 'ACME Corp' }], widthPct: 50 },
      { runs: [{ text: `Page ${PAGE_MARK}` }], widthPct: 50 },
    ]
    const value: HeaderFooter = {
      text: `ACME CorpPage ${PAGE_MARK}`,
      pageNumber: true,
      paras: [{ runs: [], cells }],
    }
    const next = hfWithoutPageMarks(value)
    expect(next.paras).toEqual([{ runs: [], cells }])
    expect(next.pageNumber).toBe(false)
    expect(next.text).toContain('ACME Corp')
  })

  it('mixed part: the dedicated page paragraph goes, the table row stays', () => {
    const row = { runs: [], cells: [{ runs: [{ text: 'Logo | Title' }] }] }
    const value: HeaderFooter = {
      text: `Logo | Title${PAGE_MARK}`,
      pageNumber: true,
      paras: [row, { align: 'center', runs: [{ text: PAGE_MARK }] }],
    }
    expect(hfWithoutPageMarks(value)).toEqual({
      text: 'Logo | Title',
      pageNumber: false,
      paras: [row],
    })
  })

  it('hfHasPageField ignores a literal # when no page number is set', () => {
    expect(hfHasPageField({ text: 'Item #5', pageNumber: false })).toBe(false)
    expect(hfHasPageField({ text: `p ${PAGE_MARK}`, pageNumber: true })).toBe(true)
    expect(hfHasPageField(null)).toBe(false)
  })
})

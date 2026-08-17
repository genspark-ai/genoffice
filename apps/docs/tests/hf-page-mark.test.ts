import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import { hfHasPageField, hfWithoutPageMarks, makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { hfFromPart, restingHfAreaVariant } from '../src/renderer/doc-state'

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

  it('paragraph w:jc lands on the element; no align means none (left via CSS, like Word)', () => {
    const value: HeaderFooter = {
      text: 'ab',
      pageNumber: false,
      paras: [{ align: 'right', runs: [{ text: 'a' }] }, { runs: [{ text: 'b' }] }],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[0].style.textAlign).toBe('right')
    expect(paras[1].style.textAlign).toBe('')
  })

  it('tabbed paragraph splits into positioned segments at its stops (three-column header)', () => {
    const value: HeaderFooter = {
      text: 'Left\tMid\tRight',
      pageNumber: false,
      paras: [
        {
          runs: [{ text: 'Left\tMid\t' }, { text: 'Right', bold: true }],
          tabStops: [
            { pos: 4513, val: 'center' },
            { pos: 9026, val: 'right' },
          ],
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const para = el.querySelector<HTMLElement>('.page-hf-para')!
    expect(para.classList.contains('page-hf-tabbed')).toBe(true)
    const segs = [...para.querySelectorAll<HTMLElement>('.page-hf-tabseg')]
    expect(segs).toHaveLength(2)
    expect(segs[0].textContent).toBe('Mid')
    expect(segs[0].classList.contains('page-hf-tabseg-center')).toBe(true)
    expect(segs[0].style.left).toBe(`${4513 / 15}px`)
    expect(segs[1].textContent).toBe('Right')
    expect(segs[1].classList.contains('page-hf-tabseg-right')).toBe(true)
    expect(segs[1].style.left).toBe(`${9026 / 15}px`)
    expect(segs[1].querySelector('span')?.style.fontWeight).toBe('600')
    // lead text stays inline
    expect(para.childNodes[0].textContent).toBe('Left')
  })

  it('ptab alignments win over stops and place at margin-relative percents', () => {
    const value: HeaderFooter = {
      text: 'L\tM\tR',
      pageNumber: false,
      paras: [
        {
          runs: [{ text: 'L\tM\tR' }],
          ptabAligns: ['center', 'right'],
          tabStops: [{ pos: 1000, val: 'left' }],
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const segs = [...el.querySelectorAll<HTMLElement>('.page-hf-tabseg')]
    expect(segs.map((s) => s.style.left)).toEqual(['50%', '100%'])
    expect(segs.map((s) => s.className)).toEqual([
      'page-hf-tabseg page-hf-tabseg-center',
      'page-hf-tabseg page-hf-tabseg-right',
    ])
  })

  it('tabbed paragraph keeps its shading and borders (mirrors the plain path)', () => {
    const value: HeaderFooter = {
      text: 'L\tR',
      pageNumber: false,
      paras: [
        {
          runs: [{ text: 'L\tR' }],
          shadingFill: 'EEEEEE',
          borders: 'b',
          borderLines: { b: { color: 'FF0000', szPt: 1 } },
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const para = el.querySelector<HTMLElement>('.page-hf-para.page-hf-tabbed')!
    expect(para.style.backgroundColor).not.toBe('')
    expect(para.style.borderBottom).toContain('rgb(255, 0, 0)')
  })

  it('oversized runs after a tab set the paragraph min-height (absolute segments add no flow height)', () => {
    const value: HeaderFooter = {
      text: '\tBig Title',
      pageNumber: false,
      paras: [{ runs: [{ text: '\tBig Title', sizeHalfPoints: 48 }] }],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const para = el.querySelector<HTMLElement>('.page-hf-tabbed')!
    expect(para.style.minHeight).toBe(`${24 * 1.3}pt`)
  })

  it('tabs without stops fall back to the implicit center/right header stops', () => {
    const value: HeaderFooter = {
      text: 'a\tb\tc',
      pageNumber: false,
      paras: [{ runs: [{ text: 'a\tb\tc' }] }],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 1, pageTotal: 1 })
    const segs = [...el.querySelectorAll<HTMLElement>('.page-hf-tabseg')]
    expect(segs.map((s) => s.style.left)).toEqual(['50%', '100%'])
    expect(segs.map((s) => s.textContent)).toEqual(['b', 'c'])
  })

  it('PAGE marks inside tab segments still substitute', () => {
    const value: HeaderFooter = {
      text: `Title\tPage ${PAGE_MARK} of ${TOTAL_PAGES_MARK}`,
      pageNumber: true,
      paras: [{ runs: [{ text: `Title\tPage ${PAGE_MARK} of ${TOTAL_PAGES_MARK}` }] }],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 3, pageTotal: 14 })
    expect(el.querySelector('.page-hf-tabseg')?.textContent).toBe('Page 3 of 14')
  })

  it('image strip follows the containing paragraph alignment (POI headerPic: jc=right)', () => {
    const value: HeaderFooter = { text: '', pageNumber: false, paras: [{ runs: [] }] }
    const el = makeGapHfEl({
      kind: 'header',
      value,
      images: [{ dataUrl: 'data:image/png;base64,', widthPx: 10, heightPx: 10, align: 'right' }],
      pageNo: 1,
      pageTotal: 1,
    })
    const wrap = el.querySelector<HTMLElement>('.page-hf-images')
    expect(wrap?.style.justifyContent).toBe('flex-end')
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
      { paras: [[{ text: 'ACME Corp' }]], widthPct: 50 },
      { paras: [[{ text: `Page ${PAGE_MARK}` }]], widthPct: 50 },
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
    const row = { runs: [], cells: [{ paras: [[{ text: 'Logo | Title' }]] }] }
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

describe('hfFromPart emptiness', () => {
  it('keeps an image-only part (logo footer) instead of dropping it', () => {
    const part = {
      text: '',
      hasPageNumber: false,
      paras: [],
      images: [{ dataUrl: 'data:image/png;base64,x' }],
    }
    expect(hfFromPart(part)).toEqual({ text: '', pageNumber: false, paras: undefined })
    expect(hfFromPart({ text: '', hasPageNumber: false, paras: [] })).toBeNull()
    expect(hfFromPart(null)).toBeNull()
  })
})

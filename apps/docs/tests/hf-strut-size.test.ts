/**
 * Header/footer strips size each paragraph's lines by its own runs (Word),
 * shrink-only: an all-8pt header measured at the 10.5pt strip strut pushed the
 * body top ~1px and cost a two-column sample its 42nd grid row on every page;
 * a strip-wide max gave 8pt letterhead lines an 11pt title's strut.
 */
import { describe, expect, it } from 'vitest'
import { hfTabSegments, makeGapHfEl } from '../src/renderer/editor/hf-dom'
import {
  fromHfParagraph,
  hfLineFactorNumber,
  hfParaLineHeightCss,
  layoutParaOfRuns,
  paraLineFactor,
  toStyleRecord,
} from '../src/renderer/editor/para-layout-spec'

const p = (runs: object[]) => ({ runs }) as never
/** inline style of a strip paragraph made of these runs */
const strip = (para: object) => toStyleRecord(fromHfParagraph(para as never))
const strut = (pt: number) => ({ '--doc-strut': `${pt}pt`, fontSize: 'var(--doc-strut)' })

describe('strip paragraph strut size', () => {
  it('smallest declared size carries the strut; larger runs lay their own boxes', () => {
    // Word probe 2026-10-29: 8pt br-lines under an 11pt first line pitch 8 x 1.15
    expect(strip(p([{ text: '\u5b66\u4f1a\u8a8c', sizeHalfPoints: 16 }]))).toMatchObject(strut(8))
    expect(
      strip(
        p([
          { text: 'a', sizeHalfPoints: 16 },
          { text: 'b', sizeHalfPoints: 18 },
        ]),
      ),
    ).toMatchObject(strut(8))
  })

  it('a text run without a size inherits the strip base through .doc-strut-mixed', () => {
    const box = fromHfParagraph(p([{ text: 'a', sizeHalfPoints: 16 }, { text: 'b' }]))
    expect(box.classes).toContain('doc-strut-mixed')
    expect(toStyleRecord(box)['--doc-strut']).toBe('min(8pt,var(--doc-base-fs,8pt))')
  })

  it('whitespace-only runs never size a line; a space-only paragraph is sized by its mark', () => {
    expect(strip(p([{ text: ' ' }, { text: 'x', sizeHalfPoints: 16 }]))).toMatchObject(strut(8))
    expect(
      strip({ runs: [{ text: ' ', sizeHalfPoints: 24 }], emptyRunSizeHalfPoints: 16 }),
    ).toMatchObject({ fontSize: '8pt' })
  })

  it('an empty paragraph lays a Latin line at the strip size', () => {
    expect(strip(p([]))).toEqual({
      direction: 'ltr',
      '--doc-line-factor': 'var(--doc-line-factor-latin,1.2)',
      lineHeight: 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    })
  })
})

describe('makeGapHfEl per-paragraph strut', () => {
  it('each paragraph shrinks to its own declared run size, not the strip-wide max', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [
          p([{ text: 'LETTERHEAD', sizeHalfPoints: 22 }]),
          p([{ text: 'street 1', sizeHalfPoints: 16 }]),
          p([{ text: 'plain' }]),
        ],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    expect(el.style.fontSize).toBe('')
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[0].style.getPropertyValue('--doc-strut')).toBe('11pt')
    expect(paras[1].style.getPropertyValue('--doc-strut')).toBe('8pt')
    expect(paras[2].style.fontSize).toBe('')
  })

  it('layout-table cell paragraphs take their own run size', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [
          {
            runs: [],
            cells: [
              { paras: [[{ text: 'a', sizeHalfPoints: 16 }], [{ text: 'b', sizeHalfPoints: 20 }]] },
            ],
          },
        ],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const cellParas = el.querySelectorAll<HTMLElement>('.page-hf-cell-para')
    expect(cellParas[0].style.getPropertyValue('--doc-strut')).toBe('8pt')
    expect(cellParas[1].style.getPropertyValue('--doc-strut')).toBe('10pt')
  })
})

describe('strip paragraph line spacing', () => {
  it('maps the style-resolved w:spacing to a plain factor multiple or a fixed height', () => {
    expect(hfParaLineHeightCss({ lineRule: 'auto', lineRawTwips: 240, lineSpacing: 1 })).toBe(
      'calc(var(--doc-line-factor,1.2) * 1)',
    )
    expect(hfParaLineHeightCss({ lineRule: 'auto', lineRawTwips: 324 })).toBe(
      'calc(var(--doc-line-factor,1.2) * 1.35)',
    )
    expect(hfParaLineHeightCss({ lineRule: 'exact', lineRawTwips: 360 })).toBe('18.0pt')
    expect(hfParaLineHeightCss({})).toBe(null)
  })

  it('sets the paragraph line-height so the probe measures a single 8pt line', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [
          {
            lineRule: 'auto',
            lineRawTwips: 240,
            lineSpacing: 1,
            runs: [{ text: 'Header', sizeHalfPoints: 16 }],
          },
          { runs: [{ text: 'Second', sizeHalfPoints: 16 }] },
        ],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[0].style.lineHeight).toBe('calc(var(--doc-line-factor,1.2) * 1)')
    expect(paras[1].style.lineHeight).toBe(
      'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    )
  })
})

describe('makeGapHfEl blank paragraph size', () => {
  it('sizes a run-less paragraph line by its mark / style size', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [{ runs: [] }, { runs: [], emptyRunSizeHalfPoints: 24 }] as never,
      },
      pageNo: 1,
      pageTotal: 1,
    })
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[0].style.fontSize).toBe('')
    expect(paras[1].style.fontSize).toBe('12pt')
  })
})

describe('strip paragraph line factor follows the runs’ face', () => {
  it('a Times New Roman 8pt line is 8 x 1.15, not the body font factor', () => {
    const runs = [{ text: 'NPSN : 2056', sizeHalfPoints: 16, fontAscii: 'Times New Roman' }]
    expect(strip(p(runs))).toMatchObject({
      '--doc-line-factor': '1.15',
      lineHeight: 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
      ...strut(8),
    })
    expect(strip(p(runs)).fontFamily).toMatch(/^'Times New Roman'/)
    expect(hfLineFactorNumber(runs as never)).toBe(1.15)
  })

  it('a footer table line keeps the inherited document multiple (docDefaults line=276)', () => {
    const el = makeGapHfEl({
      kind: 'footer',
      value: {
        text: '',
        paras: [
          {
            runs: [],
            cells: [
              {
                paras: [[{ text: 'Chairperson', sizeHalfPoints: 15, fontAscii: 'Arial' }]],
                paraProps: [{ spaceAfter: 0 }],
              },
            ],
          },
        ],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const cell = el.querySelector<HTMLElement>('.page-hf-cell-para')!
    expect(cell.style.getPropertyValue('--doc-line-factor')).toBe('1.15')
    expect(cell.style.lineHeight).toBe(
      'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    )
  })

  it('runs without a declared face take the inherited face factor at their size', () => {
    // Word probe 2026-10-29: an inherited-face 8pt strip line pitches 8 x 1.22
    expect(strip(p([{ text: 'plain', sizeHalfPoints: 16 }]))).toEqual({
      direction: 'ltr',
      '--doc-line-factor': 'var(--doc-line-factor-latin,1.2)',
      lineHeight: 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
      ...strut(8),
    })
    expect(
      hfLineFactorNumber([{ text: 'a', fontAscii: 'Times New Roman' }, { text: 'b' }] as never),
    ).toBe(null)
  })

  it('makeGapHfEl sets the factor per paragraph; a w:spacing multiple reads the same var', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [
          p([{ text: 'Alamat', sizeHalfPoints: 16, fontAscii: 'Times New Roman' }]),
          {
            lineRule: 'auto',
            lineRawTwips: 360,
            lineSpacing: 1.5,
            runs: [{ text: 'Title', sizeHalfPoints: 22, fontAscii: 'Cambria' }],
          },
          p([{ text: 'plain' }]),
        ],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[0].style.getPropertyValue('--doc-line-factor')).toBe('1.15')
    expect(paras[0].style.lineHeight).toBe(
      'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    )
    expect(paras[1].style.getPropertyValue('--doc-line-factor')).toBe('1.172')
    expect(paras[1].style.lineHeight).toBe('calc(var(--doc-line-factor,1.2) * 1.5)')
    expect(paras[2].style.getPropertyValue('--doc-line-factor')).toBe(
      'var(--doc-line-factor-latin,1.2)',
    )
    expect(paras[2].style.lineHeight).toBe(
      'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    )
  })

  it('a tabbed line floors at run size x face factor instead of the 1.3 guess', () => {
    const tabbed = hfTabSegments({
      runs: [
        { text: 'Lead\t', sizeHalfPoints: 16, fontAscii: 'Times New Roman' },
        { text: 'Right', sizeHalfPoints: 16, fontAscii: 'Times New Roman' },
      ],
      tabStops: [{ pos: 9000, align: 'right' }],
    } as never)
    expect(tabbed?.minHeightPt).toBeCloseTo(8 * 1.15, 5)
    expect(
      hfTabSegments({
        runs: [
          { text: 'Lead\t', sizeHalfPoints: 16 },
          { text: 'Right', sizeHalfPoints: 16 },
        ],
        tabStops: [{ pos: 9000, align: 'right' }],
      } as never)?.minHeightPt,
    ).toBeCloseTo(8 * 1.3, 5)
  })

  it('a CJK run after a tab floors at the East Asian face factor, not the Latin one', () => {
    const cjk = {
      text: '\u7b2c\u4e00\u7ae0',
      sizeHalfPoints: 16,
      font: 'SimSun',
      fontAscii: 'Times New Roman',
    }
    const tabbed = hfTabSegments({
      runs: [{ text: '\t', sizeHalfPoints: 16, fontAscii: 'Times New Roman' }, cjk],
      tabStops: [{ pos: 9000, align: 'right' }],
    } as never)
    const ea = paraLineFactor(layoutParaOfRuns([cjk] as never, cjk.text, null))
    expect(Number(ea)).toBeGreaterThan(1.15)
    expect(tabbed?.minHeightPt).toBeCloseTo(8 * Number(ea), 5)
    expect(hfLineFactorNumber([cjk] as never)).toBe(Number(ea))
    expect(
      hfLineFactorNumber([{ ...cjk, font: undefined, fontAscii: 'Times New Roman' }] as never),
    ).toBe(null)
  })

  it('a mixed-script run with an East Asian face but no Latin face keeps the EA factor', () => {
    const run = { text: '\u7b2c\u4e00\u7ae0 Chapter', sizeHalfPoints: 16, font: 'Microsoft YaHei' }
    const ea = Number(
      paraLineFactor(
        layoutParaOfRuns([{ ...run, text: '\u7b2c\u4e00\u7ae0' }] as never, '\u7b2c', null),
      ),
    )
    expect(ea).toBeGreaterThan(1.3)
    expect(hfLineFactorNumber([run] as never)).toBe(ea)
  })
})

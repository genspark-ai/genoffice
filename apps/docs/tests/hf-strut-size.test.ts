/**
 * Header/footer strips size each paragraph's lines by its own runs (Word),
 * shrink-only: an all-8pt header measured at the 10.5pt strip strut pushed the
 * body top ~1px and cost a two-column sample its 42nd grid row on every page;
 * a strip-wide max gave 8pt letterhead lines an 11pt title's strut.
 */
import { describe, expect, it } from 'vitest'
import {
  hfParaLineFactorStyle,
  hfParaLineHeightCss,
  hfParaStrutFontSize,
  hfRunsFaceFactor,
  hfRunsStrutPt,
  hfTabSegments,
  makeGapHfEl,
} from '../src/renderer/editor/hf-dom'
import { runsLineFactor } from '../src/renderer/line-metrics'

const p = (runs: object[]) => ({ runs }) as never

describe('hfRunsStrutPt', () => {
  it('largest declared size when every text run declares one', () => {
    expect(hfRunsStrutPt([{ text: '\u5b66\u4f1a\u8a8c', sizeHalfPoints: 16 }] as never)).toBe(8)
    expect(
      hfRunsStrutPt([
        { text: 'a', sizeHalfPoints: 16 },
        { text: 'b', sizeHalfPoints: 18 },
      ] as never),
    ).toBe(9)
  })

  it('a text run without a size inherits the strip base (null)', () => {
    expect(hfRunsStrutPt([{ text: 'a', sizeHalfPoints: 16 }, { text: 'b' }] as never)).toBe(null)
  })

  it('whitespace-only runs never veto and only size empty paragraphs', () => {
    expect(hfRunsStrutPt([{ text: ' ' }, { text: 'x', sizeHalfPoints: 16 }] as never)).toBe(8)
    expect(hfRunsStrutPt([{ text: ' ', sizeHalfPoints: 24 }] as never)).toBe(12)
    expect(
      hfRunsStrutPt([
        { text: ' ', sizeHalfPoints: 24 },
        { text: 'x', sizeHalfPoints: 16 },
      ] as never),
    ).toBe(8)
  })

  it('empty runs yield null', () => {
    expect(hfRunsStrutPt([])).toBe(null)
    expect(hfParaStrutFontSize([])).toBe(null)
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
    expect(paras[0].style.fontSize).toBe('min(11pt, var(--hf-default-fs, 10.5pt))')
    expect(paras[1].style.fontSize).toBe('min(8pt, var(--hf-default-fs, 10.5pt))')
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
    expect(cellParas[0].style.fontSize).toBe('min(8pt, var(--hf-default-fs, 10.5pt))')
    expect(cellParas[1].style.fontSize).toBe('min(10pt, var(--hf-default-fs, 10.5pt))')
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
    expect(paras[1].style.lineHeight).toBe('')
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
    expect(hfParaLineFactorStyle(runs as never)).toEqual({
      '--doc-line-factor': '1.15',
      lineHeight: 'calc(var(--doc-line-factor, 1.2) * var(--doc-line-mult, 1))',
    })
    expect(hfRunsFaceFactor(runs as never)).toBe(1.15)
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

  it('runs without a declared face keep the inherited strip factor', () => {
    expect(hfParaLineFactorStyle([{ text: 'plain', sizeHalfPoints: 16 }] as never)).toEqual({})
    expect(
      hfRunsFaceFactor([{ text: 'a', fontAscii: 'Times New Roman' }, { text: 'b' }] as never),
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
    expect(paras[2].style.getPropertyValue('--doc-line-factor')).toBe('')
    expect(paras[2].style.lineHeight).toBe('')
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
    const ea = runsLineFactor([cjk] as never, cjk.text)
    expect(Number(ea)).toBeGreaterThan(1.15)
    expect(tabbed?.minHeightPt).toBeCloseTo(8 * Number(ea), 5)
    expect(hfRunsFaceFactor([cjk] as never)).toBe(Number(ea))
    expect(
      hfRunsFaceFactor([{ ...cjk, font: undefined, fontAscii: 'Times New Roman' }] as never),
    ).toBe(null)
  })

  it('a mixed-script run with an East Asian face but no Latin face keeps the EA factor', () => {
    const run = { text: '\u7b2c\u4e00\u7ae0 Chapter', sizeHalfPoints: 16, font: 'Microsoft YaHei' }
    const ea = Number(runsLineFactor([{ ...run, text: '\u7b2c\u4e00\u7ae0' }] as never, '\u7b2c'))
    expect(ea).toBeGreaterThan(1.3)
    expect(hfRunsFaceFactor([run] as never)).toBe(ea)
  })
})

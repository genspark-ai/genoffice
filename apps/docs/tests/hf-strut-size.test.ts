/**
 * Header/footer strips size their lines by their runs (Word), shrink-only:
 * an all-8pt header measured at the 10.5pt strip strut pushed the body top
 * ~1px and cost every two-column page of SAS prod_043 its 42nd grid row.
 */
import { describe, expect, it } from 'vitest'
import { hfDeclaredStrutPt, hfParaLineHeightCss, makeGapHfEl } from '../src/renderer/editor/hf-dom'

const p = (runs: object[]) => ({ runs }) as never

describe('hfDeclaredStrutPt', () => {
  it('largest declared size when every text run declares one', () => {
    expect(hfDeclaredStrutPt([p([{ text: '学会誌', sizeHalfPoints: 16 }])])).toBe(8)
    expect(
      hfDeclaredStrutPt([
        p([
          { text: 'a', sizeHalfPoints: 16 },
          { text: 'b', sizeHalfPoints: 18 },
        ]),
      ]),
    ).toBe(9)
  })

  it('a text run without a size inherits the strip base (null)', () => {
    expect(hfDeclaredStrutPt([p([{ text: 'a', sizeHalfPoints: 16 }, { text: 'b' }])])).toBe(null)
  })

  it('whitespace-only runs never veto and only size empty paragraphs', () => {
    expect(hfDeclaredStrutPt([p([{ text: ' ' }, { text: 'x', sizeHalfPoints: 16 }])])).toBe(8)
    expect(hfDeclaredStrutPt([p([{ text: ' ', sizeHalfPoints: 24 }])])).toBe(12)
    expect(
      hfDeclaredStrutPt([
        p([
          { text: ' ', sizeHalfPoints: 24 },
          { text: 'x', sizeHalfPoints: 16 },
        ]),
      ]),
    ).toBe(8)
  })

  it('empty part yields null', () => {
    expect(hfDeclaredStrutPt([])).toBe(null)
    expect(hfDeclaredStrutPt([p([])])).toBe(null)
  })
})

describe('makeGapHfEl strip strut', () => {
  it('shrinks the strip font to the declared run size (shrink-only via min())', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: { text: '', paras: [p([{ text: '学会誌', sizeHalfPoints: 16 }])] } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    expect(el.style.fontSize).toBe('min(8pt, var(--hf-default-fs, 10.5pt))')
  })

  it('leaves the strip base when a run inherits its size', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: { text: '', paras: [p([{ text: '学会誌' }])] } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    expect(el.style.fontSize).toBe('')
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

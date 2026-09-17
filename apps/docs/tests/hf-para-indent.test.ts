import { describe, expect, it } from 'vitest'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

// header text indented beside a left logo (w:ind left=1134 right=119): Word
// narrows the text block; drawn from the margin it overlapped the logo and
// wrapped one line short
const para = (extra: object) =>
  ({ runs: [{ text: 'PROGRAM '.repeat(12).trim(), sizeHalfPoints: 24 }], ...extra }) as never

describe('header paragraph indent', () => {
  it('the strip paragraph indents by w:ind left/right (margins, clear of border padding)', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [para({ indentLeft: 1134, indentRight: 119, indentFirstLine: 300 })],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const p = el.querySelector<HTMLElement>('.page-hf-para')!
    expect(p.style.marginInlineStart).toBe('75.6px')
    expect(p.style.marginInlineEnd).toBe('7.9px')
    expect(p.style.textIndent).toBe('20px')
  })

  it('a bordered indented paragraph keeps both the indent and the border padding', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [para({ indentLeft: 1134, borders: 'lb', borderLines: { l: { szPt: 0.5 } } })],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const p = el.querySelector<HTMLElement>('.page-hf-para')!
    expect(p.style.marginInlineStart).toBe('75.6px')
    expect(p.style.paddingLeft).toBe('4px')
  })

  it('a tabbed paragraph keeps its column-relative tab layout (no indent margins)', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: {
        text: '',
        paras: [{ runs: [{ text: 'left\tright' }], indentLeft: 1134, indentRight: 119 }],
      } as never,
      pageNo: 1,
      pageTotal: 1,
    })
    const p = el.querySelector<HTMLElement>('.page-hf-para')!
    expect(p.style.marginInlineStart).toBe('')
    expect(p.style.marginInlineEnd).toBe('')
  })

  it('the height estimate wraps the indented paragraph in the narrower width', () => {
    const plain = estimateHfHeight({ text: '', paras: [para({})] }, 300)
    const indented = estimateHfHeight(
      { text: '', paras: [para({ indentLeft: 1134, indentRight: 119 })] },
      300,
    )
    expect(indented).toBeGreaterThan(plain)
    expect(indented).toBe(
      estimateHfHeight({ text: '', paras: [para({})] }, 300 - (1134 + 119) / 15),
    )
  })
})

// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { draftAutoBreaks, syncDraftBreaks } from '../src/renderer/editor/draft-view'
import type { PageCut, PageLead } from '../src/renderer/pagination-cuts'

const cut = (start: number, lead: PageLead): PageCut => ({ start, lead }) as PageCut
const block = (i: number): PageLead => ({ kind: 'block', block: i })

describe('draft view page breaks', () => {
  it('an automatic break gets a rule; the first page and a trailing blank page do not', () => {
    const cuts = [cut(0, block(0)), cut(600, block(3)), cut(1200, { kind: 'tail', block: -1 })]
    expect(draftAutoBreaks(cuts, [{}, {}, {}, {}])).toEqual([{ page: 1, block: 3, y: 600 }])
  })

  it('a break inside a paragraph or a table is automatic and placed by its flow position', () => {
    const cuts = [
      cut(0, block(0)),
      cut(500, { kind: 'line', block: 1, lineOffset: 40 }),
      cut(1000, { kind: 'row', block: 2, row: 4 }),
    ]
    expect(draftAutoBreaks(cuts, [{}, {}, {}])).toEqual([
      { page: 1, y: 500 },
      { page: 2, y: 1000 },
    ])
  })

  it('a break that shows its own mark gets no rule; one inherited from a style does', () => {
    const cuts = [cut(0, block(0)), cut(300, block(1)), cut(600, block(2)), cut(900, block(3))]
    const blocks = [{ section: 0 }, { section: 0 }, { section: 1 }, { section: 1 }]
    const marks = {
      startsWithMark: (i: number) => i === 1,
      endsInPageBreak: (i: number) => i === 2,
    }
    expect(draftAutoBreaks(cuts, blocks, marks)).toEqual([])
    expect(draftAutoBreaks(cuts, blocks).map((b) => b.page)).toEqual([1, 3])
  })

  it('a blank page and the page after it share one rule', () => {
    const cuts = [cut(0, block(0)), cut(400, block(2)), cut(400, block(2))]
    expect(draftAutoBreaks(cuts, [{}, {}, {}])).toHaveLength(1)
  })

  it('the rules follow the breaks and go when there are none', () => {
    const layer = document.createElement('div')
    syncDraftBreaks(layer, [120, 840.26])
    const tops = () =>
      Array.from(layer.querySelectorAll<HTMLElement>(':scope > .draft-break')).map(
        (el) => el.style.top,
      )
    expect(tops()).toEqual(['120px', '840.3px'])
    syncDraftBreaks(layer, [300])
    expect(tops()).toEqual(['300px'])
    syncDraftBreaks(layer, [])
    expect(layer.childElementCount).toBe(0)
  })
})

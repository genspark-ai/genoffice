import { describe, expect, it } from 'vitest'
import type { CommentInfo } from '@genoffice/docx-engine'
import { measureCommentSpots } from '../src/renderer/components/PaginationPreview'

/** stub layout: jsdom has no layout engine, so anchors declare their rects */
type Rect = { top: number; bottom: number; left: number; right: number; height: number }

function stubRects(el: HTMLElement, rects: Rect[]): void {
  el.getClientRects = () => rects as unknown as DOMRectList
}

const rect = (top: number, left: number, w = 40, h = 16): Rect => ({
  top,
  bottom: top + h,
  left,
  right: left + w,
  height: h,
})

const comment = (id: string, extra: Partial<CommentInfo> = {}): CommentInfo => ({
  id,
  author: 'A',
  text: `text ${id}`,
  ...extra,
})

function span(id: string, rects: Rect[]): HTMLElement {
  const el = document.createElement('span')
  el.className = 'doc-comment'
  el.dataset.commentIds = id
  stubRects(el, rects)
  return el
}

function para(idx: number, rects: Rect[]): HTMLElement {
  const el = document.createElement('p')
  el.dataset.idx = String(idx)
  stubRects(el, rects)
  return el
}

describe('measureCommentSpots', () => {
  it('numbers and stacks by document order, not by geometric (top, endX) order', () => {
    // two anchors on the same line: the doc-first range ends further right
    // (on the next line, at a small X) — a geometric sort would invert them
    const pm = document.createElement('div')
    const p = document.createElement('p')
    const first = span('7', [rect(100, 50), rect(116, 10, 30)])
    const second = span('8', [rect(100, 200)])
    p.append(first, second)
    pm.append(p)
    stubRects(pm, [rect(0, 0)])

    const spots = measureCommentSpots(pm, [comment('7'), comment('8')], [], 0, 1, 0)
    expect(spots.map((s) => s.id)).toEqual(['7', '8'])
    expect(spots.map((s) => s.no)).toEqual([1, 2])
    // the leader still leaves the range end (last rect of the last span)
    expect(spots[0].endX).toBe(40)
    expect(spots[0].endY).toBe(131)
  })

  it('anchors a cross-paragraph range to the blocks carrying its markers', () => {
    const pm = document.createElement('div')
    const spanEl = span('1', [rect(300, 20)])
    const p1 = para(4, [rect(40, 0, 400)])
    const p2 = para(6, [rect(80, 0, 400), rect(96, 0, 200)])
    const pMid = document.createElement('p')
    pMid.append(spanEl)
    pm.append(p1, p2, pMid)
    const blocks = [
      { docxIndex: 4, originalXml: '<w:p><w:commentRangeStart w:id="0"/><w:r/></w:p>' },
      { docxIndex: 6, originalXml: '<w:p><w:commentRangeEnd w:id="0"/><w:r/></w:p>' },
    ]

    const spots = measureCommentSpots(pm, [comment('0'), comment('1')], blocks, 0, 1, 0)
    expect(spots.map((s) => s.id)).toEqual(['0', '1'])
    expect(spots.map((s) => s.no)).toEqual([1, 2])
    // balloon top from the start block, leader from the end block's last line
    expect(spots[0].top).toBe(40)
    expect(spots[0].endX).toBe(200)
    expect(spots[0].endY).toBe(111)
  })

  it('suppresses balloons for ranges touching a table but keeps their numbers', () => {
    const pm = document.createElement('div')
    const tbl = document.createElement('div')
    tbl.dataset.idx = '2'
    stubRects(tbl, [rect(10, 0, 400, 100)])
    const p = document.createElement('p')
    const late = span('9', [rect(200, 30)])
    p.append(late)
    pm.append(tbl, p)
    const blocks = [
      {
        docxIndex: 2,
        originalXml:
          '<w:tbl><w:tr><w:tc><w:p><w:commentRangeStart w:id="5"/></w:p></w:tc>' +
          '<w:tc><w:p><w:commentRangeEnd w:id="5"/></w:p></w:tc></w:tr></w:tbl>',
      },
    ]

    const spots = measureCommentSpots(pm, [comment('5'), comment('9')], blocks, 0, 1, 0)
    // the cell-crossing thread prints no balloon (Word suppresses it) but the
    // next thread keeps Word's document-order number
    expect(spots.map((s) => s.id)).toEqual(['9'])
    expect(spots[0].no).toBe(2)
  })

  it('suppresses a table wrapped in an SDT (originalXml does not start with w:tbl)', () => {
    const pm = document.createElement('div')
    const tbl = document.createElement('div')
    tbl.dataset.idx = '0'
    stubRects(tbl, [rect(10, 0, 400, 100)])
    pm.append(tbl)
    const blocks = [
      {
        docxIndex: 0,
        originalXml:
          '<w:sdt><w:sdtPr/><w:sdtContent><w:tbl><w:tr><w:tc><w:p>' +
          '<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>' +
          '</w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>',
      },
    ]
    expect(measureCommentSpots(pm, [comment('1')], blocks, 0, 1, 0)).toEqual([])
  })

  it('orders same-block fallback threads by their marker offset in the XML', () => {
    const pm = document.createElement('div')
    const p1 = para(0, [rect(40, 0, 400)])
    const p2 = para(1, [rect(120, 0, 400)])
    pm.append(p1, p2)
    // both ranges start in block 0; "b" starts first in the XML but comes
    // later in the comments array (comments.xml order is arbitrary)
    const blocks = [
      {
        docxIndex: 0,
        originalXml:
          '<w:p><w:commentRangeStart w:id="b"/><w:r/><w:commentRangeStart w:id="a"/></w:p>',
      },
      {
        docxIndex: 1,
        originalXml: '<w:p><w:commentRangeEnd w:id="a"/><w:commentRangeEnd w:id="b"/><w:r/></w:p>',
      },
    ]
    const spots = measureCommentSpots(pm, [comment('a'), comment('b')], blocks, 0, 1, 0)
    expect(spots.map((s) => s.id)).toEqual(['b', 'a'])
    expect(spots.map((s) => s.no)).toEqual([1, 2])
  })

  it('interleaves a session-added thread among same-block file threads by span order', () => {
    const pm = document.createElement('div')
    const p = para(0, [rect(40, 0, 400)])
    const s1 = span('f1', [rect(40, 10)])
    const sNew = span('new', [rect(40, 100)])
    const s2 = span('f2', [rect(40, 200)])
    p.append(s1, sNew, s2)
    pm.append(p)
    // only the file threads have parsed markers; the session thread's span
    // sits between them in the DOM and must take the middle number
    const blocks = [
      {
        docxIndex: 0,
        originalXml:
          '<w:p><w:commentRangeStart w:id="f1"/><w:commentRangeEnd w:id="f1"/>' +
          '<w:commentRangeStart w:id="f2"/><w:commentRangeEnd w:id="f2"/></w:p>',
      },
    ]
    const spots = measureCommentSpots(
      pm,
      [comment('f2'), comment('new'), comment('f1')],
      blocks,
      0,
      1,
      0,
    )
    expect(spots.map((s) => s.id)).toEqual(['f1', 'new', 'f2'])
    expect(spots.map((s) => s.no)).toEqual([1, 2, 3])
  })

  it('numbers a session thread after a file thread sharing its first span', () => {
    const pm = document.createElement('div')
    const p = para(0, [rect(40, 0, 400)])
    // overlapping ranges: one span carries both ids, only "f" has a marker
    const shared = span('f new', [rect(40, 10)])
    p.append(shared)
    pm.append(p)
    const blocks = [
      {
        docxIndex: 0,
        originalXml: '<w:p><w:commentRangeStart w:id="f"/><w:commentRangeEnd w:id="f"/></w:p>',
      },
    ]
    const spots = measureCommentSpots(pm, [comment('new'), comment('f')], blocks, 0, 1, 0)
    expect(spots.map((s) => s.id)).toEqual(['f', 'new'])
    expect(spots.map((s) => s.no)).toEqual([1, 2])
  })

  it('keeps a provable span order over an ambiguous same-block fallback thread', () => {
    const pm = document.createElement('div')
    const p1 = para(0, [rect(40, 0, 400)])
    const sNew = span('new', [rect(40, 10)])
    const sFile = span('f', [rect(40, 200)])
    p1.append(sNew, sFile)
    const p2 = para(1, [rect(120, 0, 400)])
    pm.append(p1, p2)
    // the cross-paragraph thread "x" starts late in block 0 (offset past both
    // spans) but cannot be ordered against the session span; the session
    // thread's provable place before "f" must win
    const blocks = [
      {
        docxIndex: 0,
        originalXml:
          '<w:p><w:commentRangeStart w:id="f"/><w:commentRangeEnd w:id="f"/><w:r/>' +
          '<w:commentRangeStart w:id="x"/></w:p>',
      },
      { docxIndex: 1, originalXml: '<w:p><w:commentRangeEnd w:id="x"/></w:p>' },
    ]
    const spots = measureCommentSpots(
      pm,
      [comment('f'), comment('x'), comment('new')],
      blocks,
      0,
      1,
      0,
    )
    expect(spots.map((s) => s.id)).toEqual(['new', 'f', 'x'])
    expect(spots.map((s) => s.no)).toEqual([1, 2, 3])
  })

  it('pins a shared-span session thread before later same-block fallbacks', () => {
    const pm = document.createElement('div')
    const p1 = para(0, [rect(40, 0, 400)])
    const shared = span('f new', [rect(40, 10)])
    p1.append(shared)
    const p2 = para(1, [rect(120, 0, 400)])
    pm.append(p1, p2)
    // the session thread's first span is also early thread "f"'s, so its
    // range start is known to precede fallback "x"'s later marker
    const blocks = [
      {
        docxIndex: 0,
        originalXml:
          '<w:p><w:commentRangeStart w:id="f"/><w:commentRangeEnd w:id="f"/><w:r/>' +
          '<w:commentRangeStart w:id="x"/></w:p>',
      },
      { docxIndex: 1, originalXml: '<w:p><w:commentRangeEnd w:id="x"/></w:p>' },
    ]
    const spots = measureCommentSpots(
      pm,
      [comment('x'), comment('new'), comment('f')],
      blocks,
      0,
      1,
      0,
    )
    expect(spots.map((s) => s.id)).toEqual(['f', 'new', 'x'])
    expect(spots.map((s) => s.no)).toEqual([1, 2, 3])
  })

  it('skips replies, resolved threads, and threads without any anchor', () => {
    const pm = document.createElement('div')
    const p = document.createElement('p')
    p.append(span('1', [rect(10, 10)]))
    pm.append(p)

    const spots = measureCommentSpots(
      pm,
      [comment('1'), comment('2', { parentId: '1' }), comment('3', { done: true }), comment('4')],
      [],
      0,
      1,
      0,
    )
    expect(spots.map((s) => s.id)).toEqual(['1'])
    expect(spots[0].no).toBe(1)
  })

  it('scales measured coordinates by origin and zoom factor', () => {
    const pm = document.createElement('div')
    const p = document.createElement('p')
    p.append(span('1', [rect(220, 130, 40, 20)]))
    pm.append(p)

    const spots = measureCommentSpots(pm, [comment('1')], [], 20, 2, 30)
    expect(spots[0].top).toBe(100)
    expect(spots[0].endX).toBe(70)
    expect(spots[0].endY).toBe(109)
  })
})

/**
 * Long plain-text cells must reach `fillText` as the slice that can touch the
 * clip box, anchored on the aligned edge, so scrolling a grid of paragraph
 * cells no longer shapes megabytes of text per frame.
 */
import { HorizontalAlign } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  LONG_TEXT_MIN_CHARS,
  truncateDocForLayout,
  visibleSlice,
} from '../src/renderer/long-text-render'

const long = Array.from({ length: 10_000 }, (_, i) => String.fromCharCode(97 + (i % 26))).join('')
// 6 px per character, so 10k chars measure 60k px.
const width = long.length * 6

describe('visibleSlice', () => {
  it('keeps short or unmeasured text intact', () => {
    const short = 'x'.repeat(LONG_TEXT_MIN_CHARS)
    expect(visibleSlice(short, short.length * 6, 100, HorizontalAlign.LEFT)).toBe(short)
    expect(visibleSlice(long, 0, 100, HorizontalAlign.LEFT)).toBe(long)
    expect(visibleSlice(long, width, 0, HorizontalAlign.LEFT)).toBe(long)
  })

  it('keeps the head for left/unspecified alignment with slack', () => {
    const slice = visibleSlice(long, width, 600, HorizontalAlign.UNSPECIFIED)
    expect(long.startsWith(slice)).toBe(true)
    // 100 chars fit; 1.5× slack + 32 spare = 182.
    expect(slice.length).toBe(182)
    expect(visibleSlice(long, width, 600, HorizontalAlign.LEFT)).toBe(slice)
  })

  it('keeps the tail for right alignment and the middle for center', () => {
    const tail = visibleSlice(long, width, 600, HorizontalAlign.RIGHT)
    expect(long.endsWith(tail)).toBe(true)
    expect(tail.length).toBe(182)
    const mid = visibleSlice(long, width, 600, HorizontalAlign.CENTER)
    expect(mid.length).toBe(182)
    const start = Math.floor((long.length - mid.length) / 2)
    expect(mid).toBe(long.slice(start, start + mid.length))
  })

  it('returns the whole string when the clip box can show all of it', () => {
    expect(visibleSlice(long, width, width, HorizontalAlign.LEFT)).toBe(long)
  })
})

describe('truncateDocForLayout', () => {
  const text = Array.from({ length: 2000 }, (_, i) => String.fromCharCode(97 + (i % 26))).join('')
  const doc = {
    id: 'rich-cell',
    body: {
      dataStream: `${text}\r\n`,
      textRuns: [
        { st: 0, ed: 1000, ts: { bl: 1 } },
        { st: 1000, ed: 2000, ts: { it: 1 } },
      ],
      paragraphs: [{ startIndex: 2000 }],
      sectionBreaks: [{ startIndex: 2001 }],
    },
    documentStyle: {},
  }

  it('leaves short or multi-paragraph documents alone', () => {
    expect(truncateDocForLayout(doc, HorizontalAlign.LEFT, 2000)).toBe(doc)
    const multi = {
      ...doc,
      body: { ...doc.body, paragraphs: [{ startIndex: 5 }, { startIndex: 2000 }] },
    }
    expect(truncateDocForLayout(multi, HorizontalAlign.LEFT, 100)).toBe(multi)
  })

  it('keeps the head with runs clipped and terminators rebuilt', () => {
    const out = truncateDocForLayout(doc, HorizontalAlign.LEFT, 300)
    expect(out.body?.dataStream).toBe(`${text.slice(0, 300)}\r\n`)
    expect(out.body?.textRuns).toEqual([{ st: 0, ed: 300, ts: { bl: 1 } }])
    expect(out.body?.paragraphs).toEqual([{ startIndex: 300 }])
    expect(out.body?.sectionBreaks).toEqual([{ startIndex: 301 }])
  })

  it('keeps the tail for right alignment and shifts run offsets', () => {
    const out = truncateDocForLayout(doc, HorizontalAlign.RIGHT, 300)
    expect(out.body?.dataStream).toBe(`${text.slice(1700)}\r\n`)
    expect(out.body?.textRuns).toEqual([{ st: 0, ed: 300, ts: { it: 1 } }])
  })

  it('keeps the middle for center alignment across a run boundary', () => {
    const out = truncateDocForLayout(doc, HorizontalAlign.CENTER, 300)
    expect(out.body?.dataStream).toBe(`${text.slice(850, 1150)}\r\n`)
    expect(out.body?.textRuns).toEqual([
      { st: 0, ed: 150, ts: { bl: 1 } },
      { st: 150, ed: 300, ts: { it: 1 } },
    ])
  })
})

describe('right-to-left paragraphs', () => {
  const arabic = Array.from({ length: 2000 }, (_, i) =>
    String.fromCharCode(0x0627 + (i % 20)),
  ).join('')

  it('keeps the logical head at a right-aligned edge and the tail at a left one', () => {
    const width = arabic.length * 6
    const right = visibleSlice(arabic, width, 600, HorizontalAlign.RIGHT)
    expect(arabic.startsWith(right)).toBe(true)
    const left = visibleSlice(arabic, width, 600, HorizontalAlign.LEFT)
    expect(arabic.endsWith(left)).toBe(true)
    const general = visibleSlice(arabic, width, 600, HorizontalAlign.UNSPECIFIED)
    expect(arabic.startsWith(general)).toBe(true)
  })

  it('applies the same rule to layout documents', () => {
    const doc = {
      id: 'rich-cell',
      body: {
        dataStream: `${arabic}\r\n`,
        textRuns: [],
        paragraphs: [{ startIndex: 2000 }],
        sectionBreaks: [{ startIndex: 2001 }],
      },
      documentStyle: {},
    }
    expect(truncateDocForLayout(doc, HorizontalAlign.RIGHT, 300).body?.dataStream).toBe(
      `${arabic.slice(0, 300)}\r\n`,
    )
  })
})

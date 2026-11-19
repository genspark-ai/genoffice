// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import { mergePendingPgNum } from '../src/renderer/ai/page-setup'
import { fitTextScale } from '../src/renderer/editor/cell-fit-text'
import { chapterPrefixes } from '../src/renderer/pagination-pass'
import type { BlockBox, PageSlice } from '../src/renderer/pagination-types'

const block = (top: number, chapter?: string): BlockBox => {
  const el = document.createElement('p')
  if (chapter) el.setAttribute('data-chapter', chapter)
  return { top, height: 100, el }
}
const section = (chapter?: SectionInfo['pageNumberChapter']): SectionInfo =>
  ({ ...(chapter ? { pageNumberChapter: chapter } : {}) }) as SectionInfo
const page = (start: number, section = 0): PageSlice => ({ start, end: start + 1000, section })

describe('chapter numbers in page numbers', () => {
  const blocks = [
    block(0, '1:1'),
    block(100, '2:1.1'),
    block(1000),
    block(2500, '1:2'),
    block(3000, '1:C'),
  ]

  it('a page carries the number of the last chapter heading that starts on or before it', () => {
    const prefix = chapterPrefixes(blocks, [section({ style: 1, sep: 'hyphen' })])
    expect([0, 1000, 2000, 3000].map((start) => prefix(page(start)))).toEqual([
      '1-',
      '1-',
      '2-',
      'C-',
    ])
  })

  it('follows the heading level and the separator of the section the page is in', () => {
    const prefix = chapterPrefixes(blocks, [
      section({ style: 2, sep: 'colon' }),
      section(),
      section({ style: 1, sep: 'emDash' }),
    ])
    expect(prefix(page(0, 0))).toBe('1.1:')
    expect(prefix(page(1000, 1))).toBeUndefined()
    expect(prefix(page(2000, 2))).toBe('2\u2014')
  })

  it('a page before the first chapter heading has no chapter number', () => {
    const prefix = chapterPrefixes(
      [block(0), block(1500, '1:1')],
      [section({ style: 1, sep: 'period' })],
    )
    expect(prefix(page(0))).toBeUndefined()
    expect(prefix(page(1000))).toBe('1.')
  })
})

describe('fit text', () => {
  it('a line wider than its cell is squeezed to it, a narrower one is left to justification', () => {
    expect(fitTextScale(400, 200)).toBe(0.5)
    expect(fitTextScale(300, 200)).toBe(0.667)
    expect(fitTextScale(150, 200)).toBe(1)
    expect(fitTextScale(200.4, 200)).toBe(1)
    expect(fitTextScale(0, 200)).toBe(1)
    expect(fitTextScale(100, 0)).toBe(1)
  })
})

describe('an AI page-number change on the last section', () => {
  it('keeps the chapter number that is waiting to be saved', () => {
    const chapter = { style: 1, sep: 'colon' }
    expect(mergePendingPgNum({ fmt: 'decimal', chapter }, { fmt: 'upperRoman', start: 3 })).toEqual(
      { fmt: 'upperRoman', start: 3, chapter },
    )
    // taken off in the dialog and not saved yet: it stays off
    expect(mergePendingPgNum({ chapter: null }, { start: 2 })).toEqual({ start: 2, chapter: null })
  })

  it("leaves the file's own chapter number alone when none is pending", () => {
    expect(mergePendingPgNum(null, { fmt: 'lowerRoman' })).toEqual({ fmt: 'lowerRoman' })
    expect(mergePendingPgNum({ fmt: 'decimal' }, {})).toEqual({})
  })
})

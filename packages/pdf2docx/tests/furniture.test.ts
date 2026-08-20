/** Header/footer + page-number furniture detection unit tests (no wasm). */
import { describe, expect, it } from 'vitest'
import { detectFurniture, HF_PAGE_MARK, type FurniturePage } from '../src/analyze/furniture'
import type { PdfChar } from '../src/ir'
import { mkText } from './helpers/chars'

const HEIGHT = 792

function pageOf(index: number, ...lineChars: PdfChar[][]): FurniturePage {
  return { index, heightPt: HEIGHT, chars: lineChars.flat() }
}

const header = (text: string) => mkText(text, 72, { y: HEIGHT - 40 }).chars
const footer = (text: string) => mkText(text, 280, { y: 40 }).chars
const body = (text: string, y = 400) => mkText(text, 72, { y }).chars

describe('detectFurniture: repeated headers', () => {
  it('drops a repeated header everywhere and re-emits it as an hf slot (P17)', () => {
    const pages = [0, 1, 2, 3].map((i) =>
      pageOf(i, header('ACME REPORT'), body(`Body text ${'x'.repeat(i)}`)),
    )
    const { drop, droppedLines, hf } = detectFurniture(pages)
    // the slot re-renders as a real docx header on every page, so ALL
    // occurrences leave the body (the first kept one would double on page 1)
    expect(drop.every((s) => s.size > 0)).toBe(true)
    expect(droppedLines).toBe(4)
    const dropped = [...drop[1]!].map((c) => c.text).join('')
    expect(dropped.trim().replace(/\s+/g, ' ')).toBe('ACME REPORT')
    expect(hf).toHaveLength(1)
    expect(hf[0]).toMatchObject({
      band: 'top',
      text: 'ACME REPORT',
      pageNo: false,
      coversFirstPage: true,
    })
  })

  it('keeps top text that appears on only two pages (content, not a header)', () => {
    const pages = [
      pageOf(0, header('Chapter One'), body('a')),
      pageOf(1, header('Chapter One'), body('b')),
      pageOf(2, body('c')),
      pageOf(3, body('d')),
    ]
    const { droppedLines } = detectFurniture(pages)
    expect(droppedLines).toBe(0)
  })

  it('keeps distinct numbered headlines apart (様式第１号 vs 様式第２号)', () => {
    const pages = [0, 1, 2, 3, 4].map((i) => pageOf(i, header(`様式第${i}号`), body('本文')))
    expect(detectFurniture(pages).droppedLines).toBe(0)
  })
})

describe('detectFurniture: page numbers', () => {
  it('drops bare page numbers on every page, first included', () => {
    const pages = [0, 1, 2].map((i) => pageOf(i, footer(String(i + 1)), body('text')))
    const { drop, droppedLines } = detectFurniture(pages)
    expect(droppedLines).toBe(3)
    expect(drop[0]!.size).toBeGreaterThan(0)
  })

  it('drops decorated and worded page numbers ("— 2 —", "第 1 页 / 共12页")', () => {
    const dashed = [0, 1, 2].map((i) => pageOf(i, footer(`— ${i + 1} —`), body('text')))
    expect(detectFurniture(dashed).droppedLines).toBe(3)
    const worded = [0, 1, 2].map((i) => pageOf(i, footer(`第 ${i + 1} 页 / 共3页`), body('正文')))
    expect(detectFurniture(worded).droppedLines).toBe(3)
  })

  it('leaves numeric body text alone (not in the edge bands)', () => {
    const pages = [0, 1, 2].map((i) => pageOf(i, body('42', 400)))
    expect(detectFurniture(pages).droppedLines).toBe(0)
  })
})

describe('detectFurniture: hidden micro text (P11 C)', () => {
  it('keeps sub-readable repeated lines on every page (invisible watermark junk)', () => {
    const junk = () => mkText('PPT模板下载：www.1ppt.com/moban/', 280, { y: 40, fontSize: 1 }).chars
    const pages = [0, 1, 2, 3].map((i) => pageOf(i, junk(), body(`Body ${'x'.repeat(i)}`)))
    const { drop, droppedLines } = detectFurniture(pages)
    expect(droppedLines).toBe(0)
    expect(drop.every((s) => s.size === 0)).toBe(true)
  })

  it('still drops a readable repeated footer at the same spot', () => {
    const pages = [0, 1, 2, 3].map((i) => pageOf(i, footer('Corp Confidential'), body(`Body ${i}`)))
    const { droppedLines, hf } = detectFurniture(pages)
    // emitted as an hf slot → all four occurrences leave the body
    expect(droppedLines).toBe(4)
    expect(hf[0]).toMatchObject({ band: 'bottom', text: 'Corp Confidential' })
  })
})

describe('detectFurniture: hf re-emission (P17)', () => {
  it('turns a "Page N of M" footer into a PAGE-field slot', () => {
    const pages = [0, 1, 2, 3].map((i) =>
      pageOf(i, footer(`Page ${i + 1} of 4`), body(`Body ${i}`)),
    )
    const { hf } = detectFurniture(pages)
    expect(hf).toHaveLength(1)
    expect(hf[0]!.pageNo).toBe(true)
    expect(hf[0]!.text).toBe(`Page ${HF_PAGE_MARK} of 4`)
    expect(hf[0]!.coversFirstPage).toBe(true)
  })

  it('does not emit a pageNo slot whose numbers do not track the physical page', () => {
    // numbering starts at 5 → offset ≠ 0, reproducing it wrong is worse
    const pages = [0, 1, 2, 3].map((i) => pageOf(i, footer(String(i + 5)), body(`B ${i}`)))
    const { hf, droppedLines } = detectFurniture(pages)
    expect(hf).toHaveLength(0)
    expect(droppedLines).toBe(4) // still dropped from the body
  })

  it('flags a header that skips the first page (cover) for titlePg handling', () => {
    const pages = [
      pageOf(0, body('Cover title')),
      ...[1, 2, 3, 4].map((i) => pageOf(i, header('EU Grants: template'), body(`B${i}`))),
    ]
    const { hf } = detectFurniture(pages)
    expect(hf).toHaveLength(1)
    expect(hf[0]!.coversFirstPage).toBe(false)
  })
})

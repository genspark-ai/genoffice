import { describe, expect, it } from 'vitest'
import { buildOcrPageData } from '../src/renderer/ocr-layer'
import type { PdfOcrLine } from '../src/shared/ipc'

const GEOM = { pw: 600, ph: 800, rot: 0 }

const box = (x0: number, x1: number): [number, number, number, number] => [x0, 0.8, x1, 0.85]
const SEPARATOR: [number, number, number, number] = [0, 1, 0, 1]

describe('buildOcrPageData word joining', () => {
  it('keeps CJK characters unspaced so search matches the query as typed', () => {
    const line: PdfOcrLine = {
      text: '使用价值',
      confidence: 1,
      box: box(0.1, 0.5),
      chars: [
        { text: '使', box: box(0.1, 0.2) },
        { text: '用', box: box(0.2, 0.3) },
        { text: '价', box: box(0.3, 0.4) },
        { text: '值', box: box(0.4, 0.5) },
      ],
    }
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('使用价值\n')
    expect(data.entry.lower.includes('使用价值')).toBe(true)
    expect(data.words.map((w) => w.text)).toEqual(['使', '用', '价', '值'])
    expect(data.words.every((w) => !w.spaceAfter)).toBe(true)
  })

  it('keeps a single space where the OCR stream had a separator', () => {
    const line: PdfOcrLine = {
      text: 'hello world',
      confidence: 1,
      box: box(0.1, 0.6),
      chars: [
        ...[...'hello'].map((t) => ({ text: t, box: box(0.1, 0.3) })),
        { text: ' ', box: SEPARATOR },
        ...[...'world'].map((t) => ({ text: t, box: box(0.35, 0.6) })),
      ],
    }
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('hello world\n')
    expect(data.words.map((w) => w.text)).toEqual(['hello', 'world'])
    expect(data.words[0]!.spaceAfter).toBe(true)
    expect(data.words[1]!.spaceAfter).toBe(true)
  })

  it('drops Windows-style segmentation separators between CJK words', () => {
    // Windows OcrLine.Text space-joins its CJK words; the helper marks those
    // separators with zero boxes. They must not become spaces in the index.
    const line: PdfOcrLine = {
      text: '使用 价值',
      confidence: 1,
      box: box(0.1, 0.5),
      chars: [
        { text: '使', box: box(0.1, 0.2) },
        { text: '用', box: box(0.2, 0.3) },
        { text: ' ', box: [0, 0, 0, 0] },
        { text: '价', box: box(0.3, 0.4) },
        { text: '值', box: box(0.4, 0.5) },
      ],
    }
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('使用价值\n')
  })

  it('keeps real Korean word spaces (Hangul is space-delimited)', () => {
    const line: PdfOcrLine = {
      text: '사용 가치',
      confidence: 1,
      box: box(0.1, 0.5),
      chars: [
        { text: '사', box: box(0.1, 0.15) },
        { text: '용', box: box(0.15, 0.2) },
        { text: ' ', box: SEPARATOR },
        { text: '가', box: box(0.25, 0.3) },
        { text: '치', box: box(0.3, 0.35) },
      ],
    }
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('사용 가치\n')
    expect(data.words[data.words.length - 1]!.spaceAfter).toBe(true)
  })

  it('keeps the space on a CJK/Latin boundary', () => {
    const line: PdfOcrLine = {
      text: '对比 Windows 系统',
      confidence: 1,
      box: box(0.1, 0.7),
      chars: [
        { text: '对', box: box(0.1, 0.15) },
        { text: '比', box: box(0.15, 0.2) },
        { text: ' ', box: SEPARATOR },
        ...[...'Windows'].map((t) => ({ text: t, box: box(0.25, 0.45) })),
        { text: ' ', box: SEPARATOR },
        { text: '系', box: box(0.5, 0.55) },
        { text: '统', box: box(0.55, 0.6) },
      ],
    }
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('对比 Windows 系统\n')
  })

  it('no-space rule follows the shared script table: thai suppressed, fullwidth punct line-end', () => {
    const thai: PdfOcrLine = {
      text: 'มูลค่า การใช้',
      confidence: 1,
      box: box(0.1, 0.5),
      chars: [
        ...[...'มูลค่า'].map((t) => ({ text: t, box: box(0.1, 0.2) })),
        { text: ' ', box: [0, 0, 0, 0] as [number, number, number, number] },
        ...[...'การใช้'].map((t) => ({ text: t, box: box(0.25, 0.35) })),
      ],
    }
    const punct: PdfOcrLine = {
      text: '确认！',
      confidence: 1,
      box: box(0.1, 0.3),
      chars: [
        { text: '确', box: box(0.1, 0.15) },
        { text: '认', box: box(0.15, 0.2) },
        { text: '！', box: box(0.2, 0.25) },
      ],
    }
    const data = buildOcrPageData([thai, punct], GEOM)!
    expect(data.entry.text).toBe('มูลค่าการใช้\n确认！\n')
    expect(data.words[data.words.length - 1]!.spaceAfter).toBe(false)
  })

  it('offsets in index items match the joined text', () => {
    const line: PdfOcrLine = {
      text: '第一篇 商品',
      confidence: 1,
      box: box(0.1, 0.7),
      chars: [
        { text: '第', box: box(0.1, 0.2) },
        { text: '一', box: box(0.2, 0.3) },
        { text: '篇', box: box(0.3, 0.4) },
        { text: ' ', box: SEPARATOR },
        { text: '商', box: box(0.45, 0.55) },
        { text: '品', box: box(0.55, 0.65) },
      ],
    }
    // the Han|Han separator is dropped (queries are typed without it)
    const data = buildOcrPageData([line], GEOM)!
    expect(data.entry.text).toBe('第一篇商品\n')
    for (const item of data.entry.items) {
      expect(data.entry.text.slice(item.start, item.end).trim().length).toBeGreaterThan(0)
    }
    const shang = data.entry.items[3]!
    expect(data.entry.text.slice(shang.start, shang.end)).toBe('商')
  })
})

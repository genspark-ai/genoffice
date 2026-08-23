/** Local OCR recovery for scanned pages: policy gates, char synthesis, and
 * the pipeline integration (src/ocr.ts). The engine itself is faked — the
 * real platform engines (Vision helper) are exercised by eval scripts. */
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import type { ExtractedPage } from '../src/extract'
import { ocrLinesToChars, tryOcrScannedPage, type OcrEngine, type OcrLine } from '../src/ocr'
import { buildScannedPdf } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

/** normalized line box helper: y measured from the bottom, 0–1 */
const nbox = (x0: number, y0: number, x1: number, y1: number) => ({ x0, y0, x1, y1 })

const CJK_LINE: OcrLine = {
  text: '扫描件文字识别测试',
  confidence: 0.9,
  box: nbox(0.1, 0.85, 0.55, 0.88),
}
const LATIN_LINE: OcrLine = {
  text: 'Recovered by local OCR engine',
  confidence: 0.95,
  box: nbox(0.1, 0.8, 0.6, 0.82),
}

function scannedExtractedPage(over: Partial<ExtractedPage> = {}): ExtractedPage {
  return {
    index: 0,
    widthPt: 612,
    heightPt: 792,
    rotation: 0,
    chars: [],
    images: [],
    paths: [],
    degraded: false,
    scanned: true,
    hasStructTree: false,
    vectorRegions: [],
    badUnicodeRatio: 0,
    render: { data: new Uint8Array([1]), mime: 'image/png', pixelWidth: 612, pixelHeight: 792 },
    ...over,
  }
}

describe('ocrLinesToChars', () => {
  it('synthesizes CJK chars with a shared baseline and em-sized advances', () => {
    const chars = ocrLinesToChars([CJK_LINE], 612, 792)
    expect(chars).toHaveLength(9)
    // all chars share the line baseline
    expect(new Set(chars.map((c) => c.originY)).size).toBe(1)
    // x advances are monotonic and cover the line box
    for (let i = 1; i < chars.length; i++) {
      expect(chars[i]!.box.x0).toBeGreaterThanOrEqual(chars[i - 1]!.box.x0)
    }
    expect(chars[0]!.box.x0).toBeCloseTo(0.1 * 612, 5)
    expect(chars[chars.length - 1]!.box.x1).toBeCloseTo(0.55 * 612, 5)
    // CJK line height ≈ font size
    const lineH = (0.88 - 0.85) * 792
    expect(chars[0]!.fontSize).toBeGreaterThan(lineH * 0.8)
    expect(chars[0]!.fontSize).toBeLessThanOrEqual(lineH)
    expect(chars[0]!.script).toBe('cjk')
  })

  it('emits spaces as real glyphs so word gaps survive the line builder', () => {
    const chars = ocrLinesToChars([LATIN_LINE], 612, 792)
    expect(chars.map((c) => c.text).join('')).toBe('Recovered by local OCR engine')
    const space = chars[9]!
    expect(space.text).toBe(' ')
    expect(space.box.x1).toBeGreaterThan(space.box.x0)
  })

  it('anchors word segments to engine boxes, falls back when partial', () => {
    // two words with a wide real gap: 'ab' at 100..200, 'cd' at 500..600 —
    // the gap must survive as a space glyph spanning it (form label vs value)
    const withBoxes: OcrLine = {
      text: 'ab cd',
      confidence: 1,
      box: nbox(0.1, 0.5, 0.6, 0.52),
      chars: [
        { text: 'a', box: nbox(0.1, 0.5, 0.15, 0.52) },
        { text: 'b', box: nbox(0.14, 0.5, 0.2, 0.52) }, // overlaps 'a' (Vision does this)
        { text: ' ', box: nbox(0, 0, 0, 0) },
        { text: 'c', box: nbox(0.5, 0.5, 0.55, 0.52) },
        { text: 'd', box: nbox(0.55, 0.5, 0.6, 0.52) },
      ],
    }
    const chars = ocrLinesToChars([withBoxes], 1000, 1000)
    expect(chars.map((c) => c.text).join('')).toBe('ab cd')
    // glyph boxes are monotone and non-overlapping despite overlapping input
    for (let i = 1; i < chars.length; i++) {
      expect(chars[i]!.box.x0).toBeGreaterThanOrEqual(chars[i - 1]!.box.x1 - 1e-6)
    }
    // the word gap spans between the segment anchors
    const space = chars[2]!
    expect(space.text).toBe(' ')
    expect(space.box.x0).toBeCloseTo(200, 3)
    expect(space.box.x1).toBeCloseTo(500, 3)

    // a zero-width box on a non-space glyph invalidates the whole set
    const partial: OcrLine = {
      ...withBoxes,
      text: 'ab',
      chars: [
        { text: 'a', box: nbox(0.1, 0.5, 0.15, 0.52) },
        { text: 'b', box: nbox(0, 0, 0, 0) },
      ],
    }
    const fallback = ocrLinesToChars([partial], 1000, 1000)
    expect(fallback).toHaveLength(2)
    // distributed evenly across the whole line box instead (weights equal)
    expect(fallback[1]!.box.x1).toBeCloseTo(600, 3)
  })
})

describe('tryOcrScannedPage gates', () => {
  it('recovers a page from good lines', () => {
    const engine: OcrEngine = () => ({ lines: [CJK_LINE, LATIN_LINE], paperShare: 0.95 })
    const result = tryOcrScannedPage(scannedExtractedPage(), engine)
    expect(result).not.toBeNull()
    expect(result!.page.ocrRecovered).toBe(true)
    expect(result!.page.scanned).toBe(false)
    expect(result!.page.render).toBeUndefined()
    expect(result!.confidence).toBeCloseTo(0.925, 3)
  })

  it.each<[string, ReturnType<OcrEngine>]>([
    ['engine failure', null],
    ['no lines', { lines: [] }],
    ['low mean confidence', { lines: [{ ...CJK_LINE, confidence: 0.4 }] }],
    [
      'too few chars (photo with stray text)',
      { lines: [{ ...CJK_LINE, text: '路牌', confidence: 0.9 }] },
    ],
    ['a photo page (low paper share)', { lines: [CJK_LINE, LATIN_LINE], paperShare: 0.1 }],
    [
      'a photo on a white backdrop (tiny text coverage)',
      {
        lines: [
          { ...CJK_LINE, text: '产品图片说明文字', box: { x0: 0.4, y0: 0.1, x1: 0.48, y1: 0.11 } },
        ],
        paperShare: 0.9,
      },
    ],
  ])('falls back to the bitmap on %s', (_name, recognition) => {
    const engine: OcrEngine = () => recognition
    expect(tryOcrScannedPage(scannedExtractedPage(), engine)).toBeNull()
  })

  it('drops sub-threshold lines but keeps the rest', () => {
    const engine: OcrEngine = () => ({ lines: [CJK_LINE, { ...LATIN_LINE, confidence: 0.1 }] })
    const result = tryOcrScannedPage(scannedExtractedPage(), engine)
    expect(result).not.toBeNull()
    expect(result!.confidence).toBeCloseTo(0.9, 3)
  })

  it('falls back when the engine throws', () => {
    const engine: OcrEngine = () => {
      throw new Error('helper crashed')
    }
    expect(tryOcrScannedPage(scannedExtractedPage(), engine)).toBeNull()
  })

  it('requires a page render', () => {
    const engine: OcrEngine = () => ({ lines: [CJK_LINE, LATIN_LINE], paperShare: 0.95 })
    expect(tryOcrScannedPage(scannedExtractedPage({ render: undefined }), engine)).toBeNull()
  })
})

describe('pipeline integration', () => {
  it('converts a scanned page to editable text with an OCR engine', async () => {
    const pdfium = await loadPdfium()
    const engine: OcrEngine = () => ({ lines: [CJK_LINE, LATIN_LINE], paperShare: 0.95 })
    const result = await convertPdfToDocx(await buildScannedPdf(), {
      pdfium,
      renderScale: 1,
      ocr: engine,
    })
    expect(result.pageResults[0]!.status).toBe('ocr')
    expect(result.pageResults[0]!.confidence).toBeCloseTo(0.925, 3)
    expect(result.warnings.some((w) => /recovered via local OCR/.test(w))).toBe(true)
    expect(result.scannedDocument).toBe(false)
    const parsed = await parseDocx(result.docx)
    const text = JSON.stringify(parsed.blocks)
    expect(text).toContain('扫描件文字识别测试')
    expect(text).toContain('Recovered')
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(0)
  })

  it('never re-degrades an OCR-recovered page via the P4 confidence floor', async () => {
    // invariant guarded by !page.ocrRecovered in the fidelity floor: however
    // noisy the recognized text scores in the analyze layer, a page reported
    // as 'ocr' must actually ship editable text, never a bitmap fallback
    const noisy = '�'.repeat(12)
    const engine: OcrEngine = () => ({
      lines: [
        { text: noisy, confidence: 0.9, box: nbox(0.1, 0.85, 0.55, 0.88) },
        { text: noisy, confidence: 0.9, box: nbox(0.1, 0.845, 0.55, 0.875) },
        CJK_LINE,
      ],
      paperShare: 0.95,
    })
    const result = await convertPdfToDocx(await buildScannedPdf(), {
      pdfium: await loadPdfium(),
      renderScale: 1,
      ocr: engine,
    })
    expect(result.pageResults[0]!.status).toBe('ocr')
    const parsed = await parseDocx(result.docx)
    // the page must actually BE editable text — no full-page image fallback
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(0)
    expect(JSON.stringify(parsed.blocks)).toContain('扫描件文字识别测试')
  })

  it('keeps the bitmap fallback when the engine finds nothing', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildScannedPdf(), {
      pdfium,
      renderScale: 1,
      ocr: () => null,
    })
    expect(result.pageResults[0]!.status).toBe('scanned')
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(1)
  })
})

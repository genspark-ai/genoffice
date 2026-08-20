/**
 * End-to-end conversion through the real PDFium wasm: fixture PDF (generated
 * on the fly) → convertPdfToDocx → docx-engine parseDocx assertions.
 */
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import {
  buildCjkPdf,
  buildImagePdf,
  buildLatinPdf,
  buildScannedPdf,
  buildSpacedPdf,
  buildStreamTablePdf,
  buildTablePdf,
  buildTwoColumnPdf,
  cjkFontBytes,
  buildCheckboxFormPdf,
  buildWideScannedPdf,
} from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

const paraTexts = async (docx: Uint8Array): Promise<string[]> => {
  const parsed = await parseDocx(docx)
  return parsed.blocks
    .filter((b) => !b.hidden && b.type !== 'image')
    .map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    .filter((t) => t.length > 0)
}

describe('integration: latin PDF (pdf-lib fixture)', () => {
  it('converts a two-paragraph text PDF with correct word spacing and paragraph splits', async () => {
    const pdfium = await loadPdfium()
    const pdf = await buildLatinPdf()
    const progress: Array<[number, number]> = []
    const result = await convertPdfToDocx(pdf, {
      pdfium,
      onProgress: (p, t) => progress.push([p, t]),
    })

    expect(result.pages).toBe(1)
    expect(result.warnings).toEqual([])
    expect(progress).toEqual([[1, 1]])

    const texts = await paraTexts(result.docx)
    expect(texts).toContain('Sample Document')
    expect(texts).toContain(
      'The quick brown fox jumps over the lazy dog while the cat watches from a warm windowsill.',
    )
    expect(texts).toContain('A second paragraph sits below the first one after a wider gap.')
  })

  it('carries font size and boldness through to the runs', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildLatinPdf(), { pdfium })
    const parsed = await parseDocx(result.docx)
    const title = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('Sample Document')),
    )!
    const run = title.runs!.find((r) => r.text.includes('Sample'))!
    expect(run.sizeHalfPoints).toBe(36) // 18pt
    expect(run.bold).toBe(true)
    // centered against the page's mirrored margins, not its own right edge
    expect(title.format?.align).toBe('center')
  })
})

describe('integration: images', () => {
  it('embeds page images as docx pictures', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildImagePdf(), { pdfium })
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(1)
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('Figure below:')
  })

  it('detects scanned pages and falls back to a full-page render', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildScannedPdf(), { pdfium, renderScale: 1 })
    expect(result.warnings).toHaveLength(1)
    expect(result.warnings[0]).toMatch(/page 1: scanned/)
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(1)
  })

  it("shrinks wide-format scans to Word's 22in page ceiling instead of cropping", async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildWideScannedPdf(), { pdfium, renderScale: 1 })
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(result.docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    const sizes = [...xml.matchAll(/<w:pgSz w:w="(\d+)" w:h="(\d+)"/g)]
    expect(sizes.length).toBeGreaterThan(0)
    const maxTwips = 22 * 72 * 20
    for (const m of sizes) {
      expect(Number(m[1])).toBeLessThanOrEqual(maxTwips)
      expect(Number(m[2])).toBeLessThanOrEqual(maxTwips)
    }
    // aspect ratio survives the clamp (1800 × 1012.5 → 16:9)
    const ratio = Number(sizes[0]![1]) / Number(sizes[0]![2])
    expect(ratio).toBeCloseTo(1800 / 1012.5, 1)
  })
})

describe('integration: AcroForm checkbox widgets (P29)', () => {
  it('synthesizes glyph chars beside their labels, checked state included', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildCheckboxFormPdf(), { pdfium })
    const texts = await paraTexts(result.docx)
    const line = texts.find((t) => t.includes('Individual'))
    expect(line).toBeDefined()
    expect(line!).toContain('\u2612')
    expect(line!).toContain('\u2610')
    // glyphs interleave in reading order: box before its own label
    expect(line!.indexOf('\u2612')).toBeLessThan(line!.indexOf('Individual'))
    expect(line!.indexOf('\u2610')).toBeLessThan(line!.indexOf('Corporation'))
  })
})

describe('integration: lattice tables (pdf-lib fixture)', () => {
  it('converts a bordered table into a real docx table with merge + shading', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildTablePdf(), { pdfium })
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)

    // the paragraph above the table stays in the normal flow, before the table
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('Before the table.')
    const tableIdx = parsed.blocks.findIndex((b) => b.type === 'table')
    const paraIdx = parsed.blocks.findIndex((b) =>
      (b.runs ?? []).some((r) => r.text.includes('Before the table.')),
    )
    expect(tableIdx).toBeGreaterThan(paraIdx)

    const model = parsed.blocks[tableIdx]!.table!
    expect(model.rows).toHaveLength(2)
    // header row: three cells, texts land in their cells, first cell shaded
    expect(model.rows[0]!.map((c) => c.paras[0])).toEqual(['Name', 'Qty', 'Price'])
    expect(model.rows[0]![0]!.fill?.toUpperCase()).toBe('FFCC00')
    // bottom row: 'Total' + one merged cell spanning columns 1-2
    expect(model.rows[1]!.map((c) => c.paras[0])).toEqual(['Total', 'wide merged cell'])
    expect(model.rows[1]![1]!.colSpan).toBe(2)
  })

  it('does not misfire on plain text pages (no strokes → no tables)', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildLatinPdf(), { pdfium })
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.filter((b) => b.type === 'table')).toHaveLength(0)
  })
})

describe('integration: stream tables (P3, pdf-lib fixture)', () => {
  it('converts a borderless three-line table into a real docx table', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildStreamTablePdf(), { pdfium })
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)

    const texts = await paraTexts(result.docx)
    expect(texts).toContain('Quarterly results follow:')
    const tableBlock = parsed.blocks.find((b) => b.type === 'table')
    expect(tableBlock).toBeDefined()
    const model = tableBlock!.table!
    expect(model.rows).toHaveLength(3)
    expect(model.rows[0]!.map((c) => c.paras[0])).toEqual(['Region', 'Revenue'])
    expect(model.rows[2]!.map((c) => c.paras[0])).toEqual(['South', '890'])
    // the paragraph stays in the flow, before the table
    const tableIdx = parsed.blocks.findIndex((b) => b.type === 'table')
    const paraIdx = parsed.blocks.findIndex((b) =>
      (b.runs ?? []).some((r) => r.text.includes('Quarterly')),
    )
    expect(tableIdx).toBeGreaterThan(paraIdx)
  })

  it('does not misfire on the two-column text page', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildTwoColumnPdf(), { pdfium })
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.filter((b) => b.type === 'table')).toHaveLength(0)
  })
})

describe('integration: columns (P3, pdf-lib fixture)', () => {
  it('converts a two-column page into a w:cols section with correct reading order', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildTwoColumnPdf(), { pdfium })
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)

    const { readSectionSettings } = await import('@genoffice/docx-engine')
    const settings = readSectionSettings(parsed)
    expect(settings.columns).toBe(2)

    const all = (await paraTexts(result.docx)).join('\n')
    expect(all).toContain('Left column begins here')
    expect(all).toContain('Right column starts with')
    // reading order: whole left column before the right column
    expect(all.indexOf('smooth grey stones')).toBeLessThan(all.indexOf('Right column starts'))
  })
})

describe('integration: before_space chain (P3, pdf-lib fixture)', () => {
  it('turns big vertical whitespace into spacingBefore on title and body', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildSpacedPdf(), { pdfium })
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)

    const title = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('Spaced Title')),
    )!
    const body = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('body paragraph starts')),
    )!
    // page-top whitespace beyond the clamped 108pt margin: ~31pt → ~620 twips
    expect(title.format?.spaceBefore).toBeGreaterThan(300)
    expect(title.format?.spaceBefore).toBeLessThan(1000)
    // title → body gap ≈ 127pt minus normal leading → ~2500 twips
    expect(body.format?.spaceBefore).toBeGreaterThan(2000)
    expect(body.format?.spaceBefore).toBeLessThan(3000)
  })
})

// CJK fixtures need a system font with CJK coverage (pdf-lib's standard 14 have none)
const hasCjkFont = cjkFontBytes() !== null

describe.skipIf(!hasCjkFont)('integration: zh/ja/ko (PDFium-generated fixtures)', () => {
  it('zh: lines of one paragraph join WITHOUT inserted spaces', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: '这是第一行的中文内容测试', x: 72, y: 700 },
      { text: '第二行继续说明相关文字。', x: 72, y: 685 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    expect(result.warnings).toEqual([])
    const texts = await paraTexts(result.docx)
    expect(texts.join('\n')).toContain('这是第一行的中文内容测试第二行继续说明相关文字。')
  })

  it('zh/en mixed: scripts split into separate runs, no spaces invented', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: '中文English混排测试', x: 72, y: 700 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('中文English混排测试')
    const parsed = await parseDocx(result.docx)
    const para = parsed.blocks.find((b) => (b.runs ?? []).some((r) => r.text.includes('English')))!
    expect(para.runs!.length).toBeGreaterThanOrEqual(3) // cjk / latin / cjk
  })

  it('ja: kana + kanji text survives intact', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: 'こんにちは世界。日本語のテストです。', x: 72, y: 700 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('こんにちは世界。日本語のテストです。')
  })

  it('ko: hangul keeps its real inter-word spaces', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: '안녕하세요 세계 여러분', x: 72, y: 700 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('안녕하세요 세계 여러분')
  })
})

// Arabic/Hebrew integration reuses the same wide-coverage system font.
// PDFium's simple text objects place chars LTR in string order without shaping,
// so writing the REVERSED logical string yields a visual-order page — exactly
// what real PDF writers produce for RTL text.
describe.skipIf(!hasCjkFont)('integration: ar/he RTL (P2 — no more bitmap fallback)', () => {
  const visual = (logical: string): string => [...logical].reverse().join('')

  it('ar: converts to logical-order text with bidi paragraph and rtl runs', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: visual('مرحبا بالعالم'), x: 350, y: 700 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    // P1 degraded these pages to bitmaps; P2 must not
    expect(result.warnings).toEqual([])
    const parsed = await parseDocx(result.docx)
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.runs!.map((r) => r.text).join('')).toBe('مرحبا بالعالم')
    expect(para.format?.bidi).toBe(true)
    const arRun = para.runs!.find((r) => /[؀-ۿ]/.test(r.text))!
    expect(arRun.rtl).toBe(true)
    expect(arRun.fontCs).toBeTruthy()
  })

  it('ar: digit runs inside RTL text keep LTR order', async () => {
    const pdfium = await loadPdfium()
    // logical "رقم 25 هنا" — visual layout reverses the Arabic but keeps "25"
    const logical = 'رقم 25 هنا'
    const visualText = [...logical].reverse().join('').replace('52', '25')
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [{ text: visualText, x: 300, y: 700 }])
    const result = await convertPdfToDocx(pdf, { pdfium })
    const parsed = await parseDocx(result.docx)
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.runs!.map((r) => r.text).join('')).toBe(logical)
  })

  it('he: hebrew line comes out in logical order', async () => {
    const pdfium = await loadPdfium()
    const pdf = buildCjkPdf(pdfium, cjkFontBytes()!, [
      { text: visual('שלום עולם'), x: 400, y: 700 },
    ])
    const result = await convertPdfToDocx(pdf, { pdfium })
    expect(result.warnings).toEqual([])
    const texts = await paraTexts(result.docx)
    expect(texts).toContain('שלום עולם')
  })
})

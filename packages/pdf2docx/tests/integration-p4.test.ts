/**
 * P4 end-to-end coverage through the real PDFium wasm: style mapping,
 * list rebuild, vector-illustration rasterization and the per-page result API.
 */
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import {
  buildLatinPdf,
  buildListPdf,
  buildScannedPdf,
  buildStyledPdf,
  buildVectorArtPdf,
} from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

describe('integration: text × shape styles', () => {
  it('maps highlight fills, underlines and strikethroughs onto runs', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildStyledPdf(), { pdfium })
    expect(result.pageResults).toEqual([expect.objectContaining({ page: 1, status: 'ok' })])
    const parsed = await parseDocx(result.docx)
    const runOf = (needle: string) => {
      for (const b of parsed.blocks) {
        const run = (b.runs ?? []).find((r) => r.text.includes(needle))
        if (run) return run
      }
      return undefined
    }
    const plain = runOf('Plain opening')!
    expect(plain.underline).toBeFalsy()
    expect(plain.strike).toBeFalsy()
    expect(plain.highlight).toBeFalsy()

    expect(runOf('Yellow highlighted')!.highlight).toBe('yellow')
    expect(runOf('Underlined words')!.underline).toBe(true)
    expect(runOf('Underlined words')!.strike).toBeFalsy()
    expect(runOf('Struck through')!.strike).toBe(true)
    expect(runOf('Struck through')!.underline).toBeFalsy()
  })
})

describe('integration: lists', () => {
  it('rebuilds bullet and numbered items as real docx list paragraphs', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildListPdf(), { pdfium })
    const parsed = await parseDocx(result.docx)

    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items).toHaveLength(6)

    const bullets = items.filter((b) => b.list?.kind === 'bullet')
    expect(bullets).toHaveLength(3)
    expect(bullets.map((b) => b.runs!.map((r) => r.text).join(''))).toEqual([
      'apples and pears',
      'bread with butter',
      'milk in bottles',
    ])

    const ordered = items.filter((b) => b.list?.kind === 'ordered')
    expect(ordered).toHaveLength(3)
    expect(ordered.map((b) => b.runs!.map((r) => r.text).join(''))).toEqual([
      'first step here',
      'second step done',
      'third step ends',
    ])
    // one shared numId for the sequential run, all at level 0
    expect(new Set(ordered.map((b) => b.list!.numId)).size).toBe(1)
    expect(ordered.every((b) => b.list!.ilvl === 0)).toBe(true)

    // the intro stays a plain paragraph
    const intro = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('Shopping list intro')),
    )!
    expect(intro.type).toBe('paragraph')
  })
})

describe('integration: vector illustrations', () => {
  it('rasterizes the curve cluster into an embedded image and keeps the text', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildVectorArtPdf(), { pdfium, renderScale: 1 })
    expect(result.warnings.some((w) => /vector illustration/.test(w))).toBe(true)

    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.some((b) => b.type === 'image')).toBe(true)
    const texts = parsed.blocks
      .map((b) => (b.runs ?? []).map((r) => r.text).join(''))
      .filter((t) => t.length > 0)
    expect(texts).toContain('Chart on this page:')
    expect(texts).toContain('And a closing paragraph well below the chart.')
  })
})

describe('integration: per-page result API', () => {
  it('reports ok pages with confidence and no scanned verdict', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildLatinPdf(), { pdfium })
    expect(result.scannedDocument).toBe(false)
    expect(result.pageResults).toHaveLength(1)
    expect(result.pageResults[0]).toMatchObject({ page: 1, status: 'ok' })
    expect(result.pageResults[0]!.confidence).toBeGreaterThan(0.8)
  })

  it('flags a scanned document for the caller UI split', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(await buildScannedPdf(), { pdfium, renderScale: 1 })
    expect(result.scannedDocument).toBe(true)
    expect(result.pageResults).toEqual([expect.objectContaining({ page: 1, status: 'scanned' })])
  })
})

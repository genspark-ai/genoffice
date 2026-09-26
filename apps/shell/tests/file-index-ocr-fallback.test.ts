// OCR fallback in the content extractor: scanned PDFs (no text layer) get a
// second chance through the injected OCR hook — only for PDFs, only when the
// hook is wired, and failures degrade to the empty-text extraction.
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeEach, describe, expect, it, vi } from 'vitest'

import { extractText, SCANNED_PDF_TEXT_CHARS } from '../src/main/file-index/extract'
import { parseFileToText } from '@genoffice/file-parse'

vi.mock('@genoffice/file-parse', () => ({ parseFileToText: vi.fn() }))

const parse = vi.mocked(parseFileToText)

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'file-index-ocr-'))
  parse.mockReset()
})

function pdfWithText(text: string): string {
  const path = join(dir, `doc-${Math.random().toString(36).slice(2)}.pdf`)
  writeFileSync(path, `%PDF-1.4 ${text}`)
  return path
}

describe('extractText OCR fallback', () => {
  it('uses the OCR hook for a PDF whose text layer is empty', async () => {
    const path = pdfWithText('')
    parse.mockResolvedValue({ ok: true, kind: 'text', text: '   ' })
    const ocrPdf = vi.fn(async () => 'Recognized scanned words')
    const r = await extractText(path, { ocrPdf })
    expect(ocrPdf).toHaveBeenCalledWith(path)
    expect(r).toEqual({ kind: 'text', text: 'Recognized scanned words' })
  })

  it('keeps the extraction when OCR returns nothing', async () => {
    const path = pdfWithText('')
    parse.mockResolvedValue({ ok: true, kind: 'text', text: '' })
    const r = await extractText(path, { ocrPdf: async () => null })
    expect(r).toEqual({ kind: 'text', text: '' })
  })

  it('degrades to the empty extraction when the OCR hook throws', async () => {
    const path = pdfWithText('')
    parse.mockResolvedValue({ ok: true, kind: 'text', text: ' ' })
    const r = await extractText(path, {
      ocrPdf: async () => {
        throw new Error('engine gone')
      },
    })
    expect(r).toEqual({ kind: 'text', text: ' ' })
  })

  it('never invokes OCR when the PDF already has a text layer', async () => {
    const path = pdfWithText('')
    parse.mockResolvedValue({
      ok: true,
      kind: 'text',
      text: 'real text extracted by pdfium, long enough to clear the scanned threshold',
    })
    const ocrPdf = vi.fn(async () => 'should not be called')
    await extractText(path, { ocrPdf })
    expect(ocrPdf).not.toHaveBeenCalled()
  })

  it('never invokes OCR for non-PDF files', async () => {
    const path = join(dir, 'notes.txt')
    writeFileSync(path, '')
    parse.mockResolvedValue({ ok: true, kind: 'text', text: '' })
    const ocrPdf = vi.fn(async () => 'nope')
    await extractText(path, { ocrPdf })
    expect(ocrPdf).not.toHaveBeenCalled()
  })

  it('keeps the richer of the two texts when the scan carries a stray character', async () => {
    const path = pdfWithText('')
    parse.mockResolvedValue({ ok: true, kind: 'text', text: 'X' })
    const ocrPdf = vi.fn(async () => '') // OCR recognized nothing usable
    const r = await extractText(path, { ocrPdf })
    expect(r).toEqual({ kind: 'text', text: 'X' })
  })

  it('threshold sanity: SCANNED_PDF_TEXT_CHARS stays tiny', () => {
    expect(SCANNED_PDF_TEXT_CHARS).toBeLessThanOrEqual(64)
  })
})

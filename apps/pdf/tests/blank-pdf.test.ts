import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { blankPdfBuffer } from '../src/main/blank-pdf'

describe('blankPdfBuffer', () => {
  it('produces a loadable single-page A4 document', async () => {
    const buf = await blankPdfBuffer()
    expect(buf.subarray(0, 5).toString()).toBe('%PDF-')
    // jsdom runs tests in another realm: hand pdf-lib a plain Uint8Array copy
    const doc = await PDFDocument.load(new Uint8Array(buf))
    expect(doc.getPageCount()).toBe(1)
    const { width, height } = doc.getPage(0).getSize()
    expect(width).toBeCloseTo(595.28, 1)
    expect(height).toBeCloseTo(841.89, 1)
  })
})

import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { buildPdfFixture, writeFixture } from './helpers/fixtures'

describe('parseFileToText: pdf', () => {
  it('extracts page text via pdfjs', async () => {
    const path = writeFixture('doc.pdf', buildPdfFixture('Hello PDF parsing'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('Hello PDF parsing')
  })

  it('fails gracefully on a corrupt pdf', async () => {
    const path = writeFixture('broken.pdf', Buffer.from('%PDF-1.4 garbage'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })

  // pdf.js scans the first 1024 bytes for %PDF-, and real-world fetches often
  // leave junk in front of it (HTTP header remnants, a stray CRLF). Sniffing
  // only bytes 0..3 rejected those files outright.
  it('accepts a pdf whose header follows leading junk', async () => {
    const junk = Buffer.from('\r\n\r\nGARBAGE-HEADER\n', 'utf8')
    const path = writeFixture(
      'junk.pdf',
      Buffer.concat([junk, Buffer.from(buildPdfFixture('Hello junk'))]),
    )
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('Hello junk')
  })
})

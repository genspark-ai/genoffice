import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { pdfToText } from '../src/pdf'
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

describe('pdfToText: item spacing', () => {
  const fixtureWithStream = (stream: string): Uint8Array => {
    const placeholder = 'BT /F1 24 Tf 72 720 Td (PLACEHOLDER) Tj ET'
    const src = new TextDecoder().decode(buildPdfFixture('PLACEHOLDER'))
    return new TextEncoder().encode(
      src
        .replace(placeholder, stream)
        .replace(`/Length ${placeholder.length}`, `/Length ${stream.length}`),
    )
  }

  // pdf.js only inserts a fake space when the next run advances to the right;
  // a run drawn to the left of the previous one on the same baseline was glued.
  it('separates same-line runs that are not adjacent horizontally', async () => {
    const text = await pdfToText(
      fixtureWithStream(
        'BT /F1 24 Tf 300 720 Td (World) Tj ET BT /F1 24 Tf 72 720 Td (Hello) Tj ET',
      ),
    )
    expect(text).toBe('World Hello')
  })

  it('keeps adjacent runs glued and does not double existing spaces', async () => {
    const glued = await pdfToText(fixtureWithStream('BT /F1 24 Tf 72 720 Td [(Hel) (lo)] TJ ET'))
    expect(glued).toBe('Hello')
    const spaced = await pdfToText(
      fixtureWithStream(
        'BT /F1 24 Tf 72 720 Td (Hello) Tj ET BT /F1 24 Tf 300 720 Td (World) Tj ET',
      ),
    )
    expect(spaced).toBe('Hello World')
  })

  it('breaks the line when the baseline changes', async () => {
    const text = await pdfToText(
      fixtureWithStream(
        'BT /F1 24 Tf 72 720 Td (Hello) Tj ET BT /F1 24 Tf 72 690 Td (World) Tj ET',
      ),
    )
    expect(text).toBe('Hello\nWorld')
  })
})

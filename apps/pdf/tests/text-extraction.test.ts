import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { readPdfText } from '../src/main/read-text'
import { verifyTextEdits } from '../src/main/text-edit'

/** one page, one Helvetica run per entry, at the given baseline */
async function fixture(lines: [string, number][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const [text, y] of lines) page.drawText(text, { x: 72, y, size: 14, font })
  return doc.save({ useObjectStreams: false })
}

describe('a hyphen at the end of a line', () => {
  it('reads back as the hyphen and break the page paints', async () => {
    // pdfium folds the line-end hyphen and the break into one codepoint: U+FFFE
    // from the textpage read, U+0002 from the per-object read. A reader asking
    // for the hyphen the page draws got a noncharacter instead.
    const bytes = await fixture([
      ['a hyphen-', 700],
      ['ated word', 680],
    ])
    const page = (await readPdfText(bytes)).pages[0]!
    expect(page.text).toBe('a hyphen-\nated word')
  })

  it('verifies a save of a line that ends in a hyphen', async () => {
    // verifyTextEdits compares the per-object text, which carries U+0002, against
    // the requested line — so a correctly drawn hyphen aborted the whole save
    // with "replacement text missing from saved output".
    const bytes = await fixture([
      ['a hyphen-', 700],
      ['ated word', 680],
    ])
    const failures = await verifyTextEdits(bytes, [
      { pageIndex: 0, newText: 'a hyphen-' },
      { pageIndex: 0, newText: 'ated word' },
    ])
    expect(failures).toEqual([])
  })

  it('still fails verification for text the page does not carry', async () => {
    const bytes = await fixture([
      ['a hyphen-', 700],
      ['ated word', 680],
    ])
    const failures = await verifyTextEdits(bytes, [{ pageIndex: 0, newText: 'absent line' }])
    expect(failures).toHaveLength(1)
    expect(failures[0]!.pageIndex).toBe(0)
    expect(failures[0]!.reason).toContain('replacement text missing from saved output')
  })

  it('leaves a plain line break and an in-word hyphen alone', async () => {
    const bytes = await fixture([
      ['well-known', 700],
      ['words', 680],
    ])
    const page = (await readPdfText(bytes)).pages[0]!
    expect(page.text).toBe('well-known\nwords')
  })
})

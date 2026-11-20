/**
 * Read-back verification on overlapping text objects (genoffice#2797): PDFium's
 * shared text page keeps only the first of two same-origin strings with a common
 * prefix, so the second replacement looked missing and aborted the save even
 * though the page had drawn it. The fallback reads each page-owned object with
 * the others temporarily removed.
 */
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { verifyTextEdits } from '../src/main/text-edit'

async function overlapping(texts: string[]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const text of texts) page.drawText(text, { x: 72, y: 500, size: 12, font })
  return doc.save({ useObjectStreams: false })
}

describe('verifying overlapping text objects', () => {
  it('accepts both replacements when the objects share a prefix', async () => {
    // the shared text page reads "abc" + "d", so "abd" was reported missing
    const bytes = await overlapping(['abc', 'abd'])
    expect(
      await verifyTextEdits(bytes, [
        { pageIndex: 0, newText: 'abc' },
        { pageIndex: 0, newText: 'abd' },
      ]),
    ).toEqual([])
  })

  it('accepts a third object that shares the same prefix', async () => {
    const bytes = await overlapping(['abc', 'abd', 'abe'])
    expect(
      await verifyTextEdits(
        bytes,
        ['abc', 'abd', 'abe'].map((newText) => ({ pageIndex: 0, newText })),
      ),
    ).toEqual([])
  })

  it('still refuses text no object drew', async () => {
    const bytes = await overlapping(['abc', 'abd'])
    const failures = await verifyTextEdits(bytes, [
      { pageIndex: 0, newText: 'abc' },
      { pageIndex: 0, newText: 'abz' },
    ])
    expect(failures).toHaveLength(1)
    expect(failures[0]!.reason).toContain('replacement text missing from saved output ("abz")')
  })

  it('leaves ordinary, non-overlapping text alone', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([595, 842])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('first line', { x: 72, y: 700, size: 12, font })
    page.drawText('second line', { x: 72, y: 660, size: 12, font })
    const bytes = await doc.save({ useObjectStreams: false })
    expect(
      await verifyTextEdits(bytes, [
        { pageIndex: 0, newText: 'first line' },
        { pageIndex: 0, newText: 'second line' },
      ]),
    ).toEqual([])
  })
})

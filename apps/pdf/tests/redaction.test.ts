import { describe, expect, it } from 'vitest'
import {
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  decodePDFRawStream,
  StandardFonts,
} from 'pdf-lib'
import { readPdfText } from '../src/main/read-text'
import { redactPdf } from '../src/main/redaction'

async function fixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 200])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('PUBLIC SECRET PUBLIC', { x: 20, y: 90, size: 18, font })
  // The visible text uses a font encoding, so add a decoded-stream sentinel that
  // proves this test's whole-object scanner can actually observe input evidence.
  page.node.addContentStream(
    doc.context.register(
      PDFRawStream.of(PDFDict.withContext(doc.context), Buffer.from('% SECRET-STREAM-SENTINEL\n')),
    ),
  )
  return doc.save({ useObjectStreams: false })
}

async function decodedStreams(bytes: Uint8Array): Promise<string> {
  const doc = await PDFDocument.load(bytes)
  let out = ''
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (object instanceof PDFRawStream)
      out += Buffer.from(decodePDFRawStream(object).decode()).toString('latin1')
  }
  return out
}

async function decodedObjectText(bytes: Uint8Array): Promise<string> {
  const doc = await PDFDocument.load(bytes)
  let out = ''
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    out +=
      object instanceof PDFRawStream
        ? Buffer.from(decodePDFRawStream(object).decode()).toString('latin1')
        : object.toString()
  }
  return out
}

async function pixelGridFixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([100, 100])
  const pixels = new Uint8Array(8 * 8 * 3)
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) pixels.set(x < 4 ? [255, 0, 0] : [0, 0, 255], (y * 8 + x) * 3)
  }
  const image = doc.context.register(
    PDFRawStream.of(
      doc.context.obj({
        Type: 'XObject',
        Subtype: 'Image',
        Width: 8,
        Height: 8,
        ColorSpace: 'DeviceRGB',
        BitsPerComponent: 8,
      }),
      pixels,
    ),
  )
  const name = page.node.newXObject('Image', image)
  page.node.addContentStream(
    doc.context.register(
      PDFRawStream.of(
        PDFDict.withContext(doc.context),
        Buffer.from(`q\n80 0 0 80 10 10 cm\n${name.toString()} Do\nQ\n`),
      ),
    ),
  )
  return doc.save({ useObjectStreams: false })
}

async function embeddedImagePixels(bytes: Uint8Array): Promise<Uint8Array[]> {
  const doc = await PDFDocument.load(bytes)
  const images: Uint8Array[] = []
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (
      object instanceof PDFRawStream &&
      object.dict.get(PDFName.of('Subtype'))?.toString() === '/Image'
    ) {
      images.push(decodePDFRawStream(object).decode())
    }
  }
  return images
}

describe('native PDF redaction', () => {
  it('permanently removes a marked text region using real PDFium', async () => {
    const before = await fixture()
    expect(await decodedStreams(before)).toContain('SECRET-STREAM-SENTINEL')
    const out = await redactPdf(before, [{ pageIndex: 0, rect: [78, 84, 145, 112] }])

    const text = await readPdfText(out)
    expect(text.pages[0]?.text).toContain('PUBLIC')
    expect(text.pages[0]?.text).not.toContain('SECRET')
    // The redactor rewrites the content stream rather than leaving a visual cover over
    // the original string. The sentinel must also be absent from decoded stream bytes.
    expect(await decodedStreams(out)).not.toContain('SECRET')
  })

  it('removes selected embedded image pixels without restoring an original image stream', async () => {
    const out = await redactPdf(await pixelGridFixture(), [
      { pageIndex: 0, rect: [10, 10, 50, 90] },
    ])
    const images = await embeddedImagePixels(out)
    expect(images).toHaveLength(1)
    expect(images[0]).toHaveLength(8 * 8 * 3)
    for (let y = 0; y < 8; y++) {
      for (let x = 0; x < 8; x++) {
        const actual = [...images[0]!.slice((y * 8 + x) * 3, (y * 8 + x + 1) * 3)]
        expect(actual).toEqual(x < 4 ? [255, 255, 255] : [0, 0, 255])
      }
    }
  })

  it('removes each of two independently marked text regions', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 200])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('SECRET-ONE', { x: 20, y: 130, size: 18, font })
    page.drawText('SECRET-TWO', { x: 20, y: 60, size: 18, font })
    page.drawText('PUBLIC', { x: 20, y: 20, size: 18, font })
    const out = await redactPdf(await doc.save({ useObjectStreams: false }), [
      { pageIndex: 0, rect: [18, 125, 140, 155] },
      { pageIndex: 0, rect: [18, 55, 140, 85] },
    ])
    const text = await readPdfText(out)
    expect(text.pages[0]?.text).not.toContain('SECRET-ONE')
    expect(text.pages[0]?.text).not.toContain('SECRET-TWO')
    expect(text.pages[0]?.text).toContain('PUBLIC')
  })

  it('removes marked-content ActualText from extraction and decoded streams', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 200])
    const font = await doc.embedFont(StandardFonts.Helvetica)
    page.drawText('PUBLIC', { x: 20, y: 40, size: 18, font })
    page.node.addContentStream(
      doc.context.register(
        PDFRawStream.of(
          PDFDict.withContext(doc.context),
          Buffer.from(
            `/Span << /ActualText (SECRET-ACTUAL) >> BDC\nBT /${font.name} 18 Tf 1 0 0 1 20 100 Tm (SECRET-ACTUAL) Tj ET\nEMC\n`,
          ),
        ),
      ),
    )
    const before = await doc.save({ useObjectStreams: false })
    expect(await readPdfText(before)).toMatchObject({
      pages: [{ text: expect.stringContaining('SECRET-ACTUAL') }],
    })
    expect(await decodedStreams(before)).toContain('SECRET-ACTUAL')
    const out = await redactPdf(before, [{ pageIndex: 0, rect: [18, 95, 180, 125] }])
    expect((await readPdfText(out)).pages[0]?.text).not.toContain('SECRET-ACTUAL')
    expect(await decodedStreams(out)).not.toContain('SECRET-ACTUAL')
  })

  it('removes an intersecting annotation without leaving its contents in saved objects', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 200])
    const annotation = doc.context.register(
      doc.context.obj({
        Type: 'Annot',
        Subtype: 'Text',
        Rect: [20, 80, 180, 110],
        Contents: 'SECRET-ANNOT',
      }),
    )
    page.node.set(PDFName.of('Annots'), doc.context.obj([annotation]))
    const before = await doc.save({ useObjectStreams: false })
    expect(await decodedObjectText(before)).toContain('SECRET-ANNOT')
    const out = await redactPdf(before, [{ pageIndex: 0, rect: [20, 80, 180, 110] }])
    expect(await decodedObjectText(out)).not.toContain('SECRET-ANNOT')
  })

  it('rejects a region over a Widget because PDFium 2.15.1 retains its field value', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 200])
    const widget = doc.context.register(
      doc.context.obj({
        Type: 'Annot',
        Subtype: 'Widget',
        Rect: [20, 80, 180, 110],
        FT: 'Tx',
        T: 'secret-field',
        V: 'SECRET-FORM-VALUE',
      }),
    )
    page.node.set(PDFName.of('Annots'), doc.context.obj([widget]))
    doc.catalog.set(
      PDFName.of('AcroForm'),
      doc.context.register(doc.context.obj({ Fields: [widget] })),
    )
    await expect(
      redactPdf(await doc.save({ useObjectStreams: false }), [
        { pageIndex: 0, rect: [20, 80, 180, 110] },
      ]),
    ).rejects.toThrow('form field')
  })

  it('fails closed when the selected image bytes are shared', async () => {
    const doc = await PDFDocument.create()
    const png = await doc.embedPng(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    )
    for (const y of [20, 70]) {
      const page = doc.getPageCount() === 0 ? doc.addPage([120, 120]) : doc.getPage(0)
      page.drawImage(png, { x: 20, y, width: 30, height: 30 })
    }
    await expect(
      redactPdf(await doc.save({ useObjectStreams: false }), [
        { pageIndex: 0, rect: [20, 20, 50, 50] },
      ]),
    ).rejects.toThrow('shared image')
  })

  it('rejects malformed or out-of-range regions before mutating bytes', async () => {
    const before = await fixture()
    await expect(redactPdf(before, [{ pageIndex: 9, rect: [1, 1, 2, 2] }])).rejects.toThrow(
      'page index',
    )
    await expect(redactPdf(before, [{ pageIndex: 0, rect: [1, 1, 1, 2] }])).rejects.toThrow(
      'rectangle',
    )
  })
})

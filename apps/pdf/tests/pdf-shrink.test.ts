import { describe, expect, it } from 'vitest'
import { deflateSync } from 'node:zlib'
import { PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { resamplePixels, shrinkDocument } from '../src/main/pdf-shrink'

describe('resamplePixels', () => {
  it('nearest-samples an RGB buffer to the target size', () => {
    // 2x1 red | blue → 4x2: each quadrant copies its source pixel
    const out = resamplePixels(new Uint8Array([255, 0, 0, 0, 0, 255]), 2, 1, 3, 4, 2)
    expect(out).toEqual(
      new Uint8Array([
        255, 0, 0, 255, 0, 0, 0, 0, 255, 0, 0, 255, 255, 0, 0, 255, 0, 0, 0, 0, 255, 0, 0, 255,
      ]),
    )
  })
})

describe('shrinkDocument', () => {
  it('re-deflates weakly compressed streams and reports the count', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([200, 200])
    const payload = Buffer.from('repeated content '.repeat(500))
    const weak = deflateSync(payload, { level: 1 })
    const stream = PDFRawStream.of(
      doc.context.obj({ Length: weak.length, Filter: 'FlateDecode' }),
      new Uint8Array(weak),
    )
    doc.context.assign(doc.context.nextRef(), stream)
    const before = (await doc.save()).length
    const result = shrinkDocument(doc)
    expect(result.streamsRecompressed).toBe(1)
    const after = await doc.save()
    expect(after.length).toBeLessThan(before)
  })

  it('downsamples an oversized Flate image and rewrites its dimensions', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const w = 4000
    const h = 100
    const pixels = new Uint8Array(w * h * 3).fill(200)
    const dict = doc.context.obj({
      Type: 'XObject',
      Subtype: 'Image',
      Width: w,
      Height: h,
      ColorSpace: 'DeviceRGB',
      BitsPerComponent: 8,
      Filter: 'FlateDecode',
      Length: 1,
    })
    const ref = doc.context.nextRef()
    doc.context.assign(ref, PDFRawStream.of(dict, deflateSync(pixels)))
    // hang it off the page resources so it counts as content
    doc.getPages()[0]!.node.set(PDFName.of('XObject'), doc.context.obj({ Im0: ref }))
    const result = shrinkDocument(doc, 2480)
    expect(result.imagesDownsampled).toBe(1)
    const saved = await doc.save()
    const back = await PDFDocument.load(saved, { updateMetadata: false })
    const fresh = back.context.lookup(ref)
    expect(String((fresh as PDFRawStream).dict.get(PDFName.of('Width')))).toBe('2480')
  })

  it('leaves predictor streams and non-Flate data byte-identical', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([200, 200])
    const payload = Buffer.from('predictor data '.repeat(300))
    const packed = deflateSync(payload, { level: 1 })
    const dict = doc.context.obj({
      Length: packed.length,
      Filter: 'FlateDecode',
      DecodeParms: doc.context.obj({ Predictor: 12 }),
    })
    const ref = doc.context.nextRef()
    doc.context.assign(ref, PDFRawStream.of(dict, new Uint8Array(packed)))
    const before = await doc.save()
    const result = shrinkDocument(doc)
    expect(result.streamsRecompressed).toBe(0)
    const after = await doc.save()
    expect(after.length).toBe(before.length)
  })

  it('skips an oversized ICCBased image without throwing', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    // an array colour space, as photo scans carry: [/ICCBased <stream ref>]
    const icc = doc.context.nextRef()
    doc.context.assign(
      icc,
      PDFRawStream.of(doc.context.obj({ N: 3, Length: 2 }), new Uint8Array([0, 0])),
    )
    const w = 4000
    const pixels = new Uint8Array(w * 100 * 3).fill(200)
    const dict = doc.context.obj({
      Type: 'XObject',
      Subtype: 'Image',
      Width: w,
      Height: 100,
      ColorSpace: doc.context.obj([PDFName.of('ICCBased'), icc]),
      BitsPerComponent: 8,
      Filter: 'FlateDecode',
      Length: 1,
    })
    const ref = doc.context.nextRef()
    doc.context.assign(ref, PDFRawStream.of(dict, deflateSync(pixels)))
    // lookupMaybe on the array used to raise UnexpectedObjectTypeError and
    // abort the whole run; the array is now simply not a device colour space
    const result = shrinkDocument(doc)
    expect(result.imagesDownsampled).toBe(0)
  })

  it('keeps an /SMask pair intact when only one side could shrink', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const w = 4000
    const h = 1000
    // base: DCT (never touched) — the mask's dimensions must follow it
    const baseDict = doc.context.obj({
      Type: 'XObject',
      Subtype: 'Image',
      Width: w,
      Height: h,
      ColorSpace: 'DeviceRGB',
      BitsPerComponent: 8,
      Filter: 'DCTDecode',
      Length: 4,
    })
    const baseRef = doc.context.nextRef()
    doc.context.assign(baseRef, PDFRawStream.of(baseDict, new Uint8Array([1, 2, 3, 4])))
    // mask: oversized Flate gray — a candidate on its own, but its base is not
    const maskPixels = new Uint8Array(w * h).fill(128)
    const maskRef = doc.context.nextRef()
    doc.context.assign(
      maskRef,
      PDFRawStream.of(
        doc.context.obj({
          Type: 'XObject',
          Subtype: 'Image',
          Width: w,
          Height: h,
          ColorSpace: 'DeviceGray',
          BitsPerComponent: 8,
          Filter: 'FlateDecode',
          Length: 1,
        }),
        deflateSync(maskPixels),
      ),
    )
    baseDict.set(PDFName.of('SMask'), maskRef)
    const result = shrinkDocument(doc)
    expect(result.imagesDownsampled).toBe(0)
    expect(
      String((doc.context.lookup(maskRef) as PDFRawStream).dict.get(PDFName.of('Width'))),
    ).toBe(String(w))
  })

  it('shrinks an /SMask pair together so the dimensions still match', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const w = 4000
    const h = 1000
    const baseRef = doc.context.nextRef()
    doc.context.assign(
      baseRef,
      PDFRawStream.of(
        doc.context.obj({
          Type: 'XObject',
          Subtype: 'Image',
          Width: w,
          Height: h,
          ColorSpace: 'DeviceRGB',
          BitsPerComponent: 8,
          Filter: 'FlateDecode',
          Length: 1,
        }),
        deflateSync(new Uint8Array(w * h * 3).fill(200)),
      ),
    )
    const maskRef = doc.context.nextRef()
    const maskDict = doc.context.obj({
      Type: 'XObject',
      Subtype: 'Image',
      Width: w,
      Height: h,
      ColorSpace: 'DeviceGray',
      BitsPerComponent: 8,
      Filter: 'FlateDecode',
      Length: 1,
    })
    doc.context.assign(maskRef, PDFRawStream.of(maskDict, deflateSync(new Uint8Array(w * h))))
    ;(doc.context.lookup(baseRef) as PDFRawStream).dict.set(PDFName.of('SMask'), maskRef)
    // a mask with /Matte must never shrink apart from its base
    maskDict.set(PDFName.of('Matte'), doc.context.obj([1, 0, 0, -1, 0, h]))
    const withMatte = shrinkDocument(doc)
    expect(withMatte.imagesDownsampled).toBe(0)
    // without the Matte the pair shrinks to the same new dimensions
    maskDict.delete(PDFName.of('Matte'))
    const paired = shrinkDocument(doc)
    expect(paired.imagesDownsampled).toBe(2)
    const dims = (r: unknown) => {
      const d = (doc.context.lookup(r) as PDFRawStream).dict
      return [String(d.get(PDFName.of('Width'))), String(d.get(PDFName.of('Height')))]
    }
    expect(dims(baseRef)).toEqual(dims(maskRef))
  })
})

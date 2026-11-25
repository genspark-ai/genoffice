/** Reduce File Size in the host: what is kept of a redrawn picture, with the canvas stood in for. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { reducePictures } from '../src/renderer/reduce-pictures'

// the TIFF decoder needs a real canvas: here a TIFF decodes to a PNG that names its own size
const tiffAsPng = vi.hoisted(() => ({ decode: (_tiff: Uint8Array): Uint8Array | null => null }))
vi.mock('@genoffice/docx-engine', async (original) => ({
  ...(await original<typeof import('@genoffice/docx-engine')>()),
  tiffToDataUrlAsync: async (bytes: ArrayBuffer | Uint8Array) => {
    const png = tiffAsPng.decode(new Uint8Array(bytes as ArrayBuffer))
    return png && `data:image/png;base64,${Buffer.from(png).toString('base64')}`
  },
}))

/** the header of a PNG of w x h, followed by `size` bytes that stand for its pixels */
function png(w: number, h: number, size: number): Uint8Array {
  const out = new Uint8Array(33 + size)
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52])
  new DataView(out.buffer).setUint32(16, w)
  new DataView(out.buffer).setUint32(20, h)
  out.fill(7, 33)
  return out
}

const picture = (crop: string) =>
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="p"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  '<pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="rIdPic"/>${crop}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const docx = (crop: string, stored: Uint8Array) =>
  buildDocx({
    bodyXml: picture(crop),
    extraRels:
      '<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.png"/>',
    binaryParts: [
      {
        path: 'word/media/pic.png',
        base64: Buffer.from(stored).toString('base64'),
        extension: 'png',
        contentType: 'image/png',
      },
    ],
  })

/** a canvas whose encoder writes `encoded` bytes, and that records what it was asked to draw */
function standInCanvas(encoded: number) {
  const drawn: number[][] = []
  vi.stubGlobal('createImageBitmap', async (_blob: Blob, ...rest: unknown[]) => {
    drawn.push(rest.filter((v): v is number => typeof v === 'number'))
    return { close() {} }
  })
  vi.stubGlobal(
    'OffscreenCanvas',
    class {
      getContext() {
        return { drawImage() {} }
      }
      async convertToBlob({ type }: { type: string }) {
        return new Blob([new Uint8Array(encoded).fill(9)], { type })
      }
    },
  )
  return drawn
}

async function saved(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  return {
    part: await zip.file('word/media/pic.png')!.async('uint8array'),
    xml: await zip.file('word/document.xml')!.async('string'),
  }
}

afterEach(() => vi.unstubAllGlobals())

describe('reducing the pictures of a document', () => {
  const CROP = '<a:srcRect l="50000" t="25000"/>'

  it('cuts a cropped picture down even when the cut picture takes more bytes', async () => {
    const stored = png(800, 600, 100)
    const drawn = standInCanvas(5000)
    const { bytes, outcome } = await reducePictures(await docx(CROP, stored), {
      ppi: null,
      deleteCropped: true,
    })
    expect(outcome).toMatchObject({ reduced: 1, smallEnough: 0, bytesAfter: 5000 })
    expect(outcome.bytesBefore).toBeLessThan(outcome.bytesAfter)
    // only the right half of the lower three quarters was drawn
    expect(drawn).toEqual([[400, 150, 400, 450]])
    const out = await saved(bytes!)
    expect(out.part.length).toBe(5000)
    expect(out.xml).not.toContain('a:srcRect')
    expect(out.xml).toContain('<wp:extent cx="952500" cy="952500"/>')
  })

  it('keeps a picture as it is stored when resampling alone would not save bytes', async () => {
    const stored = png(800, 600, 100)
    standInCanvas(5000)
    const whole = await reducePictures(await docx('', stored), { ppi: 96, deleteCropped: true })
    expect(whole.bytes).toBeNull()
    expect(whole.outcome).toMatchObject({ reduced: 0, smallEnough: 1 })
    // and a crop that was not asked to be deleted stays a crop
    const kept = await reducePictures(await docx(CROP, stored), { ppi: 96, deleteCropped: false })
    expect(kept.bytes).toBeNull()
  })

  it('stores a BMP again as a PNG under a PNG name, at the size it has when that is all it needs', async () => {
    const bmp = new Uint8Array(54 + 4000)
    bmp.set([0x42, 0x4d])
    const view = new DataView(bmp.buffer)
    view.setUint32(14, 40, true)
    view.setInt32(18, 80, true)
    view.setInt32(22, 60, true)
    const source = await buildDocx({
      bodyXml: picture(''),
      extraRels:
        '<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.bmp"/>',
      binaryParts: [
        {
          path: 'word/media/pic.bmp',
          base64: Buffer.from(bmp).toString('base64'),
          extension: 'bmp',
          contentType: 'image/bmp',
        },
      ],
    })
    const drawn = standInCanvas(500)
    const { bytes, outcome } = await reducePictures(source, { ppi: null, deleteCropped: false })
    expect(outcome).toMatchObject({ reduced: 1, untouched: 0, bytesBefore: 4054, bytesAfter: 500 })
    expect(drawn).toEqual([[]])
    const zip = await JSZip.loadAsync(bytes!)
    expect(zip.file('word/media/pic.bmp')).toBeNull()
    const stored = await zip.file('word/media/pic.png')!.async('uint8array')
    expect(stored.length).toBe(500)
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
      'Target="media/pic.png"',
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="png"')
    expect(await zip.file('word/document.xml')!.async('string')).toContain(
      '<wp:extent cx="952500" cy="952500"/>',
    )
  })

  it('leaves an animated GIF as it is', async () => {
    const frame = [0x2c, 0, 0, 0, 0, 8, 0, 8, 0, 0, 2, 2, 0x4c, 0x01, 0]
    const gif = Uint8Array.from([
      ...[0x47, 0x49, 0x46, 0x38, 0x39, 0x61],
      ...[8, 0, 8, 0, 0, 0, 0],
      ...frame,
      ...frame,
      0x3b,
    ])
    const source = await buildDocx({
      bodyXml: picture(''),
      extraRels:
        '<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.gif"/>',
      binaryParts: [
        {
          path: 'word/media/pic.gif',
          base64: Buffer.from(gif).toString('base64'),
          extension: 'gif',
          contentType: 'image/gif',
        },
      ],
    })
    standInCanvas(5)
    const { bytes, outcome } = await reducePictures(source, { ppi: 96, deleteCropped: false })
    expect(bytes).toBeNull()
    expect(outcome).toMatchObject({ reduced: 0, untouched: 1 })
  })

  it('finds a selected TIFF by the PNG the editor shows it as', async () => {
    const tiff = new Uint8Array(6000)
    tiff.set([0x49, 0x49, 0x2a, 0x00])
    const shown = png(800, 600, 3000)
    tiffAsPng.decode = (bytes) => (bytes.length === tiff.length ? shown : null)
    const source = await buildDocx({
      bodyXml: picture(''),
      extraRels:
        '<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.tiff"/>',
      binaryParts: [
        {
          path: 'word/media/pic.tiff',
          base64: Buffer.from(tiff).toString('base64'),
          extension: 'tiff',
          contentType: 'image/tiff',
        },
      ],
    })
    standInCanvas(400)
    const dataUrl = (bytes: Uint8Array) =>
      `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`
    const picked = await reducePictures(source, {
      ppi: 96,
      deleteCropped: false,
      only: [dataUrl(shown)],
    })
    expect(picked.outcome).toMatchObject({ reduced: 1, bytesBefore: 6000, bytesAfter: 400 })
    const zip = await JSZip.loadAsync(picked.bytes!)
    expect(zip.file('word/media/pic.tiff')).toBeNull()
    expect((await zip.file('word/media/pic.png')!.async('uint8array')).length).toBe(400)

    // another picture selected: the TIFF is not one of them
    const other = await reducePictures(source, {
      ppi: 96,
      deleteCropped: false,
      only: [dataUrl(png(10, 10, 20))],
    })
    expect(other.bytes).toBeNull()
    expect(other.outcome).toMatchObject({ reduced: 0, untouched: 0 })
  })

  it('leaves a TIFF undecoded while no PNG is among the selection', async () => {
    const tiff = new Uint8Array(6000)
    tiff.set([0x49, 0x49, 0x2a, 0x00])
    let decodes = 0
    tiffAsPng.decode = () => {
      decodes++
      return png(800, 600, 3000)
    }
    const source = await buildDocx({
      bodyXml: picture(''),
      extraRels:
        '<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.tiff"/>',
      binaryParts: [
        {
          path: 'word/media/pic.tiff',
          base64: Buffer.from(tiff).toString('base64'),
          extension: 'tiff',
          contentType: 'image/tiff',
        },
      ],
    })
    standInCanvas(400)
    const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10]).toString('base64')
    const { bytes, outcome } = await reducePictures(source, {
      ppi: 96,
      deleteCropped: false,
      only: ['genoffice-media://doc/word/media/other.jpeg', `data:image/jpeg;base64,${jpeg}`],
    })
    expect(decodes).toBe(0)
    expect(bytes).toBeNull()
    expect(outcome).toMatchObject({ reduced: 0, untouched: 0 })
  })
})

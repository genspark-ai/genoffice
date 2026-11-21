/** Reduce File Size in the host: what is kept of a redrawn picture, with the canvas stood in for. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { reducePictures } from '../src/renderer/reduce-pictures'

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
})

import { describe, expect, it, vi } from 'vitest'
import { parseDocx } from '../src/index'
import { metafileToDataUrl } from '../src/metafile'
import { convertEmfToDataUrl, convertWmfToDataUrl } from '../src/vendor/emf-converter/index.mjs'
import { buildDocx } from './helpers/build-docx'

// Canvas is unavailable under vitest; mock the converter to test the wiring.
vi.mock('../src/vendor/emf-converter/index.mjs', () => ({
  convertEmfToDataUrl: vi.fn(async () => 'data:image/png;base64,EMFPNG'),
  convertWmfToDataUrl: vi.fn(async () => 'data:image/png;base64,WMFPNG'),
}))

const EMF_BASE64 = Buffer.from('not-a-real-emf-payload').toString('base64')

const EMF_REL =
  '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>'

const EMF_BINARY_PART = {
  path: 'word/media/image1.emf',
  base64: EMF_BASE64,
  extension: 'emf',
  contentType: 'image/x-emf',
}

const DRAWING_XML =
  '<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing>'

describe('metafileToDataUrl routing', () => {
  it('routes emf mimes to convertEmfToDataUrl with dpiScale 2', async () => {
    for (const mime of ['image/emf', 'image/x-emf']) {
      expect(await metafileToDataUrl(new Uint8Array([1, 2, 3]), mime)).toBe(
        'data:image/png;base64,EMFPNG',
      )
    }
    expect(convertEmfToDataUrl).toHaveBeenLastCalledWith(expect.any(ArrayBuffer), { dpiScale: 2 })
  })

  it('routes wmf mimes to convertWmfToDataUrl', async () => {
    for (const mime of ['image/wmf', 'image/x-wmf']) {
      expect(await metafileToDataUrl(new Uint8Array([1, 2, 3]), mime)).toBe(
        'data:image/png;base64,WMFPNG',
      )
    }
    expect(convertWmfToDataUrl).toHaveBeenLastCalledWith(expect.any(ArrayBuffer), { dpiScale: 2 })
  })
})

describe('parseDocx with emf media', () => {
  it('converts an inline emf picture to an image block', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:r>${DRAWING_XML}</w:r></w:p>`,
        extraRels: EMF_REL,
        binaryParts: [EMF_BINARY_PART],
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.brokenImage).toBeUndefined()
    expect(block.imageDataUrl).toBe('data:image/png;base64,EMFPNG')
    expect(block.imageWidthPx).toBe(96)
    expect(block.imageHeightPx).toBe(48)
  })

  it('resolves an opaque .bin part via the [Content_Types].xml fallback', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:r>${DRAWING_XML}</w:r></w:p>`,
        extraRels:
          '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bin"/>',
        binaryParts: [
          {
            path: 'word/media/image1.bin',
            base64: EMF_BASE64,
            extension: 'bin',
            contentType: 'image/x-emf',
          },
        ],
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageDataUrl).toBe('data:image/png;base64,EMFPNG')
  })

  it('converts an emf picture inside a table cell', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:tbl><w:tr><w:tc><w:p><w:r>${DRAWING_XML}</w:r></w:p></w:tc></w:tr></w:tbl>`,
        extraRels: EMF_REL,
        binaryParts: [EMF_BINARY_PART],
      }),
    )
    const cell = doc.blocks[0].table?.rows[0]?.[0]
    const run = cell?.richParas?.[0]?.runs.find((r) => r.image)
    expect(run?.image?.dataUrl).toBe('data:image/png;base64,EMFPNG')
  })

  it('keeps the empty-frame degrade when conversion fails', async () => {
    vi.mocked(convertEmfToDataUrl).mockResolvedValue(null)
    try {
      const doc = await parseDocx(
        await buildDocx({
          bodyXml: `<w:p><w:r>${DRAWING_XML}</w:r></w:p>`,
          extraRels: EMF_REL,
          binaryParts: [EMF_BINARY_PART],
        }),
      )
      const block = doc.blocks[0]
      expect(block.type).toBe('passthrough')
      expect(block.brokenImage).toBe(true)
    } finally {
      vi.mocked(convertEmfToDataUrl).mockResolvedValue('data:image/png;base64,EMFPNG')
    }
  })

  it('surfaces a converted OLE preview picture', async () => {
    const oleXml =
      '<w:p><w:r><w:object>' +
      '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:96pt;height:48pt">' +
      '<v:imagedata r:id="rId10"/></v:shape>' +
      '<o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ProgID="Excel.Sheet.12"/>' +
      '</w:object></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: oleXml, extraRels: EMF_REL, binaryParts: [EMF_BINARY_PART] }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.oleProgId).toBe('Excel.Sheet.12')
    expect(block.imageDataUrl).toBe('data:image/png;base64,EMFPNG')
  })
})

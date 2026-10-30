import { readFileSync } from 'node:fs'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'
import { installFakeCanvas } from './helpers/fake-canvas'

/**
 * MathType equations: Word stores them as an OLE embed
 * (ProgID "Equation.DSMT4") whose appearance comes from the metafile the
 * package carries next to it. Without that preview the equation has nothing to
 * draw, so the parse must turn the metafile into a picture.
 */
const SHAPETYPE =
  '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" ' +
  'path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
  '<v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
  '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'

function oleEquationXml(previewRid: string, oleRid: string): string {
  return `<w:p>${oleEquationRun(previewRid, oleRid)}</w:p>`
}

function oleEquationRun(previewRid: string, oleRid: string): string {
  return (
    '<w:r>' +
    '<w:object w:dxaOrig="1440" w:dyaOrig="720">' +
    SHAPETYPE +
    '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole="">' +
    `<v:imagedata r:id="${previewRid}" o:title=""/>` +
    '</v:shape>' +
    '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" ' +
    `DrawAspect="Content" ObjectID="_1700000000" r:id="${oleRid}"/>` +
    '</w:object></w:r>'
  )
}

const RELS =
  '<Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.%EXT%"/>' +
  '<Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>'

function fixture(name: string): string {
  return readFileSync(new URL(`./fixtures/${name}`, import.meta.url)).toString('base64')
}

let restore: () => void

beforeAll(() => {
  restore = installFakeCanvas()
})

afterAll(() => {
  restore()
})

async function parseEquation(preview: string, ext: string) {
  const bytes = await buildDocx({
    bodyXml: oleEquationXml('rId30', 'rId31'),
    extraRels: RELS.replace('%EXT%', ext),
    binaryParts: [
      {
        path: `word/media/image1.${ext}`,
        base64: preview,
        extension: ext,
        contentType: ext === 'png' ? 'image/png' : `image/x-${ext}`,
      },
      {
        path: 'word/embeddings/oleObject1.bin',
        base64: 'AAECAwQ=',
        extension: 'bin',
        contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
      },
    ],
  })
  return parseDocx(bytes)
}

describe('MathType equation previews', () => {
  it('turns an EMF preview into a picture', async () => {
    const parsed = await parseEquation(fixture('ole-caption.emf'), 'emf')
    const block = parsed.blocks[0]
    expect(block?.oleProgId).toBe('Equation.DSMT4')
    expect(block?.imageDataUrl).toMatch(/^data:image\/png;base64,/)
    // the shape declares 72x36 pt → 96x48 px
    expect(block?.imageWidthPx).toBe(96)
    expect(block?.imageHeightPx).toBe(48)
  })

  it('turns a WMF preview into a picture', async () => {
    const parsed = await parseEquation(fixture('ole-icon.wmf'), 'wmf')
    const block = parsed.blocks[0]
    expect(block?.oleProgId).toBe('Equation.DSMT4')
    expect(block?.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('shows the equation when it shares the paragraph with text', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t xml:space="preserve">Formula: </w:t></w:r>' +
        oleEquationRun('rId30', 'rId31') +
        '<w:r><w:t xml:space="preserve"> ends the line.</w:t></w:r></w:p>',
      extraRels: RELS.replace('%EXT%', 'emf'),
      binaryParts: [
        {
          path: 'word/media/image1.emf',
          base64: fixture('ole-caption.emf'),
          extension: 'emf',
          contentType: 'image/x-emf',
        },
        {
          path: 'word/embeddings/oleObject1.bin',
          base64: 'AAECAwQ=',
          extension: 'bin',
          contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
        },
      ],
    })
    const parsed = await parseDocx(bytes)
    const block = parsed.blocks[0]
    const imageRun = block?.runs?.find((r) => r.image)
    expect(imageRun, 'the equation run carries no picture').toBeTruthy()
    expect(imageRun?.image?.dataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('keeps the equation and its preview through a save', async () => {
    const parsed = await parseEquation(fixture('ole-caption.emf'), 'emf')
    const bytes = await saveDocx(parsed, [
      { kind: 'original', docxIndex: parsed.blocks[0]!.docxIndex! },
    ])
    const reparsed = await parseDocx(bytes)
    const block = reparsed.blocks[0]
    // the embed itself: ProgID, the shape and the preview relationship
    expect(block?.originalXml).toContain('ProgID="Equation.DSMT4"')
    expect(block?.originalXml).toContain('<v:imagedata')
    // and the preview picture still resolves to a rendered picture
    expect(block?.oleProgId).toBe('Equation.DSMT4')
    expect(block?.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })
})

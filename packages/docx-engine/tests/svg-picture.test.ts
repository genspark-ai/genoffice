/**
 * A picture inserted from a vector is saved the way Word saves an icon: the
 * blip embeds a bitmap every reader can show and names the SVG part in its
 * Office 2016 svgBlip extension, which Word prefers.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  inlinePictureDrawingXml,
  NEW_PICTURE_SVG_SUFFIX,
  NEW_PICTURE_TOKEN_PREFIX,
  parseDocx,
  saveDocx,
  type NewImage,
} from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'
const SVG_BASE64 = Buffer.from(SVG).toString('base64')
const OTHER_SVG_BASE64 = Buffer.from(SVG.replace('M2 2', 'M3 3')).toString('base64')

const icon = (extra: Partial<NewImage> = {}): NewImage => ({
  base64: TINY_PNG_BASE64,
  mime: 'image/png',
  svgBase64: SVG_BASE64,
  widthPx: 96,
  heightPx: 96,
  ...extra,
})

async function open(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  const documentXml = await zip.file('word/document.xml')!.async('string')
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  const types = await zip.file('[Content_Types].xml')!.async('string')
  const targetOf = (rId: string | undefined) =>
    new RegExp(`Id="${rId}"[^>]*Target="(media/[^"]+)"`).exec(rels)?.[1]
  const media = Object.keys(zip.files).filter(
    (f) => f.startsWith('word/media/') && !zip.files[f].dir,
  )
  return { zip, documentXml, rels, types, targetOf, media }
}

async function saveImages(images: NewImage[]) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>' }))
  const blocks = [
    { kind: 'original' as const, docxIndex: 0 },
    ...images.map((image) => ({ kind: 'image' as const, image })),
  ]
  return saveDocx(parsed, blocks)
}

describe('a picture with an SVG original', () => {
  it('writes the SVG part beside its bitmap fallback and declares both', async () => {
    const { zip, documentXml, types, targetOf } = await open(await saveImages([icon()]))
    const blip =
      /<a:blip r:embed="(rId\d+)"><a:extLst><a:ext uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"><asvg:svgBlip xmlns:asvg="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2016\/SVG\/main" r:embed="(rId\d+)"\/><\/a:ext><\/a:extLst><\/a:blip>/.exec(
        documentXml,
      )
    expect(blip).toBeTruthy()
    const png = targetOf(blip![1])
    const svg = targetOf(blip![2])
    expect(png).toMatch(/^media\/aidocs\d+\.png$/)
    expect(svg).toMatch(/^media\/aidocs\d+\.svg$/)
    expect(await zip.file(`word/${svg}`)!.async('string')).toBe(SVG)
    expect(await zip.file(`word/${png}`)!.async('base64')).toBe(TINY_PNG_BASE64)
    expect(types).toContain('<Default Extension="svg" ContentType="image/svg+xml"/>')
    expect(types).toMatch(/<Default Extension="png" ContentType="image\/png"\/>/)
    expect(documentXml).toContain('<wp:extent cx="914400" cy="914400"/>')
  })

  it('shares one SVG part between copies and keeps different drawings apart', async () => {
    const same = await open(await saveImages([icon(), icon()]))
    expect(same.media.filter((f) => f.endsWith('.svg'))).toHaveLength(1)
    expect(same.media.filter((f) => f.endsWith('.png'))).toHaveLength(1)
    const other = await open(await saveImages([icon(), icon({ svgBase64: OTHER_SVG_BASE64 })]))
    expect(other.media.filter((f) => f.endsWith('.svg'))).toHaveLength(2)
    expect(new Set(other.media).size).toBe(other.media.length)
  })

  it('keeps the SVG when the picture floats', async () => {
    const { documentXml } = await open(await saveImages([icon({ wrap: 'square-left' })]))
    expect(documentXml).toContain('<wp:anchor')
    expect(documentXml).toMatch(/<asvg:svgBlip [^>]*r:embed="rId\d+"\/>/)
  })

  it('reads back as the picture, shown from its bitmap, and saves again unchanged', async () => {
    const first = await saveImages([icon()])
    const parsed = await parseDocx(first)
    const picture = parsed.blocks.find((b) => b.type === 'image')
    expect(picture?.imageDataUrl).toBe(`data:image/png;base64,${TINY_PNG_BASE64}`)
    const again = await open(
      await saveDocx(
        parsed,
        parsed.blocks.map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
      ),
    )
    expect(again.documentXml).toMatch(/<asvg:svgBlip [^>]*r:embed="rId\d+"\/>/)
    expect(again.media.filter((f) => f.endsWith('.svg'))).toHaveLength(1)
  })

  it('writes a plain blip when there is no SVG', async () => {
    const { documentXml, media } = await open(await saveImages([icon({ svgBase64: undefined })]))
    expect(documentXml).toMatch(/<a:blip r:embed="rId\d+"\/>/)
    expect(documentXml).not.toContain('svgBlip')
    expect(media.filter((f) => f.endsWith('.svg'))).toHaveLength(0)
  })

  it('lands the SVG of a picture placed in a table cell', async () => {
    const token = `${NEW_PICTURE_TOKEN_PREFIX}1`
    const drawing = inlinePictureDrawingXml({
      rId: token,
      svgRId: token + NEW_PICTURE_SVG_SUFFIX,
      docPrId: token,
      widthPx: 96,
      heightPx: 96,
    })
    const table = (inner: string) =>
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: table('<w:p/>') }))
    const { zip, documentXml, targetOf } = await open(
      await saveDocx(parsed, [{ kind: 'xml', xml: table(`<w:p><w:r>${drawing}</w:r></w:p>`) }], {
        newPictures: {
          [token]: { base64: TINY_PNG_BASE64, mime: 'image/png', svgBase64: SVG_BASE64 },
        },
      }),
    )
    expect(documentXml).not.toContain(NEW_PICTURE_TOKEN_PREFIX)
    const png = targetOf(/<a:blip r:embed="(rId\d+)"/.exec(documentXml)?.[1])
    const svg = targetOf(/<asvg:svgBlip [^>]*r:embed="(rId\d+)"/.exec(documentXml)?.[1])
    expect(png).toMatch(/\.png$/)
    expect(svg).toMatch(/\.svg$/)
    expect(await zip.file(`word/${svg}`)!.async('string')).toBe(SVG)
  })
})

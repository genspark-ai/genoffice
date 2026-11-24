/** Linked pictures of a header: read for Edit > Links, and what the dialog did written into the part. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type HeaderFooter } from '../src'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const OTHER_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const graphic = (id: number, link: string) =>
  `<wp:extent cx="914400" cy="457200"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:blipFill><a:blip r:embed="rId1" r:link="${link}"/></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic>'
/** a floating logo alone in its paragraph, then a picture in a line of text */
const HEADER =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="2">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="page"><wp:posOffset>228600</wp:posOffset></wp:positionV>' +
  `${graphic(11, 'rId2').replace('<wp:docPr', '<wp:wrapNone/><wp:docPr')}</wp:anchor></w:drawing></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Company </w:t></w:r><w:r><w:drawing><wp:inline>${graphic(12, 'rId3')}</wp:inline></w:drawing></w:r></w:p>`
const RELS =
  `<Relationship Id="rId1" Type="${IMAGE_REL}" Target="media/image1.png"/>` +
  `<Relationship Id="rId2" Type="${IMAGE_REL}" Target="file:///C:/art/logo.png" TargetMode="External"/>` +
  `<Relationship Id="rId3" Type="${IMAGE_REL}" Target="file:///C:/art/mark.png" TargetMode="External"/>`

const build = () =>
  buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    withImage: true,
    headerXml: HEADER,
    headerRelsXml: RELS,
  })

async function headerOf(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/header1.xml')!.async('string')
  const rels = await zip.file('word/_rels/header1.xml.rels')!.async('string')
  const rel = (id: string | undefined) =>
    id ? (new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*>`).exec(rels)?.[0] ?? '') : ''
  const blips = [...xml.matchAll(/<a:blip\b[^>]*>/g)].map((m) => ({
    embed: rel(/r:embed="([^"]+)"/.exec(m[0])?.[1]),
    link: rel(/r:link="([^"]+)"/.exec(m[0])?.[1]),
  }))
  return { zip, xml, rels, blips }
}

describe('linked pictures of a header', () => {
  it('are read with the ordinal of their drawing, in both shapes a header picture is kept in', async () => {
    const parsed = await parseDocx(await build())
    const part = Object.values(parsed.hfParts ?? {})[0]
    expect(part.linkedPictures).toEqual([
      { drawing: 0, target: 'file:///C:/art/logo.png', docPr: '11' },
      { drawing: 1, target: 'file:///C:/art/mark.png', docPr: '12' },
    ])
    // the logo is display data of the part, the other a run of its paragraph
    expect(part.images?.map((image) => image.linkDrawing)).toEqual([0])
    const run = part.paras.flatMap((p) => p.runs).find((r) => r.image)
    expect(run?.image?.linkTarget).toBe('file:///C:/art/mark.png')
    expect(run?.image?.dataUrl).toBe(`data:image/png;base64,${TINY_PNG_BASE64}`)
  })

  it('a drawing whose id another drawing shares is listed without it', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p/>',
      withImage: true,
      headerXml: HEADER.replace('id="12"', 'id="11"'),
      headerRelsXml: RELS,
    })
    const part = Object.values((await parseDocx(bytes)).hfParts ?? {})[0]
    expect(part.linkedPictures?.map((p) => p.docPr)).toEqual([undefined, undefined])
  })

  it('Break Link, Change Source and Update Now are written into the part; its text keeps its bytes', async () => {
    const bytes = await build()
    const parsed = await parseDocx(bytes)
    const before = await headerOf(bytes)
    const header: HeaderFooter = {
      text: parsed.headerText ?? '',
      paras: parsed.headerParas ?? undefined,
      pictureOnly: true,
      pictureLinks: {
        0: { target: null },
        1: {
          target: 'file:///D:/new/mark & co.png',
          bytes: { base64: OTHER_PNG, mime: 'image/png' },
        },
      },
    }
    const after = await headerOf(await saveDocx(parsed, [], { header }))
    // the logo keeps its picture and names no file; the other names the new file and holds its bytes
    expect(after.blips[0].embed).toContain('Target="media/image1.png"')
    expect(after.blips[0].link).toBe('')
    expect(after.blips[1].link).toContain('Target="file:///D:/new/mark &amp; co.png"')
    expect(after.blips[1].link).toContain('TargetMode="External"')
    const media = /Target="([^"]+)"/.exec(after.blips[1].embed)![1]
    expect(media).not.toBe('media/image1.png')
    expect(await after.zip.file(`word/${media}`)!.async('base64')).toBe(OTHER_PNG)
    expect(after.xml).not.toContain('aidocs-')
    // nothing else of the part moved
    const neutral = (xml: string) => xml.replace(/<a:blip\b[^>]*>/g, '<a:blip/>')
    expect(neutral(after.xml)).toBe(neutral(before.xml))

    const again = await parseDocx(await saveDocx(parsed, [], { header }))
    expect(Object.values(again.hfParts ?? {})[0].linkedPictures).toEqual([
      { drawing: 1, target: 'file:///D:/new/mark & co.png', docPr: '12' },
    ])
  })

  it('a link that only names another file keeps the embedded picture', async () => {
    const parsed = await parseDocx(await build())
    const header: HeaderFooter = {
      text: parsed.headerText ?? '',
      paras: parsed.headerParas ?? undefined,
      pictureOnly: true,
      pictureLinks: { 0: { target: 'file:///D:/new/logo.png' } },
    }
    const after = await headerOf(await saveDocx(parsed, [], { header }))
    expect(after.blips[0].embed).toContain('Target="media/image1.png"')
    expect(after.blips[0].link).toContain('Target="file:///D:/new/logo.png"')
    expect(after.blips[1].link).toContain('Target="file:///C:/art/mark.png"')
  })
})

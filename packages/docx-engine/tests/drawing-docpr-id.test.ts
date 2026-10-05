/**
 * wp:docPr/@id of a newly embedded picture: a new picture must never be handed an
 * id the document already uses. A duplicate drawing-object id makes Word flag the
 * file for repair, or silently drop one of the two drawings.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

/** a paragraph whose inline picture already owns wp:docPr/@id `docPrId` */
const drawingParagraph = (docPrId: string) =>
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  `<wp:docPr id=${docPrId} name="Original"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const TEXT_PARA = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

const docPrIds = (xml: string): string[] =>
  [...xml.matchAll(/<wp:docPr\s[^>]*?\bid=["'](\d+)["']/g)].map((m) => m[1])

/** save `bodyXml` with the originals kept and one new picture appended */
async function saveInsertingOneImage(bodyXml: string): Promise<string> {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  const saved = await saveDocx(parsed, [
    { kind: 'original', docxIndex: 0 },
    { kind: 'original', docxIndex: 1 },
    { kind: 'original', docxIndex: 2 },
    {
      kind: 'image',
      image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 20, heightPx: 20 },
    },
  ])
  return documentXmlOf(saved)
}

describe('wp:docPr ids of newly embedded pictures', () => {
  // The docPr counter was seeded from the media count only, so a document with no
  // GenOffice media always started it at the 9000 base: inserting one picture next
  // to an existing <wp:docPr id="9002"> emitted a second id 9002.
  it('never mints an id the document already uses', async () => {
    const xml = await saveInsertingOneImage(
      drawingParagraph('"9002"') + drawingParagraph('"9001"') + TEXT_PARA,
    )
    const ids = docPrIds(xml)
    expect(ids).toHaveLength(3)
    expect(new Set(ids).size).toBe(ids.length)
    // the pre-existing ids survive untouched, and the new one lands above them
    expect(ids).toContain('9002')
    expect(ids).toContain('9001')
    expect(Math.max(...ids.map(Number))).toBeGreaterThan(9002)
  })

  it('reads ids written with single quotes too', async () => {
    const xml = await saveInsertingOneImage(
      drawingParagraph("'9002'") + drawingParagraph("'1'") + TEXT_PARA,
    )
    const ids = docPrIds(xml)
    expect(new Set(ids).size).toBe(ids.length)
    expect(Math.max(...ids.map(Number))).toBeGreaterThan(9002)
  })

  it('keeps ids unique when several pictures are inserted into an empty body', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TEXT_PARA }))
    const image = (): SaveBlock => ({
      kind: 'image',
      image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 20, heightPx: 20 },
    })
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      image(),
      image(),
      image(),
    ])
    const ids = docPrIds(await documentXmlOf(saved))
    expect(ids).toHaveLength(3)
    expect(new Set(ids).size).toBe(3)
  })
})

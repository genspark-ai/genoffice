/**
 * A picture inserted into a table cell is saved as part of the cell's own
 * paragraph XML, so its drawing carries a placeholder token for the image
 * relationship and the drawing id until saveDocx lands the bytes.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  inlinePictureDrawingXml,
  NEW_PICTURE_TOKEN_PREFIX,
  parseDocx,
  saveDocx,
} from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const token = (n: number) => `${NEW_PICTURE_TOKEN_PREFIX}${n}`

const cell = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const table = (cells: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr>${cells.join('')}</w:tr></w:tbl>`

const pictureRun = (n: number) =>
  `<w:p><w:r>${inlinePictureDrawingXml({ rId: token(n), docPrId: token(n), widthPx: 40, heightPx: 20 })}</w:r></w:p>`

const ORIGINAL = table([cell('<w:p><w:r><w:t>a</w:t></w:r></w:p>'), cell('<w:p/>')])

async function saveTable(
  xml: string,
  pictures: Record<string, { base64: string; mime: 'image/png' }>,
) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: ORIGINAL }))
  const bytes = await saveDocx(parsed, [{ kind: 'xml', xml }], { newPictures: pictures })
  const zip = await JSZip.loadAsync(bytes)
  return {
    zip,
    documentXml: await zip.file('word/document.xml')!.async('string'),
    rels: await zip.file('word/_rels/document.xml.rels')!.async('string'),
  }
}

describe('pictures inserted into table cells', () => {
  it('lands the bytes as a media part and swaps the placeholder for the relationship id', async () => {
    const png = { base64: TINY_PNG_BASE64, mime: 'image/png' as const }
    const { zip, documentXml, rels } = await saveTable(
      table([cell(pictureRun(1)), cell('<w:p/>')]),
      { [token(1)]: png },
    )
    expect(documentXml).not.toContain(NEW_PICTURE_TOKEN_PREFIX)
    const rId = /<a:blip r:embed="(rId\d+)"/.exec(documentXml)?.[1]
    expect(rId).toBeTruthy()
    const target = new RegExp(`Id="${rId}"[^>]*Target="(media/[^"]+)"`).exec(rels)?.[1]
    expect(target).toMatch(/^media\/aidocs\d+\.png$/)
    expect(zip.file(`word/${target}`)).toBeTruthy()
    const ids = [...documentXml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1])
    expect(ids).toHaveLength(1)
    expect(documentXml).toContain(`<pic:cNvPr id="${ids[0]}"`)
  })

  it('mints one drawing id per picture while identical bytes share one media part', async () => {
    const png = { base64: TINY_PNG_BASE64, mime: 'image/png' as const }
    const { zip, documentXml } = await saveTable(
      table([cell(pictureRun(1)), cell(pictureRun(2))]),
      { [token(1)]: png, [token(2)]: png },
    )
    const ids = [...documentXml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1])
    expect(ids).toHaveLength(2)
    expect(new Set(ids).size).toBe(2)
    const media = Object.keys(zip.files).filter(
      (f) => f.startsWith('word/media/') && !zip.files[f].dir,
    )
    expect(media).toHaveLength(1)
    const rIds = [...documentXml.matchAll(/<a:blip r:embed="(rId\d+)"/g)].map((m) => m[1])
    expect(rIds).toHaveLength(2)
    expect(rIds[0]).toBe(rIds[1])
  })

  it('leaves a token alone when no bytes were handed over for it', async () => {
    const { documentXml } = await saveTable(table([cell(pictureRun(7)), cell('<w:p/>')]), {})
    expect(documentXml).toContain(token(7))
  })
})

/**
 * Paragraphs mixing real text with inline pictures (cover pages, icon bullets)
 * must stay editable text paragraphs with run-level images — the image-block
 * classification used to drop every text run in the paragraph.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

const INLINE_IMAGE_RUN =
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

const ANCHORED_IMAGE_RUN = INLINE_IMAGE_RUN.replace(/wp:inline/g, 'wp:anchor')

describe('text + inline image mixed paragraphs', () => {
  it('keeps the text runs and carries the picture as an image run', async () => {
    const bodyXml =
      `<w:p><w:r><w:t>COVER TITLE</w:t></w:r>${INLINE_IMAGE_RUN}` +
      '<w:r><w:t>Group 1</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    const runs = block.runs!
    expect(runs.map((r) => r.text).join('')).toBe('COVER TITLEGroup 1')
    const imageRun = runs.find((r) => r.image)
    expect(imageRun?.image?.dataUrl).toMatch(/^data:image\/png;base64,/)
    expect(imageRun?.image?.widthPx).toBe(96)
  })

  it('a picture-only paragraph is still an image block', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true }))
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('an anchored picture sharing the paragraph with text keeps the text', async () => {
    const bodyXml = `<w:p>${ANCHORED_IMAGE_RUN}<w:r><w:t>Lorem ipsum body text</w:t></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.runs!.map((r) => r.text).join('')).toBe('Lorem ipsum body text')
    const imageRun = block.runs!.find((r) => r.image)
    expect(imageRun?.image?.dataUrl).toMatch(/^data:image\/png;base64,/)
    // no wrap element on the anchor = in front of text
    expect(imageRun?.image?.wrap).toBe('front')
  })

  it('carries the anchor posOffset on the run image', async () => {
    const anchored = ANCHORED_IMAGE_RUN.replace(
      '<wp:extent',
      '<wp:positionH relativeFrom="column"><wp:posOffset>1882747</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-104609</wp:posOffset></wp:positionV>' +
        '<wp:extent',
    )
    const bodyXml = `<w:p>${anchored}<w:r><w:t>wrapped text</w:t></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const imageRun = doc.blocks[0].runs!.find((r) => r.image)
    expect(imageRun?.image?.offsetXEmu).toBe(1882747)
    expect(imageRun?.image?.offsetYEmu).toBe(-104609)
  })

  it('an anchored picture without text is still an image block', async () => {
    const bodyXml = `<w:p>${ANCHORED_IMAGE_RUN}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })
})

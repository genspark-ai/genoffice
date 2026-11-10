import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const LINKED_IMAGE_URL = 'https://example.com/cfimages?u1=abc&width=2560'

const LINKED_IMAGE_PARAGRAPH =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:link="rId30"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const IMAGE_REL = `<Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${LINKED_IMAGE_URL.replace(/&/g, '&amp;')}" TargetMode="External"/>`

/** INCLUDEPICTURE field whose cached result is the linked drawing (as produced by some converters) */
const INCLUDEPICTURE_PARAGRAPH =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE \\d "${LINKED_IMAGE_URL.replace(/&/g, '&amp;')}" \\y </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:link="rId30"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

describe('linked (external) pictures', () => {
  it('exposes the external URL as imageDataUrl instead of degrading to Drawing object', async () => {
    const bytes = await buildDocx({ bodyXml: LINKED_IMAGE_PARAGRAPH, extraRels: IMAGE_REL })
    const doc = await parseDocx(bytes)
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageDataUrl).toBe(LINKED_IMAGE_URL)
  })

  it('shows the picture for INCLUDEPICTURE field paragraphs instead of a field chip', async () => {
    const bytes = await buildDocx({ bodyXml: INCLUDEPICTURE_PARAGRAPH, extraRels: IMAGE_REL })
    const doc = await parseDocx(bytes)
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageDataUrl).toBe(LINKED_IMAGE_URL)
    // still protected: saving must keep the original field XML byte-identical
    expect(block.originalXml).toContain('INCLUDEPICTURE')
  })

  it('keeps the text visible when INCLUDEPICTURE shares the paragraph with text', async () => {
    const bodyXml = INCLUDEPICTURE_PARAGRAPH.replace(
      '<w:p>',
      '<w:p><w:r><w:t>Company logo: </w:t></w:r>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml, extraRels: IMAGE_REL }))
    const block = doc.blocks[0]
    // the field passthrough keeps the text instead of an image block dropping it
    expect(block.type).toBe('passthrough')
    expect(block.previewText).toContain('Company logo:')
    expect(block.originalXml).toContain('INCLUDEPICTURE')
  })
})

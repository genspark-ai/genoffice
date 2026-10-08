import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/parse'
import { IMAGE_PARAGRAPH_XML, buildDocx } from './helpers/build-docx'

const withPPr = (pPr: string) => IMAGE_PARAGRAPH_XML.replace('<w:p>', `<w:p>${pPr}`)

describe('w:pageBreakBefore on drawing-only / textbox paragraphs', () => {
  it('survives onto the protected image block', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Title</w:t></w:r></w:p>' + withPPr('<w:pPr><w:pageBreakBefore/></w:pPr>'),
        withImage: true,
      }),
    )
    expect(doc.blocks[1].type).toBe('image')
    expect(doc.blocks[1].format?.pageBreakBefore).toBe(true)
  })

  it('ignores an explicit off value', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: withPPr('<w:pPr><w:pageBreakBefore w:val="0"/></w:pPr>'),
        withImage: true,
      }),
    )
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].format?.pageBreakBefore).toBeUndefined()
  })

  it('survives onto a textbox passthrough block', async () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'paragraph',
      xEmu: 0,
      yEmu: 0,
      widthEmu: 914400,
      heightEmu: 914400,
      id: 5,
      paragraphs: [{ runs: [{ text: 'Box' }] }],
      holderPageBreakBefore: true,
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].label).toBe('Text box')
    expect(doc.blocks[0].format?.pageBreakBefore).toBe(true)
  })

  it('does not take a pageBreakBefore from a paragraph nested in the textbox', async () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'paragraph',
      xEmu: 0,
      yEmu: 0,
      widthEmu: 914400,
      heightEmu: 914400,
      id: 5,
      paragraphs: [{ runs: [{ text: 'Box' }], format: { pageBreakBefore: true } }],
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].format?.pageBreakBefore).toBeUndefined()
  })
  it('ignores a pageBreakBefore that only lives in the tracked-change previous pPr', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: withPPr(
          '<w:pPr><w:pPrChange w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:pPr><w:pageBreakBefore/></w:pPr></w:pPrChange></w:pPr>',
        ),
        withImage: true,
      }),
    )
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].format?.pageBreakBefore).toBeUndefined()
  })

  it('keeps a live pageBreakBefore next to a pPrChange that lacked it', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Title</w:t></w:r></w:p>' +
          withPPr(
            '<w:pPr><w:pageBreakBefore/><w:pPrChange w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr>',
          ),
        withImage: true,
      }),
    )
    expect(doc.blocks[1].format?.pageBreakBefore).toBe(true)
  })
})

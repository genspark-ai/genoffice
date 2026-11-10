import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/parse'
import { IMAGE_PARAGRAPH_XML, buildDocx } from './helpers/build-docx'

const withPPr = (pPr: string) => IMAGE_PARAGRAPH_XML.replace('<w:p>', `<w:p>${pPr}`)

const TOC_OPEN =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Chapter 1</w:t></w:r></w:p>'
const tocClosingHeading = (runs: string) =>
  `<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r>${runs}</w:p>`

describe('page break ahead of the visible text of a field paragraph', () => {
  it('turns the page before the heading that closes the TOC field', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          TOC_OPEN +
          tocClosingHeading('<w:r><w:br w:type="page"/></w:r><w:r><w:t>PREAMBLE</w:t></w:r>'),
      }),
    )
    const heading = doc.blocks[1]
    expect(heading.type).toBe('passthrough')
    expect(heading.fieldDisplay).toMatchObject({ kind: 'text', left: 'PREAMBLE' })
    expect(heading.format?.pageBreakBefore).toBe(true)
  })

  it('leaves a break that follows visible text alone', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          TOC_OPEN +
          tocClosingHeading(
            '<w:r><w:t>Intro</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>PREAMBLE</w:t></w:r>',
          ),
      }),
    )
    expect(doc.blocks[1].type).toBe('passthrough')
    expect(doc.blocks[1].format?.pageBreakBefore).toBeUndefined()
  })
})

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

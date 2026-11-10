import { describe, expect, it } from 'vitest'
import { parseDocx, type TextboxDisplay } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

/**
 * The non-editable surfaces (field / formula / picture wrappers, stray anchor
 * runs) read their geometry from the property resolver. Direct values travel
 * inline; what the style chain supplies stays with the wrapper's CSS cascade.
 */

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const stylesOf = (styles: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="200"/><w:ind w:left="300"/></w:pPr></w:style>' +
  `${styles}</w:styles>`

const field = (pPr: string, instr: string, result: string) =>
  `<w:p>${pPr}<w:r><w:t>See </w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `${result}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`

describe('field paragraphs (parse-fields)', () => {
  it('carries the direct first-line and right indent of a CITATION paragraph, not the style left', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: field(
          '<w:pPr><w:ind w:right="240" w:firstLine="420"/><w:spacing w:before="60"/></w:pPr>',
          'CITATION Smi24 \\l 1033',
          '<w:r><w:t>(Smith 2024)</w:t></w:r>',
        ),
        stylesXml: stylesOf(''),
      }),
    )
    const fd = doc.blocks[0].fieldDisplay
    expect(fd).toMatchObject({
      kind: 'text',
      indentRightTwips: 240,
      indentFirstLineTwips: 420,
      spaceBeforeTwips: 60,
    })
    // Normal's left indent and after-spacing reach the wrapper through CSS
    expect(fd?.indentLeftTwips).toBeUndefined()
    expect(fd?.spaceAfterTwips).toBeUndefined()
    expect(fd?.additiveSpacing).toBeUndefined()
  })

  it('reads a direct explicit zero and w:start/w:end spellings', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: field(
          '<w:pPr><w:ind w:start="0" w:end="120" w:hanging="200"/></w:pPr>',
          'CITATION Doe21 \\l 1033',
          '<w:r><w:t>Table 1</w:t></w:r>',
        ),
        stylesXml: stylesOf(''),
      }),
    )
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      indentLeftTwips: 0,
      indentRightTwips: 120,
      indentFirstLineTwips: -200,
    })
  })

  it('flags the compat spacing mode on the display model', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: field(
          '<w:pPr><w:spacing w:before="60"/></w:pPr>',
          'CITATION Doe21 \\l 1033',
          '<w:r><w:t>(Doe 2021)</w:t></w:r>',
        ),
        stylesXml: stylesOf(''),
        extraParts: [
          {
            path: 'word/settings.xml',
            xml: `<w:settings ${NS}><w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat></w:settings>`,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
          },
        ],
      }),
    )
    expect(doc.blocks[0].fieldDisplay?.additiveSpacing).toBe(true)
  })

  it('sizes an rtl TOC result run by its w:szCs twin', async () => {
    const toc = (rPr: string) =>
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:ind w:left="440"/>' +
      '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      `<w:r>${rPr}<w:t>פרק א</w:t></w:r><w:r><w:tab/></w:r>` +
      '<w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const stylesXml = stylesOf(
      '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>',
    )
    const rtl = await parseDocx(
      await buildDocx({
        bodyXml: toc('<w:rPr><w:rtl/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr>'),
        stylesXml,
      }),
    )
    const fd = rtl.blocks[0].fieldDisplay
    expect(fd).toMatchObject({ kind: 'tocLine', szHalfPoints: 28, bold: true })
    // a styled TOC level indents through the level ladder, not its own w:ind
    expect(fd?.indentLeftTwips).toBeUndefined()
    const ltr = await parseDocx(
      await buildDocx({
        bodyXml: toc('<w:rPr><w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr>'),
        stylesXml,
      }),
    )
    expect(ltr.blocks[0].fieldDisplay?.szHalfPoints).toBe(22)
    expect(ltr.blocks[0].fieldDisplay?.bold).toBeUndefined()
  })
})

describe('picture paragraphs (imageMeta)', () => {
  const picture = (pPr: string) =>
    IMAGE_PARAGRAPH_XML.replace('<w:p>', `<w:p>${pPr}</w:p>`.slice(0, -6))

  it('takes style-only spacing from the chain, docDefaults for a style off the Normal chain', async () => {
    const stylesXml = stylesOf(
      '<w:style w:type="paragraph" w:styleId="Pic"><w:name w:val="Pic"/>' +
        '<w:pPr><w:spacing w:before="120"/></w:pPr></w:style>',
    )
    const styled = await parseDocx(
      await buildDocx({
        bodyXml: picture('<w:pPr><w:pStyle w:val="Pic"/></w:pPr>'),
        withImage: true,
        stylesXml,
      }),
    )
    expect(styled.blocks[0].type).toBe('image')
    expect(styled.blocks[0].imageParagraphSpaceBefore).toBe(120)
    // Pic is not based on Normal: its silent after-spacing is docDefaults' 160, not Normal's 200
    expect(styled.blocks[0].imageParagraphSpaceAfter).toBe(160)
    // the style's indent stays off the model: direct w:ind only
    expect(styled.blocks[0].imageParagraphIndentLeft).toBeUndefined()
    const unstyled = await parseDocx(
      await buildDocx({ bodyXml: picture(''), withImage: true, stylesXml }),
    )
    expect(unstyled.blocks[0].imageParagraphSpaceAfter).toBe(200)
    expect(unstyled.blocks[0].imageParagraphSpaceBefore).toBe(0)
  })

  it('carries HTML auto spacing as a flag', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: picture(
          '<w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1" w:before="100"/></w:pPr>',
        ),
        withImage: true,
        stylesXml: stylesOf(''),
      }),
    )
    expect(doc.blocks[0]).toMatchObject({
      imageParagraphSpaceBefore: 100,
      imageParagraphSpaceBeforeAuto: true,
      imageParagraphSpaceAfterAuto: true,
    })
  })

  it('sizes the picture line by the drawing run resolved through its chain', async () => {
    const stylesXml = stylesOf(
      '<w:style w:type="paragraph" w:styleId="Pic"><w:name w:val="Pic"/>' +
        '<w:rPr><w:rFonts w:ascii="Book Antiqua"/><w:sz w:val="20"/></w:rPr></w:style>',
    )
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: picture(
          '<w:pPr><w:pStyle w:val="Pic"/><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>',
        ),
        withImage: true,
        stylesXml,
      }),
    )
    expect(doc.blocks[0].imageMarkFont).toBe('Book Antiqua')
    expect(doc.blocks[0].imageMarkSizeHalfPoints).toBe(20)
  })
})

const TXBX = (paras: string) =>
  '<w:p><w:pPr><w:pStyle w:val="Side"/><w:ind w:left="0"/></w:pPr><w:r><w:t>Anchor text</w:t></w:r>' +
  '<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="2000000" cy="1000000"/><wp:wrapSquare wrapText="bothSides"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${paras}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'

describe('textbox anchor and content paragraphs', () => {
  it('reads the host paragraph only: a box paragraph pPr is not the anchor geometry', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: TXBX(
          '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="240"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Box</w:t></w:r></w:p>',
        ).replace(/^<w:p><w:pPr>[\s\S]*?<\/w:pPr>/, '<w:p>'),
        stylesXml: stylesOf(''),
      }),
    )
    const block = doc.blocks[0]
    expect(block.textboxes).toHaveLength(1)
    expect(block.imageAlign).toBeUndefined()
    expect(block.imageParagraphLineTwips).toBeUndefined()
    expect(block.imageParagraphSpaceAfter).toBeUndefined()
  })

  const stylesXml = stylesOf(
    '<w:style w:type="paragraph" w:styleId="Side"><w:name w:val="Side"/>' +
      '<w:pPr><w:jc w:val="right"/><w:ind w:left="720"/></w:pPr></w:style>',
  )

  it('a direct w:left=0 on the anchor paragraph cancels the style indent of the stray runs', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: TXBX(
          '<w:p><w:pPr><w:pStyle w:val="Side"/><w:ind w:left="0"/></w:pPr><w:r><w:t>Box</w:t></w:r></w:p>',
        ),
        stylesXml,
      }),
    )
    const block = doc.blocks[0]
    expect(block.strayRuns?.map((r) => r.text)).toEqual(['Anchor text'])
    expect(block.strayIndent).toEqual({ leftTwips: 0 })
    expect(block.strayAlign).toBe('right')
    const box = (block.textboxes as TextboxDisplay[])[0]
    expect(box.paras[0]).toMatchObject({ styleId: 'Side', indentLeft: 0 })
  })
})

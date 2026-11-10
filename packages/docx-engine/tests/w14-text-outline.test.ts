/**
 * Character formatting that reaches the screen only through styles.xml or the
 * w14 extension namespace: w14:textOutline strokes (run and character style),
 * character-style w:shd, and w:tblStylePr firstRow caps / theme font.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W14_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"'

const OUTLINE =
  '<w14:textOutline w14:w="9525" w14:cap="rnd" w14:cmpd="sng" w14:algn="ctr">' +
  '<w14:solidFill><w14:srgbClr w14:val="FF0000"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:solidFill>' +
  '<w14:prstDash w14:val="solid"/><w14:bevel/></w14:textOutline>'

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${W14_NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="IntenseEmphasis"><w:name w:val="Intense Emphasis"/>' +
  `<w:rPr>${OUTLINE}</w:rPr></w:style>` +
  '<w:style w:type="character" w:styleId="ChangeType"><w:name w:val="ChangeType"/>' +
  '<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="000000"/></w:rPr></w:style>' +
  '</w:styles>'

describe('w14:textOutline', () => {
  it('reads a run-level solid outline with width and alpha', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:r><w:rPr>${OUTLINE}</w:rPr><w:t>Hello</w:t></w:r></w:p>`,
        docRootExtraAttrs: W14_NS,
        stylesXml: STYLES_XML,
      }),
    )
    expect(doc.blocks[0].runs?.[0].textOutline).toEqual({
      color: 'FF0000',
      widthPt: 0.75,
      alpha: 0.4,
    })
  })

  it('carries the outline of a character style declared in an mc:Ignorable styles part', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:rStyle w:val="IntenseEmphasis"/></w:rPr><w:t>Hello</w:t></w:r></w:p>',
        docRootExtraAttrs: W14_NS,
        stylesXml: STYLES_XML,
      }),
    )
    expect(doc.styles.get('IntenseEmphasis')?.display?.textOutline).toEqual({
      color: 'FF0000',
      widthPt: 0.75,
      alpha: 0.4,
    })
    expect(doc.blocks[0].runs?.[0].styleId).toBe('IntenseEmphasis')
  })

  it('ignores noFill outlines', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w14:textOutline w14:w="9525"><w14:noFill/></w14:textOutline></w:rPr>' +
          '<w:t>x</w:t></w:r></w:p>',
        docRootExtraAttrs: W14_NS,
      }),
    )
    expect(doc.blocks[0].runs?.[0].textOutline).toBeUndefined()
  })
})

describe('linked styles', () => {
  it('copies outline and shading onto a character shell without its own rPr', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Tag"><w:name w:val="Tag"/><w:link w:val="TagChar"/>' +
      `<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="000000"/>${OUTLINE}</w:rPr></w:style>` +
      '<w:style w:type="character" w:styleId="TagChar"><w:name w:val="Tag Char"/><w:link w:val="Tag"/></w:style>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        docRootExtraAttrs: W14_NS,
        stylesXml: STYLES_XML.replace('</w:styles>', `${styles}</w:styles>`),
      }),
    )
    const shell = doc.styles.get('TagChar')?.display
    expect(shell?.shading).toBe('000000')
    expect(shell?.textOutline?.color).toBe('FF0000')
  })
})

describe('character style w:shd', () => {
  it('exposes the style shading fill for the renderer', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:rStyle w:val="ChangeType"/></w:rPr><w:t>Question:</w:t></w:r></w:p>',
        stylesXml: STYLES_XML,
      }),
    )
    expect(doc.styles.get('ChangeType')?.display?.shading).toBe('000000')
    expect(doc.styles.get('ChangeType')?.display?.color).toBeUndefined()
  })
})

describe('table style firstRow run formatting', () => {
  it('keeps caps, italic, theme font and letter spacing of the header condition', async () => {
    const style =
      '<w:style w:type="table" w:styleId="Calendar2"><w:name w:val="Calendar 2"/>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr>' +
      '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>' +
      '<w:b w:val="0"/><w:i/><w:caps/><w:smallCaps w:val="0"/>' +
      '<w:color w:val="4F81BD"/><w:spacing w:val="20"/><w:sz w:val="32"/>' +
      '</w:rPr></w:tblStylePr></w:style>'
    const table =
      '<w:tbl><w:tblPr><w:tblStyle w:val="Calendar2"/><w:tblLook w:firstRow="1"/></w:tblPr>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>May</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>M</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: table, extraStylesXml: style }))
    const firstRow = doc.styles.get('Calendar2')?.tableDisplay?.firstRow
    expect(firstRow).toMatchObject({
      caps: 'all',
      italic: true,
      color: '4F81BD',
      sizeHalfPoints: 32,
      charSpacingTwips: 20,
    })
    expect(firstRow?.bold).toBeUndefined()
    expect(doc.blocks[0].table?.repeatHeaderRows).toEqual([true, true, false])
  })
})

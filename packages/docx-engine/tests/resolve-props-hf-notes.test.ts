/**
 * Header/footer paragraphs, layout-table cell paragraphs and footnote
 * paragraphs read the property resolver: one style chain (docDefaults ->
 * table style -> paragraph style -> direct) instead of three hand-rolled
 * lookups. settings.xml doNotUseHTMLParagraphAutoSpacing turns Word's 14pt
 * HTML auto spacing into 0 (Word probe 2026-10-22, prod corpus 20261021
 * sample 017).
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const FOOTER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const FOOTNOTES_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const styles = (body: string) =>
  `${XML_DECL}<w:styles ${W}><w:docDefaults>` +
  '<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
  '</w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  `${body}</w:styles>`

const settingsPart = (flags: string) => ({
  path: 'word/settings.xml',
  xml: `${XML_DECL}<w:settings ${W}><w:compat>${flags}</w:compat></w:settings>`,
  contentType: SETTINGS_TYPE,
})

async function hfDoc(opts: {
  stylesBody?: string
  header?: string
  footer?: string
  settings?: string
}) {
  const parts = []
  let rels = ''
  let sect = ''
  if (opts.header) {
    parts.push({
      path: 'word/header1.xml',
      xml: `${XML_DECL}<w:hdr ${W}>${opts.header}</w:hdr>`,
      contentType: HEADER_TYPE,
    })
    rels += `<Relationship Id="rId60" Type="${REL}/header" Target="header1.xml"/>`
    sect += '<w:headerReference w:type="default" r:id="rId60"/>'
  }
  if (opts.footer) {
    parts.push({
      path: 'word/footer1.xml',
      xml: `${XML_DECL}<w:ftr ${W}>${opts.footer}</w:ftr>`,
      contentType: FOOTER_TYPE,
    })
    rels += `<Relationship Id="rId61" Type="${REL}/footer" Target="footer1.xml"/>`
    sect += '<w:footerReference w:type="default" r:id="rId61"/>'
  }
  if (opts.settings !== undefined) parts.push(settingsPart(opts.settings))
  return parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      stylesXml: styles(opts.stylesBody ?? ''),
      extraRels: rels,
      extraParts: parts,
      sectPrExtra: sect,
    }),
  )
}

describe('header/footer paragraphs through the resolver', () => {
  it('style-only spacing lands on the paragraph; the docDefaults line stays with the strip base', async () => {
    const doc = await hfDoc({
      stylesBody:
        '<w:style w:type="paragraph" w:styleId="Banner"><w:name w:val="Banner"/>' +
        '<w:pPr><w:spacing w:before="120"/><w:jc w:val="both"/><w:ind w:left="360"/></w:pPr></w:style>',
      header:
        '<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr><w:r><w:t>title</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>plain</w:t></w:r></w:p>',
    })
    const [banner, plain] = doc.headerParas!
    expect(banner).toMatchObject({
      align: 'justify',
      spaceBefore: 120,
      spaceAfter: 200,
      indentLeft: 360,
    })
    expect(banner.lineRule).toBeUndefined()
    expect(banner.lineRawTwips).toBeUndefined()
    expect(plain.spaceAfter).toBe(200)
    expect(plain.spaceBefore).toBeUndefined()
    expect(plain.lineRule).toBeUndefined()
  })

  it('footer auto spacing is 14pt, and 0 under doNotUseHTMLParagraphAutoSpacing', async () => {
    const footer =
      '<w:p><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/></w:pPr>' +
      '<w:r><w:t>Page</w:t></w:r></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
      '<w:p><w:pPr><w:spacing w:before="100" w:beforeAutospacing="1"/></w:pPr><w:r><w:t>cell</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl>'
    const html = await hfDoc({ footer })
    expect(html.footerParas![0]).toMatchObject({ spaceBefore: 280, spaceAfter: 280 })
    expect(html.footerParas![1].cells![0].paraProps![0]?.spaceBefore).toBe(280)
    expect(html.footerParas![1].row?.additiveSpacing).toBeUndefined()

    const additive = await hfDoc({ footer, settings: '<w:doNotUseHTMLParagraphAutoSpacing/>' })
    expect(additive.footerParas![0].spaceBefore).toBeUndefined()
    expect(additive.footerParas![0].spaceAfter).toBeUndefined()
    expect(additive.footerParas![1].cells![0].paraProps![0]?.spaceBefore).toBeUndefined()
    expect(additive.footerParas![1].row?.additiveSpacing).toBe(true)
  })
})

describe('layout-table cell paragraphs through the resolver', () => {
  it('a direct w:left="0" cancels the style indent; the table style pPr sits under the paragraph style', async () => {
    const doc = await hfDoc({
      stylesBody:
        '<w:style w:type="paragraph" w:styleId="Indented"><w:name w:val="Indented"/>' +
        '<w:pPr><w:ind w:left="720"/><w:spacing w:after="60"/></w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:style>' +
        '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>' +
        '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>',
      header:
        '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
        '<w:tc><w:p><w:pPr><w:pStyle w:val="Indented"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p><w:pPr><w:pStyle w:val="Indented"/><w:ind w:left="0"/></w:pPr><w:r><w:t>reset</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc>' +
        '</w:tr></w:tbl>',
    })
    const [styled, reset, plain] = doc.headerParas![0].cells!.map((c) => c.paraProps![0])
    expect(styled).toMatchObject({ indentLeft: 720, spaceAfter: 60, sizeHalfPoints: 16 })
    expect(reset?.indentLeft).toBeUndefined()
    expect(reset?.spaceAfter).toBe(60)
    // Grid's pPr: after 0 replaces docDefaults' 200, line 240 travels inline
    expect(plain?.spaceAfter).toBeUndefined()
    expect(plain).toMatchObject({ lineRule: 'auto', lineRawTwips: 240 })
    expect(plain?.sizeHalfPoints).toBeUndefined()
  })
})

describe('footnote paragraphs through the resolver', () => {
  it('runs and spacing inherit the FootnoteText chain, unstyled notes the docDefaults', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>a</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
          '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>',
        stylesXml: styles(
          '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>' +
            '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
            '<w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/><w:sz w:val="20"/></w:rPr></w:style>',
        ),
        extraRels: `<Relationship Id="rId40" Type="${REL}/footnotes" Target="footnotes.xml"/>`,
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml:
              `${XML_DECL}<w:footnotes ${W}>` +
              '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
              '<w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> styled</w:t></w:r>' +
              '<w:r><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:t>small</w:t></w:r></w:p></w:footnote>' +
              '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p></w:footnote>' +
              '</w:footnotes>',
            contentType: FOOTNOTES_TYPE,
          },
        ],
      }),
    )
    const [styled, plain] = doc.footnotes
    expect(styled.styleId).toBe('FootnoteText')
    expect(styled.richParas?.[0]).toEqual([
      { text: 'styled', sizeHalfPoints: 20, fontAscii: 'Garamond' },
      { text: 'small', bold: true, sizeHalfPoints: 16, fontAscii: 'Garamond' },
    ])
    expect(styled.spacing).toEqual({
      beforeTwips: 0,
      afterTwips: 0,
      lineRule: 'auto',
      lineRawTwips: 240,
    })
    // the Normal template's Times New Roman stands in for silent docDefaults fonts
    expect(plain.richParas?.[0]).toEqual([
      { text: 'plain', sizeHalfPoints: 24, fontAscii: 'Times New Roman' },
    ])
    expect(plain.spacing).toEqual({
      beforeTwips: 0,
      afterTwips: 200,
      lineRule: 'auto',
      lineRawTwips: 276,
    })
  })
})

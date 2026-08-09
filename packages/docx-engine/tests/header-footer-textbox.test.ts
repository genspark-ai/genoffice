import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Chinese government documents (repro) place the footer page number
 * inside a VML textbox: w:p > w:r > w:pict > v:shape > v:textbox > w:txbxContent,
 * with a complex PAGE field ("— PAGE —") whose cached result is a literal digit.
 * The parser must surface the textbox paragraphs instead of dropping the footer.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const FTR_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const RPR =
  '<w:rPr><w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:cs="SimSun"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'

/** structure copied from the repro document's word/footer1.xml (fonts renamed, ids shortened) */
const GOV_TEXTBOX_FOOTER =
  XML_DECL +
  `<w:ftr ${FTR_NS}><w:p><w:pPr><w:pStyle w:val="6"/></w:pPr><w:r><w:pict>` +
  '<v:shape id="Text Box 8" o:spid="_x0000_s4096" o:spt="202" type="#_x0000_t202" ' +
  'style="position:absolute;left:0pt;margin-top:0pt;height:144pt;width:144pt;mso-position-horizontal:outside;mso-position-horizontal-relative:margin;mso-wrap-style:none;z-index:251659264;" ' +
  'filled="f" stroked="f" coordsize="21600,21600">' +
  '<v:path/><v:fill on="f" focussize="0,0"/><v:stroke on="f"/><v:imagedata o:title=""/>' +
  '<o:lock v:ext="edit" aspectratio="f"/>' +
  '<v:textbox inset="0mm,0mm,0mm,0mm" style="mso-fit-shape-to-text:t;"><w:txbxContent>' +
  '<w:p><w:pPr><w:pStyle w:val="6"/><w:ind w:left="315" w:leftChars="150" w:right="315" w:rightChars="150"/>' +
  `${RPR}</w:pPr>` +
  `<w:r>${RPR}<w:t xml:space="preserve">\u2014 </w:t></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r>${RPR}<w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${RPR}<w:t>1</w:t></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>` +
  `<w:r>${RPR}<w:t xml:space="preserve"> \u2014</w:t></w:r>` +
  '</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:ftr>'

async function parseWithFooter(footerXml: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId70" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      {
        path: 'word/footer1.xml',
        xml: footerXml,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      },
    ],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId70"/>',
  })
  return parseDocx(bytes)
}

describe('footer content inside a VML textbox', () => {
  it('parses the textbox paragraph into non-empty footer content with the PAGE marker', async () => {
    const parsed = await parseWithFooter(GOV_TEXTBOX_FOOTER)
    expect(parsed.footerHasPageNumber).toBe(true)
    // the "\u2014 N \u2014" literal runs survive; the PAGE field becomes the PAGE_MARK marker
    expect(parsed.footerText).toBe(`\u2014 ${PAGE_MARK} \u2014`)
    expect(parsed.footerParas).not.toBeNull()
    expect(parsed.footerParas!.length).toBeGreaterThan(0)
    const line = parsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')
    expect(line).toBe(`\u2014 ${PAGE_MARK} \u2014`)
    // run formatting from inside the textbox is preserved
    const firstRun = parsed.footerParas![0].runs[0]
    expect(firstRun.font).toBe('SimSun')
    expect(firstRun.sizeHalfPoints).toBe(28)
  })

  it('hfParts (multi-section lookup path) also carries the textbox paragraphs', async () => {
    const parsed = await parseWithFooter(GOV_TEXTBOX_FOOTER)
    const part = parsed.hfParts?.rId70
    expect(part).toBeDefined()
    expect(part!.hasPageNumber).toBe(true)
    expect(part!.paras.length).toBeGreaterThan(0)
    expect(part!.paras.flatMap((p) => p.runs.map((r) => r.text)).join('')).toBe(
      `\u2014 ${PAGE_MARK} \u2014`,
    )
  })

  it('footer PAGE field with a cached result drops the stale number and keeps literal runs', async () => {
    // same field structure but as a plain paragraph (no textbox): cached "7" must
    // not leak into the display text; the marker takes its place
    const plainFooter =
      XML_DECL +
      `<w:ftr ${FTR_NS}><w:p>` +
      '<w:r><w:t xml:space="preserve">\u2014 </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>7</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:t xml:space="preserve"> \u2014</w:t></w:r>' +
      '</w:p></w:ftr>'
    const parsed = await parseWithFooter(plainFooter)
    expect(parsed.footerHasPageNumber).toBe(true)
    expect(parsed.footerText).toBe(`\u2014 ${PAGE_MARK} \u2014`)
    expect(parsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')).toBe(
      `\u2014 ${PAGE_MARK} \u2014`,
    )
  })

  it('an empty decorative textbox still contributes no footer paragraphs', async () => {
    const emptyTextboxFooter =
      XML_DECL +
      `<w:ftr ${FTR_NS}><w:p><w:r><w:pict>` +
      '<v:shape id="s1" o:spt="202" type="#_x0000_t202" style="position:absolute;" filled="t" stroked="f">' +
      '<v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox>' +
      '</v:shape></w:pict></w:r></w:p></w:ftr>'
    const parsed = await parseWithFooter(emptyTextboxFooter)
    expect(parsed.footerParas ?? []).toHaveLength(0)
  })
})
